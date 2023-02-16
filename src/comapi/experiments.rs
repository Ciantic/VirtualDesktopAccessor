use crate::hresult::HRESULT;

use super::raw::*;

use super::desktop::*;
use super::interfaces::*;
use super::*;

use std::borrow::Borrow;
use std::cell::RefCell;
use std::ffi::c_void;
use std::pin::Pin;
use std::rc::Rc;
use std::{
    sync::{Arc, Condvar, Mutex},
    time::Duration,
};
use windows::core::Vtable;
use windows::Win32::UI::Shell::Common::IObjectArray;
use windows::{
    core::{GUID, HSTRING},
    Win32::{
        Foundation::HWND,
        System::{
            Com::{CoInitializeEx, COINIT_APARTMENTTHREADED},
            Threading::{
                CreateThread, GetCurrentThreadId, WaitForSingleObject, THREAD_CREATION_FLAGS,
            },
        },
        UI::WindowsAndMessaging::{
            DispatchMessageW, GetMessageW, PostQuitMessage, TranslateMessage, MSG, WM_USER,
        },
    },
};

unsafe extern "system" fn handler(_arg: *mut c_void) -> u32 {
    let mut msg = MSG::default();
    unsafe {
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    0
}

fn create_thread() {
    let thread_id = None;
    let handle = unsafe {
        CreateThread(
            None,
            0,
            Some(handler),
            None,
            THREAD_CREATION_FLAGS::default(),
            thread_id,
        )
    }
    .unwrap();

    // Join the thread
    unsafe { WaitForSingleObject(handle, u32::MAX) };
}

fn debug_desktop(desktop_new: &IVirtualDesktop, prefix: &str) {
    let mut gid = GUID::default();
    unsafe { desktop_new.get_id(&mut gid).panic_if_failed() };

    let mut name = HSTRING::new();
    unsafe { desktop_new.get_name(&mut name).panic_if_failed() };

    println!(
        "{}: {:?} {:?} {:?}",
        prefix,
        gid,
        name.to_string(),
        std::thread::current().id()
    );
}

struct SimpleVirtualDesktopNotificationWrapper {
    cookie: u32,
    ptr: IVirtualDesktopNotification,
    number_times_desktop_changed: Rc<RefCell<u32>>,
}

impl SimpleVirtualDesktopNotificationWrapper {
    pub fn new() -> Result<Pin<Box<SimpleVirtualDesktopNotificationWrapper>>> {
        println!(
            "Notification service created in thread {:?}",
            std::thread::current().id()
        );
        let provider = get_iservice_provider()?;
        let service = get_ivirtual_desktop_notification_service(&provider)?;
        let number_times_desktop_changed = Rc::new(RefCell::new(0));

        let ptr = SimpleVirtualDesktopNotification {
            number_times_desktop_changed: number_times_desktop_changed.clone(),
        };
        let mut notification = Pin::new(Box::new(SimpleVirtualDesktopNotificationWrapper {
            cookie: 0,
            ptr: ptr.into(),
            number_times_desktop_changed,
        }));

        let mut cookie = 0;
        unsafe {
            service
                .register(notification.ptr.as_raw(), &mut cookie)
                .panic_if_failed();
            assert_ne!(cookie, 0);
        }
        notification.cookie = cookie;
        println!(
            "Registered notification {} {:?}",
            cookie,
            std::thread::current().id()
        );

        Ok(notification)
    }
}

impl Drop for SimpleVirtualDesktopNotificationWrapper {
    fn drop(&mut self) {
        com_sta();
        let provider = get_iservice_provider().unwrap();
        let service = get_ivirtual_desktop_notification_service(&provider).unwrap();
        let cookie = self.cookie.borrow();

        println!("Drop notification with cookie {}", *cookie);

        unsafe { service.unregister(*cookie).panic_if_failed() };
    }
}

#[derive(Clone)]
#[windows::core::implement(IVirtualDesktopNotification)]
struct SimpleVirtualDesktopNotification {
    number_times_desktop_changed: Rc<RefCell<u32>>,
}

// Allow unused variable warnings
#[allow(unused_variables)]
impl IVirtualDesktopNotification_Impl for SimpleVirtualDesktopNotification {
    unsafe fn current_virtual_desktop_changed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_old: ComIn<IVirtualDesktop>,
        desktop_new: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        // debug_desktop(&desktop_new, "Desktop changed");
        *self.number_times_desktop_changed.borrow_mut() += 1;
        HRESULT(0)
    }

    unsafe fn virtual_desktop_wallpaper_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop wallpaper changed");
        HRESULT(0)
    }

    unsafe fn virtual_desktop_created(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop created");
        HRESULT(0)
    }

    unsafe fn virtual_desktop_destroy_begin(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        // Desktop destroyed is not anymore in the stack
        debug_desktop(&desktop_destroyed, "Desktop destroy begin");
        debug_desktop(&desktop_fallback, "Desktop destroy fallback");
        HRESULT(0)
    }

    unsafe fn virtual_desktop_destroy_failed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        HRESULT(0)
    }

    unsafe fn virtual_desktop_destroyed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        // Desktop destroyed is not anymore in the stack
        debug_desktop(&desktop_destroyed, "Desktop destroyed");
        debug_desktop(&desktop_fallback, "Desktop destroyed fallback");
        HRESULT(0)
    }

    unsafe fn virtual_desktop_is_per_monitor_changed(&self, is_per_monitor: i32) -> HRESULT {
        println!("Desktop is per monitor changed: {}", is_per_monitor != 0);
        HRESULT(0)
    }

    unsafe fn virtual_desktop_moved(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
        old_index: i64,
        new_index: i64,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop moved");
        HRESULT(0)
    }

    unsafe fn virtual_desktop_name_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop renamed");
        HRESULT(0)
    }

    unsafe fn view_virtual_desktop_changed(&self, view: IApplicationView) -> HRESULT {
        let mut hwnd = HWND::default();
        view.get_thumbnail_window(&mut hwnd);
        println!("View in desktop changed, HWND {:?}", hwnd);
        HRESULT(0)
    }
}

#[cfg(test)]
mod tests {

    use windows::Win32::{
        Foundation::{LPARAM, WPARAM},
        System::Com::{CoCreateInstance, CoUninitialize, CLSCTX_ALL},
        UI::WindowsAndMessaging::PostThreadMessageW,
    };

    use super::*;

    #[test]
    fn test_mta_access_violation() {
        com_mta();

        let current_desktop = get_current_desktop().unwrap();

        let notification_thread = std::thread::spawn(move || {
            com_mta();
            println!("Notification thread {:?}", std::thread::current().id());
            let _notification = SimpleVirtualDesktopNotificationWrapper::new().unwrap();
            std::thread::sleep(Duration::from_secs(4));
            let value = _notification.number_times_desktop_changed.borrow_mut();
            *value
        });
        let switcher = std::thread::spawn(|| {
            for _ in 0..50 {
                switch_desktop(0).unwrap();
                std::thread::sleep(Duration::from_millis(4));
                switch_desktop(1).unwrap();
            }
        });
        let threads = (0..999)
            .map(|_| {
                std::thread::spawn(|| {
                    com_sta();
                    get_desktop(get_desktop(0).get_id().unwrap())
                        .get_index()
                        .unwrap();
                    std::thread::sleep(std::time::Duration::from_millis(1));
                    get_desktop(get_desktop(1).get_id().unwrap())
                        .get_index()
                        .unwrap();
                    std::thread::sleep(std::time::Duration::from_millis(2));
                    get_desktop(get_desktop(2).get_id().unwrap())
                        .get_index()
                        .unwrap();
                    // let desktops = get_desktops().unwrap();
                    // println!("Desktops {:?}", desktops);
                })
            })
            .collect::<Vec<_>>();

        // Join all threads
        for thread in threads {
            thread.join().unwrap();
        }
        switcher.join().unwrap();

        // Delay of 100ms
        std::thread::sleep(std::time::Duration::from_millis(100));

        // Switch back to original desktop
        switch_desktop(current_desktop).unwrap();
    }

    #[test]
    fn test_drop_after_counitialize() {
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
        }

        let provider: IServiceProvider = unsafe {
            CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_ALL)
                .map_err(map_win_err)
                .unwrap()
        };

        drop(provider); // Without this we get IUnknown Release access violation

        unsafe {
            CoUninitialize();
        }
    }

    #[test]
    fn test_sta_notifications() {
        com_sta();

        let winapi_thread_id_pair = Arc::new((Mutex::new(0 as u32), Condvar::new()));
        let winapi_thread_id_pair_2 = Arc::clone(&winapi_thread_id_pair);

        // Notification thread is also in STA mode, and it sends out the Windows API compatible thread ID using a mutex and condvar
        let notification_thread = std::thread::spawn(move || {
            println!("Notification thread {:?}", std::thread::current().id());

            com_sta();
            let _notification = SimpleVirtualDesktopNotificationWrapper::new().unwrap();

            {
                // Send the current thread id to parent thread
                let (lock, cvar) = &*winapi_thread_id_pair_2;
                let mut started = lock.lock().unwrap();
                *started = unsafe { GetCurrentThreadId() };
                cvar.notify_one();
            }

            // STA message loop, this is required as the notifications are pushed to message queue to be processed
            let mut msg = MSG::default();
            unsafe {
                while GetMessageW(&mut msg, HWND::default(), 0, 0).as_bool() {
                    if msg.message == WM_USER + 0x10 {
                        PostQuitMessage(0);
                    }
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
            }

            // Return the number of times the desktop changed
            let value = _notification.number_times_desktop_changed.borrow_mut();
            *value
        });

        // Get the thread id sent by the notification thread
        let win_thread_id = {
            let (lock, cvar) = &*winapi_thread_id_pair;
            let mut started = lock.lock().unwrap();
            while *started == 0 {
                started = cvar.wait(started).unwrap();
            }
            println!("Started? {}", started);
            *started
        };

        // Start switching desktops in rapid fashion
        let current_desktop = get_current_desktop().unwrap().get_index().unwrap();

        for _ in 0..1999 {
            switch_desktop(0).unwrap();
            // std::thread::sleep(Duration::from_millis(4));
            switch_desktop(1).unwrap();
        }

        // Finally return to same desktop we were
        std::thread::sleep(Duration::from_millis(13));
        switch_desktop(current_desktop).unwrap();

        // Windows pushes the notification events to the queue, but it takes a while for them to be processed, I don't know a way to wait out until the push queue is empty
        //
        // 8 seconds is not accurate, increase if test fails
        std::thread::sleep(Duration::from_secs(8));

        // Send a message to the notification thread to quit and join it
        unsafe {
            PostThreadMessageW(
                win_thread_id,
                WM_USER + 0x10,
                WPARAM::default(),
                LPARAM::default(),
            );
        }

        let changes = notification_thread.join().unwrap();

        std::thread::sleep(Duration::from_secs(1));
        std::thread::sleep(Duration::from_secs(1));

        println!("Desktop changes {}", changes);

        // 5*2 + 1 = 11
        // assert_eq!(changes, 3999);
    }

    #[test]
    fn test_mta_notifications() {
        // Note: Test gives access violation if this thread is in MTA mode
        // Note: Test gives access violation if this thread is in STA mode and notification thread is in MTA mode

        com_sta();

        let notification_thread = std::thread::spawn(move || {
            com_mta();

            println!("Notification thread {:?}", std::thread::current().id());
            let _notification = SimpleVirtualDesktopNotificationWrapper::new().unwrap();
            std::thread::sleep(Duration::from_secs(15));
            let value = _notification.number_times_desktop_changed.borrow_mut();
            *value
        });

        // Start switching desktops in rapid fashion
        let current_desktop = get_current_desktop().unwrap().get_index().unwrap();

        for _ in 0..999 {
            switch_desktop(0).unwrap();
            // std::thread::sleep(Duration::from_millis(4));
            switch_desktop(1).unwrap();
        }

        // Finally return to same desktop we were
        std::thread::sleep(Duration::from_millis(13));
        switch_desktop(current_desktop).unwrap();
        let changes = notification_thread.join().unwrap();
        println!("Desktop changes {}", changes);
        // 5*2 + 1 = 11
        assert_eq!(changes, 1999);
    }

    #[test] // TODO: Commented out, use only on occasion when needed!
    fn test_listener_manual() {
        println!("Main thread is {:?}", std::thread::current().id());

        std::thread::spawn(|| {
            com_sta();

            println!("Notification thread {:?}", std::thread::current().id());
            let _notification = SimpleVirtualDesktopNotificationWrapper::new().unwrap();
            let mut msg = MSG::default();
            unsafe {
                while GetMessageW(&mut msg, HWND::default(), 0, 0).as_bool() {
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
            }
        })
        .join()
        .unwrap();
    }

    /// This test switched desktop and prints out the changed desktop
    #[test]
    fn test_register_notifications() {
        let _notification = SimpleVirtualDesktopNotificationWrapper::new();
        let manager = get_ivirtual_desktop_manager_internal().unwrap();

        // Get current desktop
        let mut current_desk: Option<IVirtualDesktop> = None;
        unsafe {
            manager
                .get_current_desktop(0, &mut current_desk)
                .panic_if_failed();
        }
        assert_eq!(current_desk.is_none(), false);
        let current_desk = current_desk.unwrap();

        let mut gid = GUID::default();
        unsafe { current_desk.get_id(&mut gid).panic_if_failed() };

        let mut name = HSTRING::new();
        unsafe { current_desk.get_name(&mut name).panic_if_failed() };

        println!("Current desktop: {} {:?}", name.to_string_lossy(), gid);

        // Get adjacent desktop
        let mut next_idesk: Option<IVirtualDesktop> = None;
        unsafe {
            manager
                .get_adjacent_desktop(ComIn::new(&current_desk), 3, &mut next_idesk)
                .panic_if_failed();
        }
        let next_desk = next_idesk.unwrap();
        let mut gid = GUID::default();
        unsafe { next_desk.get_id(&mut gid).panic_if_failed() };

        let mut name = HSTRING::new();
        unsafe { next_desk.get_name(&mut name).panic_if_failed() };

        // Switch to next desktop and back again
        unsafe {
            manager
                .switch_desktop(0, ComIn::new(&next_desk.into()))
                .panic_if_failed()
        };
        unsafe {
            manager
                .switch_desktop(0, ComIn::new(&current_desk))
                .panic_if_failed()
        };
        std::thread::sleep(Duration::from_millis(5)); // This is not accurate, increase when needed

        // TODO: Test that desktop changed twice
        // let mut desktop_changed_count = 0;
        // while let Ok(_) = rx.try_recv() {
        //     desktop_changed_count += 1;
        // }
        // assert_eq!(desktop_changed_count, 2);
    }

    #[test]
    fn test_list_desktops() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap() };

        let manager: IVirtualDesktopManagerInternal =
            get_ivirtual_desktop_manager_internal().unwrap();

        // let desktops: *mut IObjectArray = std::ptr::null_mut();
        let mut desktops = None;

        unsafe { manager.get_desktops(0, &mut desktops).panic_if_failed() };

        let desktops = desktops.unwrap();

        // Iterate desktops
        let count = unsafe { desktops.GetCount().unwrap() };
        assert_ne!(count, 0);

        for i in 0..count {
            let desktop: IVirtualDesktop = unsafe { desktops.GetAt(i).unwrap() };

            let mut gid = GUID::default();
            unsafe { desktop.get_id(&mut gid).panic_if_failed() };

            let mut name = HSTRING::new();
            unsafe { desktop.get_name(&mut name).panic_if_failed() };

            println!("Desktop: {} {:?}", name.to_string_lossy(), gid);
        }
    }
}
