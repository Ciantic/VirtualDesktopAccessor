use crate::hresult::HRESULT;

use super::raw2::*;

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
        let number_times_desktop_changed = Rc::new(RefCell::new(0));

        let ptr = SimpleVirtualDesktopNotification {
            number_times_desktop_changed: number_times_desktop_changed.clone(),
        };
        let mut notification = Pin::new(Box::new(SimpleVirtualDesktopNotificationWrapper {
            cookie: 0,
            ptr: ptr.into(),
            number_times_desktop_changed,
        }));

        notification.cookie = com_objects().register_for_notifications(&notification.ptr)?;
        println!(
            "Registered notification {} {:?}",
            notification.cookie,
            std::thread::current().id()
        );

        Ok(notification)
    }
}

impl Drop for SimpleVirtualDesktopNotificationWrapper {
    fn drop(&mut self) {
        let cookie = self.cookie.borrow();
        com_objects().unregister_for_notifications(*cookie).unwrap();
    }
}

#[windows::core::implement(IVirtualDesktopNotification)]
pub struct SimpleVirtualDesktopNotification {
    number_times_desktop_changed: Rc<RefCell<u32>>,
}

fn debug_desktop(desktop_new: &IVirtualDesktop, prefix: &str) {
    let mut gid = GUID::default();
    unsafe { desktop_new.get_id(&mut gid).panic_if_failed() };

    let name = "";

    // let mut name = HSTRING::new();
    // unsafe { desktop_new.get_name(&mut name).panic_if_failed() };

    println!(
        "{}: {:?} {:?} {:?}",
        prefix,
        gid,
        name.to_string(),
        std::thread::current().id()
    );
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
        let index = get_current_desktop().unwrap().get_index().unwrap();
        debug_desktop(&desktop_new, &format!("Desktop changed {}", index));
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
mod test {
    use super::*;
    use crate::comapi::raw2::{com_objects, ComObjectsAsResult};
    use std::{
        pin::Pin,
        sync::{Arc, Condvar, Mutex},
        time::Duration,
    };
    use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

    #[test]
    fn test_switch_desktops_rapidly() {
        let objects = com_objects();
        println!("This thread is {:?}", std::thread::current().id());
        let notifications_thread = std::thread::spawn(|| {
            let _notifications = SimpleVirtualDesktopNotificationWrapper::new().unwrap();
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
        });

        // Start switching desktops in rapid fashion
        let current_desktop = objects.get_current_desktop().unwrap();

        for _ in 0..4999 {
            com_objects().switch_desktop(&0.into()).unwrap();
            // std::thread::sleep(Duration::from_millis(4));
            com_objects().switch_desktop(&1.into()).unwrap();
        }

        // Finally return to same desktop we were
        std::thread::sleep(Duration::from_millis(13));
        objects.switch_desktop(&current_desktop).unwrap();
        std::thread::sleep(Duration::from_millis(13));
        println!("End of program");

        notifications_thread.join().unwrap();
    }

    #[test]
    fn test_listener_manual() {
        println!("This thread is {:?}", std::thread::current().id());
        let notifications_thread = std::thread::spawn(|| {
            let _notifications = SimpleVirtualDesktopNotificationWrapper::new().unwrap();
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
        });

        notifications_thread.join().unwrap();
    }
}
