pub mod interfaces;

use self::interfaces::*;
use crate::{DesktopID, Error, HRESULT};
use std::{ffi::c_void, time::Duration};
use windows::{
    core::{Interface, Vtable, HSTRING},
    Win32::{
        System::Com::{
            CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
            COINIT_MULTITHREADED,
        },
        UI::Shell::Common::IObjectArray,
    },
};

type Result<T> = std::result::Result<T, Error>;

struct ComInit();

impl ComInit {
    pub fn new() -> Self {
        unsafe {
            println!("CoInitializeEx COINIT_APARTMENTTHREADED");
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
        }
        ComInit()
    }
}

impl Drop for ComInit {
    fn drop(&mut self) {
        unsafe {
            println!("CoUninitialize");
            CoUninitialize();
        }
    }
}

thread_local! {
    static COM_INIT: ComInit = ComInit::new();
}

fn map_win_err(er: ::windows::core::Error) -> Error {
    Error::ComError(HRESULT::from_i32(er.code().0))
}

pub fn create_service_provider() -> Result<IServiceProvider> {
    return unsafe {
        CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_ALL).map_err(map_win_err)
    };
}

pub fn create_vd_notification_service(
    provider: &IServiceProvider,
) -> Result<IVirtualDesktopNotificationService> {
    let mut obj = std::ptr::null_mut::<c_void>();
    unsafe {
        provider
            .query_service(
                &CLSID_IVirtualNotificationService,
                &IVirtualDesktopNotificationService::IID,
                &mut obj,
            )
            .as_result()?
    }
    assert_eq!(obj.is_null(), false);

    Ok(unsafe { IVirtualDesktopNotificationService::from_raw(obj) })
}

pub fn create_vd_manager(provider: &IServiceProvider) -> Result<IVirtualDesktopManagerInternal> {
    let mut obj = std::ptr::null_mut::<c_void>();
    unsafe {
        provider
            .query_service(
                &CLSID_VirtualDesktopManagerInternal,
                &IVirtualDesktopManagerInternal::IID,
                &mut obj,
            )
            .as_result()?;
    }
    assert_eq!(obj.is_null(), false);

    Ok(unsafe { IVirtualDesktopManagerInternal::from_raw(obj) })
}

pub fn get_desktops_array(manager: &IVirtualDesktopManagerInternal) -> Result<IObjectArray> {
    let mut desktops = None;
    unsafe { manager.get_desktops(0, &mut desktops).as_result()? }
    Ok(desktops.unwrap())
}

pub fn get_desktop_number(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
) -> Result<i32> {
    let desktops = get_desktops_array(manager)?;
    let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
    for i in 0..count {
        let d: IVirtualDesktop = unsafe { desktops.GetAt(i).map_err(map_win_err)? };
        if d == *desktop {
            return Ok(i as i32);
        }
    }
    Err(Error::DesktopNotFound)
}

pub fn get_desktop_by_number(
    manager: &IVirtualDesktopManagerInternal,
    index: u32,
) -> Result<IVirtualDesktop> {
    let desktops = get_desktops_array(manager)?;
    let desktop: IVirtualDesktop = unsafe { desktops.GetAt(index).map_err(map_win_err)? };
    Ok(desktop)
}

pub fn get_desktops(manager: &IVirtualDesktopManagerInternal) -> Result<Vec<IVirtualDesktop>> {
    let mut desktops = None;
    unsafe { manager.get_desktops(0, &mut desktops).as_result()? }
    let desktops = desktops.unwrap();
    let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
    let mut idesktops = Vec::with_capacity(count as usize);
    for i in 0..count {
        let desktop: IVirtualDesktop = unsafe { desktops.GetAt(i).map_err(map_win_err)? };
        idesktops.push(desktop);
    }
    Ok(idesktops)
}

pub fn get_current_desktop(manager: &IVirtualDesktopManagerInternal) -> Result<IVirtualDesktop> {
    let mut desktop = None;
    unsafe { manager.get_current_desktop(0, &mut desktop).as_result()? }
    desktop.ok_or(Error::DesktopNotFound)
}

pub fn go_to_desktop_number(number: u32) -> Result<()> {
    let provider = create_service_provider()?;
    let manager = create_vd_manager(&provider)?;
    let desktop = get_desktop_by_number(&manager, number)?;
    unsafe { manager.switch_desktop(0, ComIn::new(&desktop)).as_result() }
}

pub fn get_current_desktop_number() -> Result<u32> {
    let provider = create_service_provider()?;
    let manager = create_vd_manager(&provider)?;
    let desktops = get_desktops(&manager)?;
    let current = get_current_desktop(&manager)?;
    for (i, desktop) in desktops.iter().enumerate() {
        if desktop == &current {
            return Ok(i as u32);
        }
    }
    Err(Error::DesktopNotFound)
}

pub fn debug_desktop(desktop_new: &IVirtualDesktop, prefix: &str) {
    let mut gid = DesktopID::default();
    unsafe { desktop_new.get_id(&mut gid).panic_if_failed() };

    let mut name = HSTRING::new();
    unsafe { desktop_new.get_name(&mut name).panic_if_failed() };

    let manager = create_vd_manager(&create_service_provider().unwrap()).unwrap();
    let number = get_desktop_number(&manager, &desktop_new).unwrap_or(-1);

    println!("{}: {} {:?} {:?}", prefix, number, gid, name.to_string());
}

mod tests {

    use std::{
        rc::Rc,
        sync::{mpsc::Sender, Mutex},
    };

    use windows::Win32::System::Com::{CoIncrementMTAUsage, COINIT_APARTMENTTHREADED};

    use super::*;

    #[derive(Clone)]
    #[windows::core::implement(IVirtualDesktopNotification)]
    struct TestVDNotifications {
        cookie: Rc<Mutex<u32>>,
        number_times_desktop_changed: Rc<Sender<()>>,
    }

    impl TestVDNotifications {
        pub fn new(number_times_desktop_changed: Sender<()>) -> Result<Self> {
            let provider = create_service_provider()?;
            let service = create_vd_notification_service(&provider)?;
            let notification = TestVDNotifications {
                cookie: Rc::new(Mutex::new(0)),
                number_times_desktop_changed: Rc::new(number_times_desktop_changed),
            };
            let inotification: IVirtualDesktopNotification = notification.clone().into();

            let mut cookie = 0;
            unsafe {
                service
                    .register(ComIn::unsafe_new_no_clone(inotification), &mut cookie)
                    .panic_if_failed();
                assert_ne!(cookie, 0);
            }
            *notification.cookie.lock().unwrap() = cookie;

            Ok(notification)
        }
    }

    impl Drop for TestVDNotifications {
        fn drop(&mut self) {
            let provider = create_service_provider().unwrap();
            let service = create_vd_notification_service(&provider).unwrap();
            let cookie = *self.cookie.lock().unwrap();
            unsafe { service.unregister(cookie) };
        }
    }

    // Allow unused variable warnings
    #[allow(unused_variables)]
    impl IVirtualDesktopNotification_Impl for TestVDNotifications {
        unsafe fn current_virtual_desktop_changed(
            &self,
            monitors: ComIn<IObjectArray>,
            desktop_old: ComIn<IVirtualDesktop>,
            desktop_new: ComIn<IVirtualDesktop>,
        ) -> HRESULT {
            debug_desktop(&desktop_new, "Desktop changed");
            self.number_times_desktop_changed.send(()).unwrap();
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
            let mut hwnd = 0 as _;
            view.get_thumbnail_window(&mut hwnd);
            println!("View in desktop changed, HWND {}", hwnd);
            HRESULT(0)
        }
    }

    #[test] // TODO: Commented out, use only on occasion when needed!
    fn test_threading_two() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap() };
        let (tx, rx) = std::sync::mpsc::channel();
        let notification = TestVDNotifications::new(tx);

        let current_desktop = get_current_desktop_number().unwrap();

        for _ in 0..999 {
            go_to_desktop_number(0).unwrap();
            // std::thread::sleep(Duration::from_millis(4));
            go_to_desktop_number(1).unwrap();
        }
        std::thread::sleep(Duration::from_millis(3));
        go_to_desktop_number(current_desktop).unwrap();
    }

    #[test] // TODO: Commented out, use only on occasion when needed!
    fn test_listener_manual() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap() };
        let (tx, rx) = std::sync::mpsc::channel();
        let notification = TestVDNotifications::new(tx);

        std::thread::sleep(Duration::from_secs(12));
    }

    /// This test switched desktop and prints out the changed desktop
    #[test]
    fn test_register_notifications() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap() };
        let (tx, rx) = std::sync::mpsc::channel();
        let notification = TestVDNotifications::new(tx);

        let provider = create_service_provider().unwrap();
        let service = create_vd_notification_service(&provider).unwrap();
        let manager = create_vd_manager(&provider).unwrap();

        // Get current desktop
        let mut current_desk: Option<IVirtualDesktop> = None;
        unsafe {
            manager
                .get_current_desktop(0, &mut current_desk)
                .panic_if_failed();
        }
        assert_eq!(current_desk.is_none(), false);
        let current_desk = current_desk.unwrap();

        let mut gid = DesktopID::default();
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
        let mut gid = DesktopID::default();
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

        // Test that desktop changed twice
        let mut desktop_changed_count = 0;
        while let Ok(_) = rx.try_recv() {
            desktop_changed_count += 1;
        }
        assert_eq!(desktop_changed_count, 2);
    }

    #[test]
    fn test_list_desktops() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap() };

        let provider = create_service_provider().unwrap();
        let manager: IVirtualDesktopManagerInternal = create_vd_manager(&provider).unwrap();

        // let desktops: *mut IObjectArray = std::ptr::null_mut();
        let mut desktops = None;

        unsafe { manager.get_desktops(0, &mut desktops).panic_if_failed() };

        let desktops = desktops.unwrap();

        // Iterate desktops
        let count = unsafe { desktops.GetCount().unwrap() };
        assert_ne!(count, 0);

        for i in 0..count {
            let desktop: IVirtualDesktop = unsafe { desktops.GetAt(i).unwrap() };

            let mut gid = DesktopID::default();
            unsafe { desktop.get_id(&mut gid).panic_if_failed() };

            let mut name = HSTRING::new();
            unsafe { desktop.get_name(&mut name).panic_if_failed() };

            println!("Desktop: {} {:?}", name.to_string_lossy(), gid);
        }
    }
}
