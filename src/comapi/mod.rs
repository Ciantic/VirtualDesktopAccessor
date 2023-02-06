pub mod interfaces;

use self::interfaces::*;
use crate::{DesktopID, Error, HRESULT};
use std::{ffi::c_void, time::Duration};
use windows::{
    core::{Interface, Vtable, HSTRING},
    Win32::{
        System::Com::{CoCreateInstance, CoInitializeEx, CLSCTX_ALL, COINIT_MULTITHREADED},
        UI::Shell::Common::IObjectArray,
    },
};

type Result<T> = std::result::Result<T, Error>;

pub fn create_service_provider() -> Result<IServiceProvider> {
    return unsafe {
        CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_ALL)
            .map_err(|er| Error::ComError(HRESULT::from_i32(er.code().0)))
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

mod tests {

    use std::sync::{mpsc::Sender, Mutex};

    use super::*;

    #[windows::core::implement(IVirtualDesktopNotification)]
    struct TestVDNotifications {
        number_times_desktop_changed: Sender<()>,
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
            let mut gid = DesktopID::default();
            unsafe { desktop_new.get_id(&mut gid).panic_if_failed() };

            let mut name = HSTRING::new();
            unsafe { desktop_new.get_name(&mut name).panic_if_failed() };

            println!("Desktop changed: {:?} {:?}", gid, name.to_string());
            self.number_times_desktop_changed.send(()).unwrap();
            HRESULT(0)
        }

        unsafe fn virtual_desktop_wallpaper_changed(
            &self,
            desktop: ComIn<IVirtualDesktop>,
            name: HSTRING,
        ) -> HRESULT {
            HRESULT(0)
        }

        unsafe fn virtual_desktop_created(
            &self,
            monitors: ComIn<IObjectArray>,
            desktop: ComIn<IVirtualDesktop>,
        ) -> HRESULT {
            HRESULT(0)
        }

        unsafe fn virtual_desktop_destroy_begin(
            &self,
            monitors: ComIn<IObjectArray>,
            desktop_destroyed: ComIn<IVirtualDesktop>,
            desktop_fallback: ComIn<IVirtualDesktop>,
        ) -> HRESULT {
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
            HRESULT(0)
        }

        unsafe fn virtual_desktop_is_per_monitor_changed(&self, is_per_monitor: i32) -> HRESULT {
            HRESULT(0)
        }

        unsafe fn virtual_desktop_moved(
            &self,
            monitors: ComIn<IObjectArray>,
            desktop: ComIn<IVirtualDesktop>,
            old_index: i64,
            new_index: i64,
        ) -> HRESULT {
            HRESULT(0)
        }

        unsafe fn virtual_desktop_name_changed(
            &self,
            desktop: ComIn<IVirtualDesktop>,
            name: HSTRING,
        ) -> HRESULT {
            HRESULT(0)
        }

        unsafe fn view_virtual_desktop_changed(&self, view: IApplicationView) -> HRESULT {
            HRESULT(0)
        }
    }

    /// This test switched desktop and prints out the changed desktop
    #[test]
    fn test_register_notifications() {
        unsafe { CoInitializeEx(None, COINIT_MULTITHREADED).unwrap() };

        let provider = create_service_provider().unwrap();
        let service = create_vd_notification_service(&provider).unwrap();
        let manager = create_vd_manager(&provider).unwrap();
        let (tx, rx) = std::sync::mpsc::channel();
        let notification = TestVDNotifications {
            number_times_desktop_changed: tx,
        };
        let mut registration_cookie = 0;
        unsafe {
            service
                .register(ComIn::new(&notification.into()), &mut registration_cookie)
                .panic_if_failed();
            assert_ne!(registration_cookie, 0);
        }

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
        std::thread::sleep(Duration::from_millis(50));
        unsafe {
            manager
                .switch_desktop(0, ComIn::new(&current_desk))
                .panic_if_failed()
        };
        std::thread::sleep(Duration::from_millis(230));

        // Test that desktop changed twice
        let mut desktop_changed_count = 0;
        while let Ok(_) = rx.try_recv() {
            desktop_changed_count += 1;
        }
        assert_eq!(desktop_changed_count, 2);

        // Unregister notifications
        unsafe {
            service.unregister(registration_cookie).panic_if_failed();
        }
    }

    #[test]
    fn test_list_desktops() {
        unsafe { CoInitializeEx(None, COINIT_MULTITHREADED).unwrap() };

        let provider: IServiceProvider =
            unsafe { CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_ALL).unwrap() };

        let mut obj = std::ptr::null_mut::<c_void>();
        unsafe {
            provider
                .query_service(
                    &CLSID_VirtualDesktopManagerInternal,
                    &IVirtualDesktopManagerInternal::IID,
                    &mut obj,
                )
                .panic_if_failed();
        }
        assert_eq!(obj.is_null(), false);

        let manager: IVirtualDesktopManagerInternal =
            unsafe { IVirtualDesktopManagerInternal::from_raw(obj) };

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
