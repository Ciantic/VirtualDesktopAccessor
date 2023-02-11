pub mod interfaces;

use self::interfaces::*;
use crate::{Error, HRESULT};
use std::ffi::c_void;
use windows::{
    core::{Interface, Vtable, GUID, HSTRING},
    Win32::{
        System::Com::{
            CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED,
        },
        UI::Shell::Common::IObjectArray,
    },
};

type HWND = u32;

type APPID_PWSTR = *mut *mut std::ffi::c_void;

type Result<T> = std::result::Result<T, Error>;

pub(crate) struct ComInit();

impl ComInit {
    pub fn new() -> Self {
        unsafe {
            // Notice: Only COINIT_APARTMENTTHREADED works correctly!
            //
            // Not COINIT_MULTITHREADED or CoIncrementMTAUsage, they cause a seldom crashes in threading tests.
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
        }
        ComInit()
    }
}

impl Drop for ComInit {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

thread_local! {
    pub(crate) static COM_INIT: ComInit = ComInit::new();
}

fn map_win_err(er: ::windows::core::Error) -> Error {
    Error::ComError(HRESULT::from_i32(er.code().0))
}

fn create_service_provider() -> Result<IServiceProvider> {
    COM_INIT.with(|_| unsafe {
        CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_ALL).map_err(map_win_err)
    })
}

fn create_vd_notification_service(
    provider: &IServiceProvider,
) -> Result<IVirtualDesktopNotificationService> {
    COM_INIT.with(|_| {
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
    })
}

fn create_vd_manager(provider: &IServiceProvider) -> Result<IVirtualDesktopManagerInternal> {
    COM_INIT.with(|_| {
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
    })
}

fn create_view_collection(provider: &IServiceProvider) -> Result<IApplicationViewCollection> {
    COM_INIT.with(|_| {
        let mut obj = std::ptr::null_mut::<c_void>();
        unsafe {
            provider
                .query_service(
                    &IApplicationViewCollection::IID,
                    &IApplicationViewCollection::IID,
                    &mut obj,
                )
                .as_result()?;
        }
        assert_eq!(obj.is_null(), false);

        Ok(unsafe { IApplicationViewCollection::from_raw(obj) })
    })
}

fn create_pinned_apps(provider: &IServiceProvider) -> Result<IVirtualDesktopPinnedApps> {
    COM_INIT.with(|_| {
        let mut obj = std::ptr::null_mut::<c_void>();
        unsafe {
            provider
                .query_service(
                    &CLSID_VirtualDesktopPinnedApps,
                    &IVirtualDesktopPinnedApps::IID,
                    &mut obj,
                )
                .as_result()?;
        }
        assert_eq!(obj.is_null(), false);

        Ok(unsafe { IVirtualDesktopPinnedApps::from_raw(obj) })
    })
}

fn get_idesktops_array(manager: &IVirtualDesktopManagerInternal) -> Result<IObjectArray> {
    COM_INIT.with(|_| {
        let mut desktops = None;
        unsafe { manager.get_desktops(0, &mut desktops).as_result()? }
        Ok(desktops.unwrap())
    })
}

fn get_idesktop_number(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
) -> Result<u32> {
    COM_INIT.with(|_| {
        let desktops = get_idesktops_array(manager)?;
        let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
        for i in 0..count {
            let d: IVirtualDesktop = unsafe { desktops.GetAt(i).map_err(map_win_err)? };
            if d == *desktop {
                return Ok(i);
            }
        }
        Err(Error::DesktopNotFound)
    })
}

fn get_idesktop_wallpaper(desktop: &IVirtualDesktop) -> Result<String> {
    COM_INIT.with(|_| {
        let mut name = HSTRING::default();
        unsafe { desktop.get_wallpaper(&mut name).as_result()? }
        Ok(name.to_string_lossy())
    })
}

fn set_idesktop_wallpaper(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
    wallpaper_path: &str,
) -> Result<()> {
    COM_INIT.with(|_| {
        let name = HSTRING::from(wallpaper_path);
        unsafe {
            manager
                .set_wallpaper(ComIn::new(&desktop), name)
                .as_result()?
        }
        Ok(())
    })
}

fn get_idesktop_guid(desktop: &IVirtualDesktop) -> Result<GUID> {
    COM_INIT.with(|_| {
        let mut guid = GUID::default();
        unsafe { desktop.get_id(&mut guid).as_result()? }
        Ok(guid)
    })
}

fn get_idesktop_name(desktop: &IVirtualDesktop) -> Result<String> {
    COM_INIT.with(|_| {
        let mut name = HSTRING::default();
        unsafe { desktop.get_name(&mut name).as_result()? }
        Ok(name.to_string_lossy())
    })
}

fn set_idesktop_name(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
    name: &str,
) -> Result<()> {
    COM_INIT.with(|_| {
        let name = HSTRING::from(name);
        unsafe { manager.set_name(ComIn::new(&desktop), name).as_result()? }
        Ok(())
    })
}

fn get_idesktop_by_number(
    manager: &IVirtualDesktopManagerInternal,
    index: u32,
) -> Result<IVirtualDesktop> {
    COM_INIT.with(|_| {
        let desktops = get_idesktops_array(manager)?;
        let desktop: IVirtualDesktop = unsafe { desktops.GetAt(index).map_err(map_win_err)? };
        Ok(desktop)
    })
}

fn get_idesktop_by_guid(
    manager: &IVirtualDesktopManagerInternal,
    guid: &GUID,
) -> Result<IVirtualDesktop> {
    COM_INIT.with(|_| {
        let mut idesktop = None;
        unsafe {
            manager.find_desktop(guid, &mut idesktop).as_result()?;
        }
        idesktop.ok_or(Error::DesktopNotFound)
    })
}

fn get_idesktops(manager: &IVirtualDesktopManagerInternal) -> Result<Vec<IVirtualDesktop>> {
    COM_INIT.with(|_| {
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
    })
}

fn get_current_idesktop(manager: &IVirtualDesktopManagerInternal) -> Result<IVirtualDesktop> {
    COM_INIT.with(|_| {
        let mut desktop = None;
        unsafe { manager.get_current_desktop(0, &mut desktop).as_result()? }
        desktop.ok_or(Error::DesktopNotFound)
    })
}

fn switch_to_idesktop(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
) -> Result<()> {
    COM_INIT.with(|_| {
        unsafe {
            manager
                .switch_desktop(0, ComIn::new(&desktop))
                .as_result()?
        }
        Ok(())
    })
}

fn create_idesktop(manager: &IVirtualDesktopManagerInternal) -> Result<IVirtualDesktop> {
    COM_INIT.with(|_| {
        let mut desktop = None;
        unsafe { manager.create_desktop(0, &mut desktop).as_result()? }
        desktop.ok_or(Error::CreateDesktopFailed)
    })
}

fn remove_idesktop(
    manager: &IVirtualDesktopManagerInternal,
    remove_desktop: &IVirtualDesktop,
    fallback_desktop: &IVirtualDesktop,
) -> Result<()> {
    COM_INIT.with(|_| unsafe {
        manager
            .remove_desktop(ComIn::new(remove_desktop), ComIn::new(fallback_desktop))
            .as_result()
            .map_err(|_| Error::RemoveDesktopFailed)
    })
}

fn get_iapplication_id_for_view(view: &IApplicationView) -> Result<APPID_PWSTR> {
    COM_INIT.with(|_| {
        let mut app_id: APPID_PWSTR = std::ptr::null_mut();
        unsafe {
            view.get_app_user_model_id(&mut app_id as *mut _ as *mut _)
                .as_result()?
        }
        Ok(app_id)
    })
}

fn get_iapplication_view_for_hwnd(
    view_collection: &IApplicationViewCollection,
    hwnd: HWND,
) -> Result<IApplicationView> {
    COM_INIT.with(|_| {
        let mut view = None;
        unsafe {
            view_collection
                .get_view_for_hwnd(hwnd, &mut view)
                .as_result()
                .map_err(|er| {
                    // View does not exist
                    if er == Error::ComError(HRESULT(0x8002802B)) {
                        Error::WindowNotFound
                    } else {
                        er
                    }
                })?
        }
        view.ok_or(Error::WindowNotFound)
    })
}

fn is_view_pinned(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<bool> {
    COM_INIT.with(|_| {
        let mut is_pinned = false;
        unsafe {
            apps.is_view_pinned(ComIn::new(&view), &mut is_pinned)
                .as_result()?
        }
        Ok(is_pinned)
    })
}

fn pin_view(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<()> {
    COM_INIT.with(|_| unsafe { apps.pin_view(ComIn::new(&view)).as_result() })
}

fn upin_view(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<()> {
    COM_INIT.with(|_| unsafe { apps.unpin_view(ComIn::new(&view)).as_result() })
}

fn is_app_id_pinned(apps: &IVirtualDesktopPinnedApps, app_id: APPID_PWSTR) -> Result<bool> {
    COM_INIT.with(|_| {
        let mut is_pinned = false;
        unsafe {
            apps.is_app_pinned(app_id as *mut _, &mut is_pinned)
                .as_result()?
        }
        Ok(is_pinned)
    })
}

fn pin_app_id(apps: &IVirtualDesktopPinnedApps, app_id: APPID_PWSTR) -> Result<()> {
    COM_INIT.with(|_| unsafe { apps.pin_app(app_id as *mut _).as_result() })
}

fn unpin_app_id(apps: &IVirtualDesktopPinnedApps, app_id: APPID_PWSTR) -> Result<()> {
    COM_INIT.with(|_| unsafe { apps.unpin_app(app_id as *mut _).as_result() })
}

pub mod pinning {
    type HWND = u32;
    use super::*;

    /// Is window pinned?
    pub fn is_pinned_window(hwnd: HWND) -> Result<bool> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let view_collection = create_view_collection(&provider)?;
            let apps = create_pinned_apps(&provider)?;
            let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
            is_view_pinned(&apps, view)
        })
    }

    /// Pin window
    pub fn pin_window(hwnd: HWND) -> Result<()> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let view_collection = create_view_collection(&provider)?;
            let apps = create_pinned_apps(&provider)?;
            let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
            pin_view(&apps, view)
        })
    }

    /// Unpin window
    pub fn unpin_window(hwnd: HWND) -> Result<()> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let view_collection = create_view_collection(&provider)?;
            let apps = create_pinned_apps(&provider)?;
            let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
            upin_view(&apps, view)
        })
    }

    /// Is pinned app
    pub fn is_pinned_app(hwnd: HWND) -> Result<bool> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let view_collection = create_view_collection(&provider)?;
            let apps = create_pinned_apps(&provider)?;
            let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
            let app_id = get_iapplication_id_for_view(&view)?;
            is_app_id_pinned(&apps, app_id)
        })
    }

    /// Pin app
    pub fn pin_app(hwnd: HWND) -> Result<()> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let view_collection = create_view_collection(&provider)?;
            let apps = create_pinned_apps(&provider)?;
            let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
            let app_id = get_iapplication_id_for_view(&view)?;
            pin_app_id(&apps, app_id)
        })
    }

    /// Unpin app
    pub fn unpin_app(hwnd: HWND) -> Result<()> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let view_collection = create_view_collection(&provider)?;
            let apps = create_pinned_apps(&provider)?;
            let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
            let app_id = get_iapplication_id_for_view(&view)?;
            unpin_app_id(&apps, app_id)
        })
    }
}

pub mod normal {
    use super::*;
    use std::fmt::Debug;
    use windows::core::GUID;

    #[derive(Copy, Clone, PartialEq)]
    pub struct Desktop {
        pub(crate) id: GUID,
    }

    impl Debug for Desktop {
        fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            write!(f, "Desktop({:?})", self.id)
        }
    }

    impl Desktop {
        pub(crate) fn empty() -> Desktop {
            Desktop {
                id: GUID::default(),
            }
        }

        pub fn new() -> Result<Desktop> {
            COM_INIT.with(|_| {
                let provider = create_service_provider()?;
                let manager = create_vd_manager(&provider)?;
                let desktop = create_idesktop(&manager)?;
                let id = get_idesktop_guid(&desktop)?;
                Ok(Desktop { id })
            })
        }

        pub fn get_name(&self) -> Result<String> {
            COM_INIT.with(|_| {
                let provider = create_service_provider()?;
                let manager = create_vd_manager(&provider)?;
                let desktop = get_idesktop_by_guid(&manager, &self.id)?;
                get_idesktop_name(&desktop)
            })
        }

        pub fn set_name(&self, name: &str) -> Result<()> {
            COM_INIT.with(|_| {
                let provider = create_service_provider()?;
                let manager = create_vd_manager(&provider)?;
                let idesktop = get_idesktop_by_guid(&manager, &self.id)?;
                set_idesktop_name(&manager, &idesktop, name)
            })
        }

        pub fn get_index(&self) -> Result<u32> {
            COM_INIT.with(|_| {
                let provider = create_service_provider()?;
                let manager = create_vd_manager(&provider)?;
                let idesktop = get_idesktop_by_guid(&manager, &self.id)?;
                let index = get_idesktop_number(&manager, &idesktop)?;
                Ok(index)
            })
        }

        pub fn get_wallpaper(&self) -> Result<String> {
            COM_INIT.with(|_| {
                let provider = create_service_provider()?;
                let manager = create_vd_manager(&provider)?;
                let idesktop = get_idesktop_by_guid(&manager, &self.id)?;
                get_idesktop_wallpaper(&idesktop)
            })
        }

        pub fn set_wallpaper(&self, path: &str) -> Result<()> {
            COM_INIT.with(|_| {
                let provider = create_service_provider()?;
                let manager = create_vd_manager(&provider)?;
                let idesktop = get_idesktop_by_guid(&manager, &self.id)?;
                set_idesktop_wallpaper(&manager, &idesktop, path)
            })
        }

        pub fn switch_to(&self) -> Result<()> {
            COM_INIT.with(|_| {
                let provider = create_service_provider()?;
                let manager = create_vd_manager(&provider)?;
                let idesktop = get_idesktop_by_guid(&manager, &self.id)?;
                switch_to_idesktop(&manager, &idesktop)
            })
        }

        pub fn remove(&self, fallback_desktop: &Desktop) -> Result<()> {
            COM_INIT.with(|_| {
                let provider = create_service_provider()?;
                let manager = create_vd_manager(&provider)?;
                let idesktop = get_idesktop_by_guid(&manager, &self.id)?;
                let fallback_idesktop = get_idesktop_by_guid(&manager, &fallback_desktop.id)?;
                remove_idesktop(&manager, &idesktop, &fallback_idesktop)
            })
        }

        pub fn get_id(&self) -> GUID {
            self.id
        }
    }

    pub fn get_desktops() -> Result<Vec<Desktop>> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let manager = create_vd_manager(&provider)?;
            let desktops: Result<Vec<Desktop>> = get_idesktops(&manager)?
                .into_iter()
                .map(|d| -> Result<Desktop> {
                    let mut desktop = Desktop::empty();
                    unsafe { d.get_id(&mut desktop.id).as_result()? };
                    Ok(desktop)
                })
                .collect();
            Ok(desktops?)
        })
    }

    pub fn get_desktop_by_guid(guid: GUID) -> Result<Desktop> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let manager = create_vd_manager(&provider)?;
            let desktop = get_idesktop_by_guid(&manager, &guid)?;
            let id = get_idesktop_guid(&desktop)?;
            Ok(Desktop { id })
        })
    }
}

mod numbered {
    use super::*;

    pub fn go_to_desktop_number(number: u32) -> Result<()> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let manager = create_vd_manager(&provider)?;
            let desktop = get_idesktop_by_number(&manager, number)?;
            switch_to_idesktop(&manager, &desktop)
        })
    }

    pub fn get_current_desktop_number() -> Result<u32> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let manager = create_vd_manager(&provider)?;
            let desktops = get_idesktops(&manager)?;
            let current = get_current_idesktop(&manager)?;
            for (i, desktop) in desktops.iter().enumerate() {
                if desktop == &current {
                    return Ok(i as u32);
                }
            }
            Err(Error::DesktopNotFound)
        })
    }

    pub fn set_name_by_desktop_number(number: u32, name: &str) -> Result<()> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let manager = create_vd_manager(&provider)?;
            let desktop = get_idesktop_by_number(&manager, number)?;
            let name = HSTRING::from(name);
            unsafe { manager.set_name(ComIn::new(&desktop), name).as_result() }
        })
    }

    pub fn get_name_by_desktop_number(number: u32) -> Result<String> {
        COM_INIT.with(|_| {
            let provider = create_service_provider()?;
            let manager = create_vd_manager(&provider)?;
            let desktop = get_idesktop_by_number(&manager, number)?;
            let mut name = HSTRING::new();
            unsafe { desktop.get_name(&mut name).as_result()? }
            Ok(name.to_string())
        })
    }
}

mod tests {
    use std::{
        rc::Rc,
        sync::{mpsc::Sender, Mutex},
        time::Duration,
    };

    use super::numbered::*;
    use super::*;

    fn debug_desktop(desktop_new: &IVirtualDesktop, prefix: &str) {
        let mut gid = GUID::default();
        unsafe { desktop_new.get_id(&mut gid).panic_if_failed() };

        let mut name = HSTRING::new();
        unsafe { desktop_new.get_name(&mut name).panic_if_failed() };

        let manager = create_vd_manager(&create_service_provider().unwrap()).unwrap();
        let number = get_idesktop_number(&manager, &desktop_new).unwrap_or(99999);

        println!("{}: {} {:?} {:?}", prefix, number, gid, name.to_string());
    }

    #[derive(Clone)]
    #[windows::core::implement(IVirtualDesktopNotification)]
    struct TestVDNotifications {
        cookie: Rc<Mutex<u32>>,
        number_times_desktop_changed: Rc<Sender<()>>,
    }

    impl TestVDNotifications {
        pub fn new(number_times_desktop_changed: Sender<()>) -> Result<Self> {
            COM_INIT.with(|_| {
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
            })
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
        let (tx, rx) = std::sync::mpsc::channel();
        let notification = TestVDNotifications::new(tx);

        std::thread::sleep(Duration::from_secs(12));
    }

    /// This test switched desktop and prints out the changed desktop
    #[test]
    fn test_register_notifications() {
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

            let mut gid = GUID::default();
            unsafe { desktop.get_id(&mut gid).panic_if_failed() };

            let mut name = HSTRING::new();
            unsafe { desktop.get_name(&mut name).panic_if_failed() };

            println!("Desktop: {} {:?}", name.to_string_lossy(), gid);
        }
    }
}
