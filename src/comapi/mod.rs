pub mod interfaces;

use self::interfaces::*;
use crate::{Error, HRESULT};
use std::{cell::RefCell, ffi::c_void, sync::Mutex};
use windows::{
    core::{Interface, Vtable, GUID, HSTRING},
    Win32::{
        System::Com::{
            CoCreateInstance, CoDecrementMTAUsage, CoIncrementMTAUsage, CoInitializeEx,
            CoUninitialize, CLSCTX_ALL, COINIT, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED,
            CO_MTA_USAGE_COOKIE,
        },
        UI::Shell::Common::IObjectArray,
    },
};

type HWND_ = u32;

type APPID_PWSTR = *mut *mut std::ffi::c_void;

type Result<T> = std::result::Result<T, Error>;

enum ComInit {
    CoInitializeEx(COINIT),
    CoIncrementMTAUsage(CO_MTA_USAGE_COOKIE),
}

// Notice: Only COINIT_APARTMENTTHREADED works correctly for everything but listener!
//
// Not COINIT_MULTITHREADED or CoIncrementMTAUsage, they cause a seldom crashes in threading tests.

impl ComInit {
    pub fn new_ex(dwcoinit: COINIT) -> Self {
        unsafe {
            #[cfg(debug_assertions)]
            println!(
                "CoInitializeEx {:?} {:?}",
                dwcoinit,
                std::thread::current().id()
            );
            CoInitializeEx(None, dwcoinit).unwrap();
        }
        ComInit::CoInitializeEx(dwcoinit)
    }

    pub fn new_increment_mta() -> Self {
        let cookie = unsafe {
            #[cfg(debug_assertions)]
            println!("CoIncrementMTAUsage {:?}", std::thread::current().id());
            CoIncrementMTAUsage().unwrap()
        };
        ComInit::CoIncrementMTAUsage(cookie)
    }
}

impl Drop for ComInit {
    fn drop(&mut self) {
        match &self {
            ComInit::CoIncrementMTAUsage(cookie) => unsafe {
                #[cfg(debug_assertions)]
                println!(
                    "CoDecrementMTAUsage {:?} {:?}",
                    cookie,
                    std::thread::current().id()
                );
                CoDecrementMTAUsage(cookie.clone()).unwrap();
            },
            ComInit::CoInitializeEx(_) => unsafe {
                println!("CoUninitialize {:?}", std::thread::current().id());
                CoUninitialize();
            },
        }
    }
}

thread_local! {
    static COM_INIT: RefCell<Option<ComInit>>  = RefCell::new(None);
}

pub fn com_sta() {
    COM_INIT.with(|f| {
        let mut m = f.borrow_mut();
        if m.is_none() {
            // Single threaded apartment = COINIT_APARTMENTTHREADED
            *m = Some(ComInit::new_ex(COINIT_APARTMENTTHREADED));
        }
    });
}

pub fn com_mta() {
    COM_INIT.with(|f| {
        let mut m = f.borrow_mut();
        if m.is_none() {
            // Multi threaded apartment = COINIT_MULTITHREADED
            *m = Some(ComInit::new_ex(COINIT_MULTITHREADED));
        }
    });
}

// thread_local! {
//     pub(crate) static COM_INIT: ComInit = ComInit::new(COINIT_APARTMENTTHREADED);
// }

fn map_win_err(er: ::windows::core::Error) -> Error {
    Error::ComError(HRESULT::from_i32(er.code().0))
}

fn get_iservice_provider() -> Result<IServiceProvider> {
    com_sta();
    unsafe { CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_ALL).map_err(map_win_err) }
}

fn get_ivirtual_desktop_notification_service(
    provider: &IServiceProvider,
) -> Result<IVirtualDesktopNotificationService> {
    com_sta();
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

fn get_ivirtual_desktop_manager(provider: &IServiceProvider) -> Result<IVirtualDesktopManager> {
    com_sta();
    let mut obj = std::ptr::null_mut::<c_void>();
    unsafe {
        provider
            .query_service(
                &IVirtualDesktopManager::IID,
                &IVirtualDesktopManager::IID,
                &mut obj,
            )
            .as_result()?;
    }
    assert_eq!(obj.is_null(), false);

    Ok(unsafe { IVirtualDesktopManager::from_raw(obj) })
}

fn get_ivirtual_desktop_manager_internal(
    provider: &IServiceProvider,
) -> Result<IVirtualDesktopManagerInternal> {
    com_sta();
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

fn get_iapplication_view_collection(
    provider: &IServiceProvider,
) -> Result<IApplicationViewCollection> {
    com_sta();
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
}

fn get_ivirtual_desktop_pinned_apps(
    provider: &IServiceProvider,
) -> Result<IVirtualDesktopPinnedApps> {
    com_sta();
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
}

fn get_idesktops_array(manager: &IVirtualDesktopManagerInternal) -> Result<IObjectArray> {
    com_sta();
    let mut desktops = None;
    unsafe { manager.get_desktops(0, &mut desktops).as_result()? }
    Ok(desktops.unwrap())
}

fn get_idesktop_number(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
) -> Result<u32> {
    com_sta();
    let desktops = get_idesktops_array(manager)?;
    let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
    for i in 0..count {
        let d: IVirtualDesktop = unsafe { desktops.GetAt(i).map_err(map_win_err)? };
        if d == *desktop {
            return Ok(i);
        }
    }
    Err(Error::DesktopNotFound)
}

fn get_idesktop_wallpaper(desktop: &IVirtualDesktop) -> Result<String> {
    com_sta();
    let mut name = HSTRING::default();
    unsafe { desktop.get_wallpaper(&mut name).as_result()? }
    Ok(name.to_string_lossy())
}

fn set_idesktop_wallpaper(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
    wallpaper_path: &str,
) -> Result<()> {
    com_sta();
    let name = HSTRING::from(wallpaper_path);
    unsafe {
        manager
            .set_wallpaper(ComIn::new(&desktop), name)
            .as_result()?
    }
    Ok(())
}

fn get_idesktop_guid(desktop: &IVirtualDesktop) -> Result<GUID> {
    com_sta();
    let mut guid = GUID::default();
    unsafe { desktop.get_id(&mut guid).as_result()? }
    Ok(guid)
}

fn get_idesktop_name(desktop: &IVirtualDesktop) -> Result<String> {
    com_sta();
    let mut name = HSTRING::default();
    unsafe { desktop.get_name(&mut name).as_result()? }
    Ok(name.to_string_lossy())
}

fn set_idesktop_name(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
    name: &str,
) -> Result<()> {
    com_sta();
    let name = HSTRING::from(name);
    unsafe { manager.set_name(ComIn::new(&desktop), name).as_result()? }
    Ok(())
}

fn get_idesktop_by_number(
    manager: &IVirtualDesktopManagerInternal,
    index: u32,
) -> Result<IVirtualDesktop> {
    com_sta();
    let desktops = get_idesktops_array(manager)?;
    let desktop = unsafe { desktops.GetAt(index).map_err(map_win_err) };
    desktop.map_err(|_| Error::DesktopNotFound)
}

fn get_idesktop_by_guid(
    manager: &IVirtualDesktopManagerInternal,
    guid: &GUID,
) -> Result<IVirtualDesktop> {
    com_sta();
    let mut idesktop = None;
    unsafe {
        manager.find_desktop(guid, &mut idesktop).as_result()?;
    }
    idesktop.ok_or(Error::DesktopNotFound)
}

fn get_idesktops(manager: &IVirtualDesktopManagerInternal) -> Result<Vec<IVirtualDesktop>> {
    com_sta();
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

fn get_current_idesktop(manager: &IVirtualDesktopManagerInternal) -> Result<IVirtualDesktop> {
    com_sta();
    let mut desktop = None;
    unsafe { manager.get_current_desktop(0, &mut desktop).as_result()? }
    desktop.ok_or(Error::DesktopNotFound)
}

fn switch_to_idesktop(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
) -> Result<()> {
    com_sta();
    unsafe {
        manager
            .switch_desktop(0, ComIn::new(&desktop))
            .as_result()?
    }
    Ok(())
}

fn create_idesktop(manager: &IVirtualDesktopManagerInternal) -> Result<IVirtualDesktop> {
    com_sta();
    let mut desktop = None;
    unsafe { manager.create_desktop(0, &mut desktop).as_result()? }
    desktop.ok_or(Error::CreateDesktopFailed)
}

fn move_view_to_desktop(
    manager: &IVirtualDesktopManagerInternal,
    view: &IApplicationView,
    desktop: &IVirtualDesktop,
) -> Result<()> {
    com_sta();
    unsafe {
        manager
            .move_view_to_desktop(ComIn::new(view), ComIn::new(desktop))
            .as_result()
    }
}

fn remove_idesktop(
    manager: &IVirtualDesktopManagerInternal,
    remove_desktop: &IVirtualDesktop,
    fallback_desktop: &IVirtualDesktop,
) -> Result<()> {
    com_sta();
    unsafe {
        manager
            .remove_desktop(ComIn::new(remove_desktop), ComIn::new(fallback_desktop))
            .as_result()
            .map_err(|_| Error::RemoveDesktopFailed)
    }
}

fn get_iapplication_id_for_view(view: &IApplicationView) -> Result<APPID_PWSTR> {
    com_sta();
    let mut app_id: APPID_PWSTR = std::ptr::null_mut();
    unsafe {
        view.get_app_user_model_id(&mut app_id as *mut _ as *mut _)
            .as_result()?
    }
    Ok(app_id)
}

fn get_iapplication_view_for_hwnd(
    view_collection: &IApplicationViewCollection,
    hwnd: HWND_,
) -> Result<IApplicationView> {
    com_sta();
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
}

fn is_view_pinned(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<bool> {
    com_sta();
    let mut is_pinned = false;
    unsafe {
        apps.is_view_pinned(ComIn::new(&view), &mut is_pinned)
            .as_result()?
    }
    Ok(is_pinned)
}

fn pin_view(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<()> {
    com_sta();
    unsafe { apps.pin_view(ComIn::new(&view)).as_result() }
}

fn upin_view(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<()> {
    com_sta();
    unsafe { apps.unpin_view(ComIn::new(&view)).as_result() }
}

fn is_app_id_pinned(apps: &IVirtualDesktopPinnedApps, app_id: APPID_PWSTR) -> Result<bool> {
    com_sta();
    let mut is_pinned = false;
    unsafe {
        apps.is_app_pinned(app_id as *mut _, &mut is_pinned)
            .as_result()?
    }
    Ok(is_pinned)
}

fn pin_app_id(apps: &IVirtualDesktopPinnedApps, app_id: APPID_PWSTR) -> Result<()> {
    com_sta();
    unsafe { apps.pin_app(app_id as *mut _).as_result() }
}

fn unpin_app_id(apps: &IVirtualDesktopPinnedApps, app_id: APPID_PWSTR) -> Result<()> {
    com_sta();
    unsafe { apps.unpin_app(app_id as *mut _).as_result() }
}

fn _is_window_on_current_desktop(manager: &IVirtualDesktopManager, hwnd: HWND_) -> Result<bool> {
    com_sta();
    let mut is_on_desktop = false;
    unsafe {
        manager
            .is_window_on_current_desktop(hwnd, &mut is_on_desktop)
            .as_result()?
    }
    Ok(is_on_desktop)
}

fn get_idesktop_by_window(
    manager_internal: &IVirtualDesktopManagerInternal,
    manager: &IVirtualDesktopManager,
    hwnd: HWND_,
) -> Result<IVirtualDesktop> {
    com_sta();
    let mut desktop_id = GUID::default();
    unsafe {
        manager
            .get_desktop_by_window(hwnd, &mut desktop_id)
            .as_result()
            .map_err(|er| match er {
                // Window does not exist
                Error::ComError(HRESULT(0x8002802B)) => Error::WindowNotFound,
                _ => er,
            })?
    }
    if desktop_id == GUID::default() {
        return Err(Error::WindowNotFound);
    }

    get_idesktop_by_guid(manager_internal, &desktop_id)
}

pub mod windowing {
    type HWND = u32;
    use super::*;

    pub fn move_window_to_desktop_number(hwnd: HWND, number: u32) -> Result<()> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = get_idesktop_by_number(&manager, number)?;
        let vc = get_iapplication_view_collection(&provider)?;
        let view = get_iapplication_view_for_hwnd(&vc, hwnd)?;
        move_view_to_desktop(&manager, &view, &desktop)
    }

    pub fn is_window_on_current_desktop(hwnd: HWND) -> Result<bool> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager(&provider)?;
        _is_window_on_current_desktop(&manager, hwnd)
    }

    pub fn is_window_on_desktop_number(hwnd: HWND, number: u32) -> Result<bool> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager(&provider)?;
        let man2 = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = get_idesktop_by_number(&man2, number)?;
        let desktop2 = get_idesktop_by_window(&man2, &manager, hwnd)?;
        let g1 = get_idesktop_guid(&desktop);
        let g2 = get_idesktop_guid(&desktop2);
        Ok(g1 == g2)
    }

    pub fn get_desktop_number_by_window(hwnd: HWND) -> Result<u32> {
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager(&provider)?;
        let man2 = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop2 = get_idesktop_by_window(&man2, &manager, hwnd)?;
        get_idesktop_number(&man2, &desktop2)
    }

    /// Is window pinned?
    pub fn is_pinned_window(hwnd: HWND) -> Result<bool> {
        com_sta();
        let provider = get_iservice_provider()?;
        let view_collection = get_iapplication_view_collection(&provider)?;
        let apps = get_ivirtual_desktop_pinned_apps(&provider)?;
        let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
        is_view_pinned(&apps, view)
    }

    /// Pin window
    pub fn pin_window(hwnd: HWND) -> Result<()> {
        com_sta();
        let provider = get_iservice_provider()?;
        let view_collection = get_iapplication_view_collection(&provider)?;
        let apps = get_ivirtual_desktop_pinned_apps(&provider)?;
        let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
        pin_view(&apps, view)
    }

    /// Unpin window
    pub fn unpin_window(hwnd: HWND) -> Result<()> {
        com_sta();
        let provider = get_iservice_provider()?;
        let view_collection = get_iapplication_view_collection(&provider)?;
        let apps = get_ivirtual_desktop_pinned_apps(&provider)?;
        let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
        upin_view(&apps, view)
    }

    /// Is pinned app
    pub fn is_pinned_app(hwnd: HWND) -> Result<bool> {
        com_sta();
        let provider = get_iservice_provider()?;
        let view_collection = get_iapplication_view_collection(&provider)?;
        let apps = get_ivirtual_desktop_pinned_apps(&provider)?;
        let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
        let app_id = get_iapplication_id_for_view(&view)?;
        is_app_id_pinned(&apps, app_id)
    }

    /// Pin app
    pub fn pin_app(hwnd: HWND) -> Result<()> {
        com_sta();
        let provider = get_iservice_provider()?;
        let view_collection = get_iapplication_view_collection(&provider)?;
        let apps = get_ivirtual_desktop_pinned_apps(&provider)?;
        let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
        let app_id = get_iapplication_id_for_view(&view)?;
        pin_app_id(&apps, app_id)
    }

    /// Unpin app
    pub fn unpin_app(hwnd: HWND) -> Result<()> {
        com_sta();
        let provider = get_iservice_provider()?;
        let view_collection = get_iapplication_view_collection(&provider)?;
        let apps = get_ivirtual_desktop_pinned_apps(&provider)?;
        let view = get_iapplication_view_for_hwnd(&view_collection, hwnd)?;
        let app_id = get_iapplication_id_for_view(&view)?;
        unpin_app_id(&apps, app_id)
    }
}

pub mod normal {
    use super::*;
    use std::fmt::Debug;
    use windows::core::GUID;

    #[derive(Copy, Clone, PartialEq)]
    pub struct Desktop(GUID);

    impl Debug for Desktop {
        fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            write!(f, "Desktop({:?})", self.0)
        }
    }

    impl Desktop {
        pub(crate) fn empty() -> Desktop {
            Desktop(GUID::default())
        }

        pub fn get_id(&self) -> GUID {
            self.0
        }

        pub fn get_name(&self) -> Result<String> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager = get_ivirtual_desktop_manager_internal(&provider)?;
            let desktop = get_idesktop_by_guid(&manager, &self.get_id())?;
            get_idesktop_name(&desktop)
        }

        pub fn set_name(&self, name: &str) -> Result<()> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager = get_ivirtual_desktop_manager_internal(&provider)?;
            let idesktop = get_idesktop_by_guid(&manager, &self.get_id())?;
            set_idesktop_name(&manager, &idesktop, name)
        }

        pub fn get_index(&self) -> Result<u32> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager = get_ivirtual_desktop_manager_internal(&provider)?;
            let idesktop = get_idesktop_by_guid(&manager, &self.get_id())?;
            let index = get_idesktop_number(&manager, &idesktop)?;
            Ok(index)
        }

        pub fn get_wallpaper(&self) -> Result<String> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager = get_ivirtual_desktop_manager_internal(&provider)?;
            let idesktop = get_idesktop_by_guid(&manager, &self.get_id())?;
            get_idesktop_wallpaper(&idesktop)
        }

        pub fn set_wallpaper(&self, path: &str) -> Result<()> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager = get_ivirtual_desktop_manager_internal(&provider)?;
            let idesktop = get_idesktop_by_guid(&manager, &self.get_id())?;
            set_idesktop_wallpaper(&manager, &idesktop, path)
        }

        pub fn switch_to(&self) -> Result<()> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager = get_ivirtual_desktop_manager_internal(&provider)?;
            let idesktop = get_idesktop_by_guid(&manager, &self.get_id())?;
            switch_to_idesktop(&manager, &idesktop)
        }

        pub fn remove(&self, fallback_desktop: &Desktop) -> Result<()> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager = get_ivirtual_desktop_manager_internal(&provider)?;
            let idesktop = get_idesktop_by_guid(&manager, &self.get_id())?;
            let fallback_idesktop = get_idesktop_by_guid(&manager, &fallback_desktop.0)?;
            remove_idesktop(&manager, &idesktop, &fallback_idesktop)
        }

        pub fn has_window(&self, hwnd: HWND_) -> Result<bool> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager_internal = get_ivirtual_desktop_manager_internal(&provider)?;
            let manager = get_ivirtual_desktop_manager(&provider)?;
            let desktop = get_idesktop_by_window(&manager_internal, &manager, hwnd)?;
            let desktop_id = get_idesktop_guid(&desktop)?;
            Ok(desktop_id == self.get_id())
        }

        pub fn move_window(&self, hwnd: HWND_) -> Result<()> {
            com_sta();
            let provider = get_iservice_provider()?;
            let manager = get_ivirtual_desktop_manager_internal(&provider)?;
            let vc = get_iapplication_view_collection(&provider)?;
            let view = get_iapplication_view_for_hwnd(&vc, hwnd)?;
            let idesktop = get_idesktop_by_guid(&manager, &self.get_id())?;
            move_view_to_desktop(&manager, &view, &idesktop)
        }
    }

    pub fn create_desktop() -> Result<Desktop> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = create_idesktop(&manager)?;
        let id = get_idesktop_guid(&desktop)?;
        Ok(Desktop(id))
    }

    pub fn get_current_desktop() -> Result<Desktop> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = get_current_idesktop(&manager)?;
        let id = get_idesktop_guid(&desktop)?;
        Ok(Desktop(id))
    }

    pub fn get_desktop_count() -> Result<u32> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktops = get_idesktops_array(&manager)?;
        unsafe { desktops.GetCount().map_err(map_win_err) }
    }
    pub fn get_desktops() -> Result<Vec<Desktop>> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktops: Result<Vec<Desktop>> = get_idesktops(&manager)?
            .into_iter()
            .map(|d| -> Result<Desktop> {
                let mut desktop = Desktop::empty();
                unsafe { d.get_id(&mut desktop.0).as_result()? };
                Ok(desktop)
            })
            .collect();
        Ok(desktops?)
    }

    pub fn get_desktop_by_index(index: u32) -> Result<Desktop> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = get_idesktop_by_number(&manager, index)?;
        let id = get_idesktop_guid(&desktop)?;
        Ok(Desktop(id))
    }

    pub fn get_desktop_by_guid(guid: &GUID) -> Result<Desktop> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = get_idesktop_by_guid(&manager, &guid)?;
        let id = get_idesktop_guid(&desktop)?;
        Ok(Desktop(id))
    }

    pub fn get_desktop_by_window(hwnd: HWND_) -> Result<Desktop> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager_internal = get_ivirtual_desktop_manager_internal(&provider)?;
        let manager = get_ivirtual_desktop_manager(&provider)?;
        let desktop = get_idesktop_by_window(&manager_internal, &manager, hwnd)?;
        let id = get_idesktop_guid(&desktop)?;
        Ok(Desktop(id))
    }
}

pub mod numbered {
    use super::*;

    pub fn go_to_desktop_number(number: u32) -> Result<()> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = get_idesktop_by_number(&manager, number)?;
        switch_to_idesktop(&manager, &desktop)
    }

    pub fn get_current_desktop_number() -> Result<u32> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktops = get_idesktops(&manager)?;
        let current = get_current_idesktop(&manager)?;
        for (i, desktop) in desktops.iter().enumerate() {
            if desktop == &current {
                return Ok(i as u32);
            }
        }
        Err(Error::DesktopNotFound)
    }

    pub fn set_name_by_desktop_number(number: u32, name: &str) -> Result<()> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = get_idesktop_by_number(&manager, number)?;
        let name = HSTRING::from(name);
        unsafe { manager.set_name(ComIn::new(&desktop), name).as_result() }
    }

    pub fn get_name_by_desktop_number(number: u32) -> Result<String> {
        com_sta();
        let provider = get_iservice_provider()?;
        let manager = get_ivirtual_desktop_manager_internal(&provider)?;
        let desktop = get_idesktop_by_number(&manager, number)?;
        let mut name = HSTRING::new();
        unsafe { desktop.get_name(&mut name).as_result()? }
        Ok(name.to_string())
    }
}

pub mod listener {}

mod tests {
    use std::{
        pin::Pin,
        rc::Rc,
        sync::{mpsc::Sender, Mutex},
        time::Duration,
    };

    use windows::Win32::{
        Foundation::HWND,
        System::Threading::{CreateThread, WaitForSingleObject, THREAD_CREATION_FLAGS},
        UI::WindowsAndMessaging::{
            DispatchMessageW, GetMessageW, PostQuitMessage, PostThreadMessageW, TranslateMessage,
            MSG, WM_QUIT, WM_USER,
        },
    };

    use super::normal::*;
    use super::numbered::*;
    use super::*;

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
        com_sta();
        let mut gid = GUID::default();
        unsafe { desktop_new.get_id(&mut gid).panic_if_failed() };

        let mut name = HSTRING::new();
        unsafe { desktop_new.get_name(&mut name).panic_if_failed() };

        let manager =
            get_ivirtual_desktop_manager_internal(&get_iservice_provider().unwrap()).unwrap();
        let number = get_idesktop_number(&manager, &desktop_new).unwrap_or(99999);

        println!(
            "{}: {} {:?} {:?} {:?}",
            prefix,
            number,
            gid,
            name.to_string(),
            std::thread::current().id()
        );
    }

    struct TestVDNotificationsWrapper {
        ptr: IVirtualDesktopNotification,
    }

    impl TestVDNotificationsWrapper {
        pub fn new(
            number_times_desktop_changed: Sender<()>,
        ) -> Result<Box<TestVDNotificationsWrapper>> {
            println!("CREATED IN THREAD {:?}", std::thread::current().id());
            com_sta();
            let provider = get_iservice_provider()?;
            let service = get_ivirtual_desktop_notification_service(&provider)?;
            let ptr = TestVDNotifications {};
            let notification = Box::new(TestVDNotificationsWrapper {
                // cookie: Rc::new(Mutex::new(0)),
                // number_times_desktop_changed: Rc::new(number_times_desktop_changed),
                ptr: ptr.into(), // service: service.clone(),
            });

            let mut cookie = 0;
            unsafe {
                service
                    .register(notification.ptr.as_raw(), &mut cookie)
                    .panic_if_failed();
                assert_ne!(cookie, 0);
            }
            // *notification.cookie.lock().unwrap() = cookie;
            println!(
                "Registered notification {} {:?}",
                cookie,
                std::thread::current().id()
            );

            Ok(notification)
        }
    }

    #[derive(Clone)]
    #[windows::core::implement(IVirtualDesktopNotification)]
    struct TestVDNotifications {
        // cookie: Rc<Mutex<u32>>,
        // number_times_desktop_changed: Rc<Sender<()>>,
        // service: IVirtualDesktopNotificationService,
    }

    impl Drop for TestVDNotifications {
        fn drop(&mut self) {
            // let provider = get_iservice_provider().unwrap();
            // let service = get_ivirtual_desktop_notification_service(&provider).unwrap();
            // let cookie = *self.cookie.lock().unwrap();
            println!("Drop notification");
            // unsafe { service.unregister(cookie).panic_if_failed() };
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
            // NOTE: This is a bit redundant, but I know that Windows calls our listener thread with COINIT_MULTITHREADED, so this is a note.
            com_mta();
            // println!("Changed desktop {:?}", std::thread::current().id());
            debug_desktop(&desktop_new, "Desktop changed");
            // self.number_times_desktop_changed.send(()).unwrap();
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
        com_sta();

        let (tx, rx) = std::sync::mpsc::channel();
        let notification_thread = std::thread::spawn(|| {
            com_sta();
            println!("Notification thread {:?}", std::thread::current().id());
            let _notification = TestVDNotificationsWrapper::new(tx).unwrap();
            let mut msg = MSG::default();
            unsafe {
                while GetMessageW(&mut msg, HWND::default(), 0, 0).as_bool() {
                    if (msg.message == WM_USER + 0x10) {
                        PostQuitMessage(0);
                    }
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
            }
        });

        let current_desktop = get_current_desktop_number().unwrap();

        for _ in 0..999 {
            go_to_desktop_number(0).unwrap();
            // std::thread::sleep(Duration::from_millis(4));
            go_to_desktop_number(1).unwrap();
        }
        std::thread::sleep(Duration::from_millis(3));
        go_to_desktop_number(current_desktop).unwrap();
    }

    #[test]
    fn test_initialize() {

        /*

        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap();
        }
        println!("CoInitializeEx COINIT_APARTMENTTHREADED");
        std::thread::spawn(|| unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED).unwrap();
            println!("CoInitializeEx COINIT_MULTITHREADED");
            CoUninitialize();
            println!("CoUninitialize COINIT_MULTITHREADED");
        })
        .join()
        .unwrap();

        unsafe {
            CoUninitialize();
        }
        println!("CoUninitialize COINIT_APARTMENTTHREADED");
         */
    }

    #[test] // TODO: Commented out, use only on occasion when needed!
    fn test_listener_manual() {
        println!("Main thread is {:?}", std::thread::current().id());

        let (tx, rx) = std::sync::mpsc::channel();
        std::thread::spawn(|| {
            println!("Notification thread {:?}", std::thread::current().id());
            let _notification = TestVDNotificationsWrapper::new(tx).unwrap();
            let mut msg = MSG::default();
            unsafe {
                while (GetMessageW(&mut msg, HWND::default(), 0, 0).as_bool()) {
                    TranslateMessage(&msg);
                    DispatchMessageW(&msg);
                }
            }
        })
        .join()
        .unwrap();
        // while sleep

        // std::thread::sleep(Duration::from_secs(1));
        // // go_to_desktop_number(0).unwrap();
        // // std::thread::sleep(Duration::from_millis(4));
        // // go_to_desktop_number(2).unwrap();
        // std::thread::sleep(Duration::from_secs(1));
        // std::thread::sleep(Duration::from_secs(1));
        // std::thread::sleep(Duration::from_secs(1));
        // std::thread::sleep(Duration::from_secs(1));
        // std::thread::sleep(Duration::from_secs(1));
        // std::thread::sleep(Duration::from_secs(1));
    }

    /// This test switched desktop and prints out the changed desktop
    #[test]
    fn test_register_notifications() {
        let (tx, rx) = std::sync::mpsc::channel();
        let notification = TestVDNotificationsWrapper::new(tx);

        let provider = get_iservice_provider().unwrap();
        let service = get_ivirtual_desktop_notification_service(&provider).unwrap();
        let manager = get_ivirtual_desktop_manager_internal(&provider).unwrap();

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

        let provider = get_iservice_provider().unwrap();
        let manager: IVirtualDesktopManagerInternal =
            get_ivirtual_desktop_manager_internal(&provider).unwrap();

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
