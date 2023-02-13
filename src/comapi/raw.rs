/// Purpose of this module is to provide helpers to access functions in interfaces module, not for direct consumption
///
/// All functions here either take in a reference to an interface or initializes a com interace.
use super::interfaces::*;
use super::Result;
use crate::{Error, HRESULT};
use std::{cell::RefCell, ffi::c_void};
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

type APPIDPWSTR = *mut *mut std::ffi::c_void;

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
            let _ = CoInitializeEx(None, dwcoinit);
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

pub fn map_win_err(er: ::windows::core::Error) -> Error {
    Error::ComError(HRESULT::from_i32(er.code().0))
}

pub fn get_iservice_provider() -> Result<IServiceProvider> {
    com_sta();
    unsafe { CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_ALL).map_err(map_win_err) }
}

pub fn get_ivirtual_desktop_notification_service(
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

pub fn get_ivirtual_desktop_manager(provider: &IServiceProvider) -> Result<IVirtualDesktopManager> {
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

pub fn get_ivirtual_desktop_manager_internal(
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

pub fn get_iapplication_view_collection(
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

pub fn get_ivirtual_desktop_pinned_apps(
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

pub fn get_idesktops_array(manager: &IVirtualDesktopManagerInternal) -> Result<IObjectArray> {
    com_sta();
    let mut desktops = None;
    unsafe { manager.get_desktops(0, &mut desktops).as_result()? }
    Ok(desktops.unwrap())
}

pub fn get_idesktop_number(
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

pub fn get_idesktop_wallpaper(desktop: &IVirtualDesktop) -> Result<String> {
    com_sta();
    let mut name = HSTRING::default();
    unsafe { desktop.get_wallpaper(&mut name).as_result()? }
    Ok(name.to_string_lossy())
}

pub fn set_idesktop_wallpaper(
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

pub fn get_idesktop_guid(desktop: &IVirtualDesktop) -> Result<GUID> {
    com_sta();
    let mut guid = GUID::default();
    unsafe { desktop.get_id(&mut guid).as_result()? }
    Ok(guid)
}

pub fn get_idesktop_name(desktop: &IVirtualDesktop) -> Result<String> {
    com_sta();
    let mut name = HSTRING::default();
    unsafe { desktop.get_name(&mut name).as_result()? }
    Ok(name.to_string_lossy())
}

pub fn set_idesktop_name(
    manager: &IVirtualDesktopManagerInternal,
    desktop: &IVirtualDesktop,
    name: &str,
) -> Result<()> {
    com_sta();
    let name = HSTRING::from(name);
    unsafe { manager.set_name(ComIn::new(&desktop), name).as_result()? }
    Ok(())
}

pub fn get_idesktop_by_number(
    manager: &IVirtualDesktopManagerInternal,
    index: u32,
) -> Result<IVirtualDesktop> {
    com_sta();
    let desktops = get_idesktops_array(manager)?;
    let desktop = unsafe { desktops.GetAt(index).map_err(map_win_err) };
    desktop.map_err(|_| Error::DesktopNotFound)
}

pub fn get_idesktop_by_guid(
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

pub fn get_idesktops(manager: &IVirtualDesktopManagerInternal) -> Result<Vec<IVirtualDesktop>> {
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

pub fn get_current_idesktop(manager: &IVirtualDesktopManagerInternal) -> Result<IVirtualDesktop> {
    com_sta();
    let mut desktop = None;
    unsafe { manager.get_current_desktop(0, &mut desktop).as_result()? }
    desktop.ok_or(Error::DesktopNotFound)
}

pub fn switch_to_idesktop(
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

pub fn create_idesktop(manager: &IVirtualDesktopManagerInternal) -> Result<IVirtualDesktop> {
    com_sta();
    let mut desktop = None;
    unsafe { manager.create_desktop(0, &mut desktop).as_result()? }
    desktop.ok_or(Error::CreateDesktopFailed)
}

pub fn move_view_to_desktop(
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

pub fn remove_idesktop(
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

pub fn get_iapplication_id_for_view(view: &IApplicationView) -> Result<APPIDPWSTR> {
    com_sta();
    let mut app_id: APPIDPWSTR = std::ptr::null_mut();
    unsafe {
        view.get_app_user_model_id(&mut app_id as *mut _ as *mut _)
            .as_result()?
    }
    Ok(app_id)
}

pub fn get_iapplication_view_for_hwnd(
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

pub fn is_view_pinned(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<bool> {
    com_sta();
    let mut is_pinned = false;
    unsafe {
        apps.is_view_pinned(ComIn::new(&view), &mut is_pinned)
            .as_result()?
    }
    Ok(is_pinned)
}

pub fn pin_view(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<()> {
    com_sta();
    unsafe { apps.pin_view(ComIn::new(&view)).as_result() }
}

pub fn upin_view(apps: &IVirtualDesktopPinnedApps, view: IApplicationView) -> Result<()> {
    com_sta();
    unsafe { apps.unpin_view(ComIn::new(&view)).as_result() }
}

pub fn is_app_id_pinned(apps: &IVirtualDesktopPinnedApps, app_id: APPIDPWSTR) -> Result<bool> {
    com_sta();
    let mut is_pinned = false;
    unsafe {
        apps.is_app_pinned(app_id as *mut _, &mut is_pinned)
            .as_result()?
    }
    Ok(is_pinned)
}

pub fn pin_app_id(apps: &IVirtualDesktopPinnedApps, app_id: APPIDPWSTR) -> Result<()> {
    com_sta();
    unsafe { apps.pin_app(app_id as *mut _).as_result() }
}

pub fn unpin_app_id(apps: &IVirtualDesktopPinnedApps, app_id: APPIDPWSTR) -> Result<()> {
    com_sta();
    unsafe { apps.unpin_app(app_id as *mut _).as_result() }
}

pub fn _is_window_on_current_desktop(
    manager: &IVirtualDesktopManager,
    hwnd: HWND_,
) -> Result<bool> {
    com_sta();
    let mut is_on_desktop = false;
    unsafe {
        manager
            .is_window_on_current_desktop(hwnd, &mut is_on_desktop)
            .as_result()?
    }
    Ok(is_on_desktop)
}

pub fn get_idesktop_by_window(
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
