use super::{
    interfaces::{ComIn, IVirtualDesktop, IVirtualDesktopManagerInternal},
    *,
};
use std::fmt::Debug;
use windows::{
    core::{GUID, HSTRING},
    Win32::Foundation::HWND,
};

use super::raw::*;

#[derive(Copy, Clone, PartialEq, Debug)]
enum DesktopInternal {
    Index(u32),
    Guid(GUID),
    IndexGuid(u32, GUID),
}

/// You can construct Desktop instance with `get_desktop` by index or GUID.
#[derive(Copy, Clone, PartialEq, Debug)]
pub struct Desktop(DesktopInternal);

// Impl from u32 to DesktopTest
impl From<u32> for Desktop {
    fn from(index: u32) -> Self {
        Desktop(DesktopInternal::Index(index))
    }
}

// Impl from i32 to DesktopTest
impl From<i32> for Desktop {
    fn from(index: i32) -> Self {
        Desktop(DesktopInternal::Index(index as u32))
    }
}

// Impl from GUID to DesktopTest
impl From<GUID> for Desktop {
    fn from(guid: GUID) -> Self {
        Desktop(DesktopInternal::Guid(guid))
    }
}

// Impl from &GUID to DesktopTest
impl From<&GUID> for Desktop {
    fn from(guid: &GUID) -> Self {
        Desktop(DesktopInternal::Guid(*guid))
    }
}
impl Desktop {
    fn get_ivirtual_desktop(
        &self,
        manager: &IVirtualDesktopManagerInternal,
    ) -> Result<IVirtualDesktop> {
        match &self.0 {
            DesktopInternal::Index(index) => get_idesktop_by_index(manager, *index),
            DesktopInternal::Guid(guid) => get_idesktop_by_guid(manager, guid),
            DesktopInternal::IndexGuid(_, guid) => get_idesktop_by_guid(manager, guid),
        }
    }

    fn internal_get_id(&self, manager: &IVirtualDesktopManagerInternal) -> Result<GUID> {
        match &self.0 {
            DesktopInternal::Index(index) => {
                com_sta();
                let idesktop = get_idesktop_by_index(manager, *index)?;
                get_idesktop_guid(&idesktop)
            }
            DesktopInternal::Guid(guid) => Ok(*guid),
            DesktopInternal::IndexGuid(_, guid) => Ok(*guid),
        }
    }

    /// Get the GUID of the desktop
    pub fn get_id(&self) -> Result<GUID> {
        match &self.0 {
            DesktopInternal::Index(index) => {
                com_sta();
                let manager = get_ivirtual_desktop_manager_internal_noparams()?;
                let idesktop = get_idesktop_by_index(&manager, *index)?;
                get_idesktop_guid(&idesktop)
            }
            DesktopInternal::Guid(guid) => Ok(*guid),
            DesktopInternal::IndexGuid(_, guid) => Ok(*guid),
        }
    }

    pub fn get_index(&self) -> Result<u32> {
        match &self.0 {
            DesktopInternal::Index(index) => Ok(*index),
            DesktopInternal::Guid(guid) => {
                com_sta();
                let manager = get_ivirtual_desktop_manager_internal_noparams()?;
                let idesktop = get_idesktop_by_guid(&manager, guid)?;
                get_idesktop_index(&manager, &idesktop)
            }
            DesktopInternal::IndexGuid(index, _) => Ok(*index),
        }
    }

    /// Get desktop name
    pub fn get_name(&self) -> Result<String> {
        com_sta();
        let manager = get_ivirtual_desktop_manager_internal_noparams()?;
        let idesk = self.get_ivirtual_desktop(&manager);
        get_idesktop_name(&idesk?)
    }

    /// Set desktop name
    pub fn set_name(&self, name: &str) -> Result<()> {
        com_sta();
        let manager = get_ivirtual_desktop_manager_internal_noparams()?;
        let idesk = self.get_ivirtual_desktop(&manager);
        set_idesktop_name(&manager, &idesk?, name)
    }

    /// Get desktop wallpaper path
    pub fn get_wallpaper(&self) -> Result<String> {
        com_sta();
        let manager = get_ivirtual_desktop_manager_internal_noparams()?;
        let idesk = self.get_ivirtual_desktop(&manager);
        get_idesktop_wallpaper(&idesk?)
    }

    /// Set desktop wallpaper path
    pub fn set_wallpaper(&self, path: &str) -> Result<()> {
        com_sta();
        let manager = get_ivirtual_desktop_manager_internal_noparams()?;
        let idesk = self.get_ivirtual_desktop(&manager);
        set_idesktop_wallpaper(&manager, &idesk?, path)
    }
}

/// Get desktop by index or GUID
///
/// # Examples
/// * Get first desktop by index `get_desktop(0)`
/// * Get second desktop by index `get_desktop(1)`
/// * Get desktop by GUID `get_desktop(GUID(0, 0, 0, [0, 0, 0, 0, 0, 0, 0, 0]))`
///
/// Note: This function does not check if the desktop exists.
pub fn get_desktop<T>(desktop: T) -> Desktop
where
    T: Into<Desktop>,
{
    desktop.into()
}

/// Test if desktop exists
// pub fn desktop_exists<T>(desktop: T) -> Result<bool>
// where
//     T: Into<Desktop>,
// {
//     com_sta();
//     let provider = get_iservice_provider()?;
//     let manager = get_ivirtual_desktop_manager_internal(&provider)?;
//     let idesktop = desktop.into().get_ivirtual_desktop(&manager)?;
//     get_idesktop_guid(&idesktop).map(|_| true)
// }

/// Switch desktop by index or GUID
pub fn switch_desktop<T>(desktop: T) -> Result<()>
where
    T: Into<Desktop> + Copy + Clone + Send + 'static,
{
    // com_sta_thread(move || {
    com_sta();
    let manager = get_ivirtual_desktop_manager_internal_noparams()?;
    let idesktop = desktop.into().get_ivirtual_desktop(&manager)?;
    switch_to_idesktop(&manager, &idesktop)
    // })
}

/// Remove desktop by index or GUID
pub fn remove_desktop<T, F>(desktop: T, fallback_desktop: F) -> Result<()>
where
    T: Into<Desktop>,
    F: Into<Desktop>,
{
    com_sta();
    let manager = get_ivirtual_desktop_manager_internal_noparams()?;
    let idesktop = desktop.into().get_ivirtual_desktop(&manager)?;
    let fallback_idesktop = fallback_desktop.into().get_ivirtual_desktop(&manager)?;
    remove_idesktop(&manager, &idesktop, &fallback_idesktop)
}

/// Is window on desktop by index or GUID
pub fn is_window_on_desktop<T>(desktop: T, hwnd: HWND) -> Result<bool>
where
    T: Into<Desktop>,
{
    com_sta();
    let provider = get_iservice_provider()?;
    let manager_internal = get_ivirtual_desktop_manager_internal_for_provider(&provider)?;
    let manager = get_ivirtual_desktop_manager(&provider)?;

    // Get desktop of the window
    let desktop_win = get_idesktop_by_window(&manager_internal, &manager, hwnd)?;
    let desktop_win_id = get_idesktop_guid(&desktop_win)?;

    // If ID matches with given desktop, return true
    Ok(desktop_win_id == desktop.into().internal_get_id(&manager_internal)?)
}

/// Move window to desktop by index or GUID
pub fn move_window_to_desktop<T>(desktop: T, hwnd: HWND) -> Result<()>
where
    T: Into<Desktop>,
{
    com_sta();
    let provider = get_iservice_provider()?;
    let manager = get_ivirtual_desktop_manager_internal_for_provider(&provider)?;
    let vc = get_iapplication_view_collection(&provider)?;
    let view = get_iapplication_view_for_hwnd(&vc, hwnd)?;
    let idesktop = desktop.into().get_ivirtual_desktop(&manager)?;
    move_view_to_desktop(&manager, &view, &idesktop)
}

/// Create desktop
pub fn create_desktop() -> Result<Desktop> {
    com_sta();
    let manager = get_ivirtual_desktop_manager_internal_noparams()?;
    let desktop = create_idesktop(&manager)?;
    let id = get_idesktop_guid(&desktop)?;
    Ok(Desktop(DesktopInternal::Guid(id)))
}

/// Get current desktop
pub fn get_current_desktop() -> Result<Desktop> {
    com_sta();
    let manager = get_ivirtual_desktop_manager_internal_noparams()?;
    let desktop = get_current_idesktop(&manager)?;
    let id = get_idesktop_guid(&desktop)?;
    Ok(Desktop(DesktopInternal::Guid(id)))
}

/// Get all desktops
pub fn get_desktops() -> Result<Vec<Desktop>> {
    com_sta();
    let manager = get_ivirtual_desktop_manager_internal_noparams()?;
    get_idesktops(&manager)?
        .into_iter()
        .enumerate()
        .map(|(i, d)| -> Result<Desktop> {
            let mut guid = GUID::default();
            unsafe { d.get_id(&mut guid).as_result()? };
            Ok(Desktop(DesktopInternal::IndexGuid(i as u32, guid)))
        })
        .collect()
}

/// Get desktop by window
pub fn get_desktop_by_window(hwnd: HWND) -> Result<Desktop> {
    com_sta();
    let provider = get_iservice_provider()?;
    let manager_internal = get_ivirtual_desktop_manager_internal_for_provider(&provider)?;
    let manager = get_ivirtual_desktop_manager(&provider)?;
    let desktop = get_idesktop_by_window(&manager_internal, &manager, hwnd)?;
    let id = get_idesktop_guid(&desktop)?;
    Ok(Desktop(DesktopInternal::Guid(id)))
}

/// Get desktop count
pub fn get_desktop_count() -> Result<u32> {
    com_sta();

    let manager = get_ivirtual_desktop_manager_internal_noparams()?;
    let desktops = get_idesktops_array(&manager)?;
    unsafe { desktops.GetCount().map_err(map_win_err) }
}

pub fn is_window_on_current_desktop(hwnd: HWND) -> Result<bool> {
    com_sta();
    let provider = get_iservice_provider()?;
    let manager = get_ivirtual_desktop_manager(&provider)?;
    _is_window_on_current_desktop(&manager, hwnd)
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
