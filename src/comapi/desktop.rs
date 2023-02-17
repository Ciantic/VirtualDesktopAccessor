use super::{
    interfaces::{ComIn, IVirtualDesktop, IVirtualDesktopManagerInternal},
    *,
};
use std::{convert::TryFrom, fmt::Debug, rc::Rc};
use windows::{
    core::{GUID, HSTRING},
    Win32::Foundation::HWND,
};

use super::raw2::*;

/// You can construct Desktop instance with `get_desktop` by index or GUID.
#[derive(Copy, Clone, Debug)]
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

impl From<DesktopInternal> for Desktop {
    fn from(desktop: DesktopInternal) -> Self {
        Desktop(desktop)
    }
}

impl From<Desktop> for DesktopInternal {
    fn from(desktop: Desktop) -> Self {
        desktop.0
    }
}

impl<'a> TryFrom<ComIn<'a, IVirtualDesktop>> for Desktop {
    type Error = Error;

    fn try_from(desktop: ComIn<IVirtualDesktop>) -> Result<Self> {
        Ok(Desktop(DesktopInternal::try_from(desktop)?))
    }
}
impl Desktop {
    pub fn try_eq(&self, other: &Desktop) -> Result<bool> {
        self.0.try_eq(&other.0)
    }

    /// Get the GUID of the desktop
    pub fn get_id(&self) -> Result<GUID> {
        com_objects().get_desktop_id(&self.0)
    }

    pub fn get_index(&self) -> Result<u32> {
        com_objects().get_desktop_index(&self.0)
    }

    /// Get desktop name
    pub fn get_name(&self) -> Result<String> {
        com_objects().get_desktop_name(&self.0)
    }

    /// Set desktop name
    pub fn set_name(&self, name: &str) -> Result<()> {
        com_objects().set_desktop_name(&self.0, name)
    }

    /// Get desktop wallpaper path
    pub fn get_wallpaper(&self) -> Result<String> {
        com_objects().get_desktop_wallpaper(&self.0)
    }

    /// Set desktop wallpaper path
    pub fn set_wallpaper(&self, path: &str) -> Result<()> {
        com_objects().set_desktop_wallpaper(&self.0, path)
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

/// Switch desktop by index or GUID
pub fn switch_desktop<T>(desktop: T) -> Result<()>
where
    T: Into<Desktop> + Copy + Clone + Send + 'static,
{
    com_objects().switch_desktop(&desktop.into().into())
}

/// Remove desktop by index or GUID
pub fn remove_desktop<T, F>(desktop: T, fallback_desktop: F) -> Result<()>
where
    T: Into<Desktop>,
    F: Into<Desktop>,
{
    com_objects().remove_desktop(&desktop.into().into(), &fallback_desktop.into().into())
}

/// Is window on desktop by index or GUID
pub fn is_window_on_desktop<T>(desktop: T, hwnd: HWND) -> Result<bool>
where
    T: Into<Desktop>,
{
    com_objects().is_window_on_desktop(&hwnd, &desktop.into().into())
}

/// Move window to desktop by index or GUID
pub fn move_window_to_desktop<T>(desktop: T, hwnd: &HWND) -> Result<()>
where
    T: Into<Desktop>,
{
    Ok(com_objects().move_window_to_desktop(hwnd, &desktop.into().into())?)
}

/// Create desktop
pub fn create_desktop() -> Result<Desktop> {
    Ok(com_objects().create_desktop()?.into())
}

/// Get current desktop
pub fn get_current_desktop() -> Result<Desktop> {
    Ok(com_objects().get_current_desktop()?.into())
}

/// Get all desktops
pub fn get_desktops() -> Result<Vec<Desktop>> {
    Ok(com_objects()
        .get_desktops()?
        .into_iter()
        .map(Desktop)
        .collect())
}

/// Get desktop by window
pub fn get_desktop_by_window(hwnd: HWND) -> Result<Desktop> {
    com_objects().get_desktop_by_window(&hwnd).map(Desktop)
}

/// Get desktop count
pub fn get_desktop_count() -> Result<u32> {
    com_objects().get_desktop_count()
}

pub fn is_window_on_current_desktop(hwnd: HWND) -> Result<bool> {
    com_objects().is_window_on_current_desktop(&hwnd)
}

/// Is window pinned?
pub fn is_pinned_window(hwnd: HWND) -> Result<bool> {
    com_objects().is_pinned_window(&hwnd)
}

/// Pin window
pub fn pin_window(hwnd: HWND) -> Result<()> {
    com_objects().pin_window(&hwnd)
}

/// Unpin window
pub fn unpin_window(hwnd: HWND) -> Result<()> {
    com_objects().unpin_window(&hwnd)
}

/// Is pinned app
pub fn is_pinned_app(hwnd: HWND) -> Result<bool> {
    com_objects().is_pinned_app(&hwnd)
}

/// Pin app
pub fn pin_app(hwnd: HWND) -> Result<()> {
    com_objects().pin_app(&hwnd)
}

/// Unpin app
pub fn unpin_app(hwnd: HWND) -> Result<()> {
    com_objects().unpin_app(&hwnd)
}
