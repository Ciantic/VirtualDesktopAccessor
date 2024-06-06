use super::comobjects::*;
use super::interfaces_multi::{ComIn, IVirtualDesktop};
use super::*;
use std::{convert::TryFrom, fmt::Debug};
use windows::{core::GUID, Win32::Foundation::HWND};

/// You can construct Desktop instance with `get_desktop(5)` by index or GUID.
#[derive(Copy, Clone, Debug)]
pub struct Desktop(DesktopInternal);

impl Eq for Desktop {}

impl PartialEq for Desktop {
    fn eq(&self, other: &Self) -> bool {
        let a = self.0;
        let b = other.0;
        match (&a, &b) {
            (DesktopInternal::Index(a), DesktopInternal::Index(b)) => a == b,
            (DesktopInternal::Guid(a), DesktopInternal::Guid(b)) => a == b,
            (DesktopInternal::IndexGuid(a, b), DesktopInternal::IndexGuid(c, d)) => {
                a == c && b == d
            }
            (DesktopInternal::Index(a), DesktopInternal::IndexGuid(b, _)) => a == b,
            (DesktopInternal::IndexGuid(a, _), DesktopInternal::Index(b)) => a == b,
            (DesktopInternal::Guid(a), DesktopInternal::IndexGuid(_, b)) => a == b,
            (DesktopInternal::IndexGuid(_, a), DesktopInternal::Guid(b)) => a == b,
            _ => with_com_objects(move |f| Ok(f.get_desktop_id(&a)? == f.get_desktop_id(&b)?))
                .unwrap_or(false),
        }
    }
}

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

impl TryFrom<IVirtualDesktop> for Desktop {
    type Error = Error;

    fn try_from(desktop: IVirtualDesktop) -> Result<Self> {
        Ok(Desktop(DesktopInternal::try_from(&desktop)?))
    }
}
impl<'a> TryFrom<&'a ComIn<'a, IVirtualDesktop>> for Desktop {
    type Error = Error;

    fn try_from(desktop: &'a ComIn<'a, IVirtualDesktop>) -> Result<Self> {
        Ok(Desktop(DesktopInternal::try_from(desktop)?))
    }
}
impl<'a> TryFrom<ComIn<'a, IVirtualDesktop>> for Desktop {
    type Error = Error;

    fn try_from(desktop: ComIn<'a, IVirtualDesktop>) -> Result<Self> {
        Ok(Desktop(DesktopInternal::try_from(&desktop)?))
    }
}
impl Desktop {
    /// Get the GUID of the desktop
    pub fn get_id(&self) -> Result<GUID> {
        let internal = self.0;
        with_com_objects(move |o| o.get_desktop_id(&internal))
    }

    pub fn get_index(&self) -> Result<u32> {
        let internal = self.0;
        with_com_objects(move |o| o.get_desktop_index(&internal))
    }

    /// Get desktop name
    pub fn get_name(&self) -> Result<String> {
        let internal = self.0;
        with_com_objects(move |o| o.get_desktop_name(&internal))
    }

    /// Set desktop name
    pub fn set_name(&self, name: &str) -> Result<()> {
        let internal = self.0;
        let name_ = name.to_owned();
        with_com_objects(move |o| o.set_desktop_name(&internal, &name_))
    }

    /// Get desktop wallpaper path
    pub fn get_wallpaper(&self) -> Result<String> {
        let internal = self.0;
        with_com_objects(move |o| o.get_desktop_wallpaper(&internal))
    }

    /// Set desktop wallpaper path
    pub fn set_wallpaper(&self, path: &str) -> Result<()> {
        let internal = self.0;
        let path_ = path.to_owned();
        with_com_objects(move |o| o.set_desktop_wallpaper(&internal, &path_))
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
    T: Into<Desktop>,
    T: Send + 'static + Copy,
{
    with_com_objects(move |o| o.switch_desktop(&desktop.into().into()))
}

/// Remove desktop by index or GUID
pub fn remove_desktop<T>(desktop: T, fallback_desktop: T) -> Result<()>
where
    T: Into<Desktop>,
    T: Send + 'static + Copy,
{
    with_com_objects(move |o| {
        o.remove_desktop(&desktop.into().into(), &fallback_desktop.into().into())
    })
}

/// Is window on desktop by index or GUID
pub fn is_window_on_desktop<T>(desktop: T, hwnd: HWND) -> Result<bool>
where
    T: Into<Desktop>,
    T: Send + 'static + Copy,
{
    with_com_objects(move |o| o.is_window_on_desktop(&hwnd, &desktop.into().into()))
}

/// Move window to desktop by index or GUID
pub fn move_window_to_desktop<T>(desktop: T, hwnd: &HWND) -> Result<()>
where
    T: Into<Desktop>,
    T: Send + 'static + Copy,
{
    let hwnd = *hwnd;
    with_com_objects(move |o| o.move_window_to_desktop(&hwnd, &desktop.into().into()))
}

/// Create desktop
pub fn create_desktop() -> Result<Desktop> {
    with_com_objects(|o| o.create_desktop().map(Desktop))
}

/// Get current desktop
pub fn get_current_desktop() -> Result<Desktop> {
    with_com_objects(|o| o.get_current_desktop().map(Desktop))
}

/// Get all desktops
pub fn get_desktops() -> Result<Vec<Desktop>> {
    with_com_objects(|o| Ok(o.get_desktops()?.into_iter().map(Desktop).collect()))
}

/// Get desktop by window
pub fn get_desktop_by_window(hwnd: HWND) -> Result<Desktop> {
    with_com_objects(move |o| o.get_desktop_by_window(&hwnd).map(Desktop))
}

/// Get desktop count
pub fn get_desktop_count() -> Result<u32> {
    with_com_objects(|o| o.get_desktop_count())
}

pub fn is_window_on_current_desktop(hwnd: HWND) -> Result<bool> {
    with_com_objects(move |o| o.is_window_on_current_desktop(&hwnd))
}

/// Is window pinned?
pub fn is_pinned_window(hwnd: HWND) -> Result<bool> {
    with_com_objects(move |o| o.is_pinned_window(&hwnd))
}

/// Pin window
pub fn pin_window(hwnd: HWND) -> Result<()> {
    with_com_objects(move |o| o.pin_window(&hwnd))
}

/// Unpin window
pub fn unpin_window(hwnd: HWND) -> Result<()> {
    with_com_objects(move |o| o.unpin_window(&hwnd))
}

/// Is pinned app
pub fn is_pinned_app(hwnd: HWND) -> Result<bool> {
    with_com_objects(move |o| o.is_pinned_app(&hwnd))
}

/// Pin app
pub fn pin_app(hwnd: HWND) -> Result<()> {
    with_com_objects(move |o| o.pin_app(&hwnd))
}

/// Unpin app
pub fn unpin_app(hwnd: HWND) -> Result<()> {
    with_com_objects(move |o| o.unpin_app(&hwnd))
}
