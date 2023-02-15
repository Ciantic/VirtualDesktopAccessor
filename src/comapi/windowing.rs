use windows::Win32::Foundation::HWND;

use super::raw::*;
use super::Result;

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
