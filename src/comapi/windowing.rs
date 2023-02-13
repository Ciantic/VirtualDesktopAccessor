type HWND = u32;
use super::raw::*;
use super::Result;

pub fn move_window_to_desktop_index(hwnd: HWND, number: u32) -> Result<()> {
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

pub fn is_window_on_desktop_index(hwnd: HWND, number: u32) -> Result<bool> {
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

pub fn get_desktop_index_by_window(hwnd: HWND) -> Result<u32> {
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
