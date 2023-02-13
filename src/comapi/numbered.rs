use windows::core::HSTRING;

use crate::Error;

use super::interfaces::ComIn;
use super::raw::*;
use super::Result;

pub fn get_desktop_count() -> Result<u32> {
    com_sta();
    let provider = get_iservice_provider()?;
    let manager = get_ivirtual_desktop_manager_internal(&provider)?;
    let desktops = get_idesktops_array(&manager)?;
    unsafe { desktops.GetCount().map_err(map_win_err) }
}

pub fn switch_to_desktop_index(number: u32) -> Result<()> {
    com_sta();
    let provider = get_iservice_provider()?;
    let manager = get_ivirtual_desktop_manager_internal(&provider)?;
    let desktop = get_idesktop_by_number(&manager, number)?;
    switch_to_idesktop(&manager, &desktop)
}

pub fn get_current_desktop_index() -> Result<u32> {
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

pub fn set_name_by_desktop_index(number: u32, name: &str) -> Result<()> {
    com_sta();
    let provider = get_iservice_provider()?;
    let manager = get_ivirtual_desktop_manager_internal(&provider)?;
    let desktop = get_idesktop_by_number(&manager, number)?;
    let name = HSTRING::from(name);
    unsafe { manager.set_name(ComIn::new(&desktop), name).as_result() }
}

pub fn get_name_by_desktop_index(number: u32) -> Result<String> {
    com_sta();
    let provider = get_iservice_provider()?;
    let manager = get_ivirtual_desktop_manager_internal(&provider)?;
    let desktop = get_idesktop_by_number(&manager, number)?;
    let mut name = HSTRING::new();
    unsafe { desktop.get_name(&mut name).as_result()? }
    Ok(name.to_string())
}
