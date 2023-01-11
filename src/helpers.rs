//! This module contains numbered helpers, if you prefer handling your desktops by a number.
//!
//! This is currently the most stable API.

use crate::{
    get_current_desktop, get_desktop_by_index, get_desktop_by_window, get_desktops,
    get_index_by_desktop, go_to_desktop, is_window_on_desktop, move_window_to_desktop,
    set_desktop_name, Error, HWND,
};

/// Get number of desktops
pub fn get_desktop_count() -> Result<usize, Error> {
    Ok(get_desktops()?.len())
}

/// Get current desktop number
pub fn get_current_desktop_number() -> Result<usize, Error> {
    get_index_by_desktop(&get_current_desktop()?)
}

/// Get desktop number by window
pub fn get_desktop_number_by_window(hwnd: HWND) -> Result<usize, Error> {
    get_index_by_desktop(&get_desktop_by_window(hwnd)?)
}

/// Is window on desktop number
pub fn is_window_on_desktop_number(hwnd: HWND, number: usize) -> Result<bool, Error> {
    is_window_on_desktop(hwnd, &get_desktop_by_index(number)?)
}

/// Rename desktop
pub fn rename_desktop_number(number: usize, name: &str) -> Result<(), Error> {
    set_desktop_name(&get_desktop_by_index(number)?, name)
}

/// Get name by desktop number
pub fn get_name_by_desktop_number(number: usize) -> Result<String, Error> {
    get_desktop_by_index(number)?.get_name()
}

/// Move window to desktop number
pub fn move_window_to_desktop_number(hwnd: HWND, number: usize) -> Result<(), Error> {
    move_window_to_desktop(hwnd, &get_desktop_by_index(number)?)
}

/// Go to desktop number
pub fn go_to_desktop_number(number: usize) -> Result<(), Error> {
    go_to_desktop(&get_desktop_by_index(number)?)
}
