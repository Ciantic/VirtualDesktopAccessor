use std::convert::TryInto;

use winapi::shared::{guiddef::GUID, windef::HWND};
use winvd::{
    get_current_desktop, get_desktop_by_guid, get_desktop_by_index, get_desktop_by_window,
    helpers::*, is_pinned_app, is_pinned_window, is_window_on_current_desktop,
    is_window_on_desktop, move_window_to_desktop, pin_app, pin_window, unpin_app, unpin_window,
    DesktopID,
};

#[no_mangle]
pub extern "C" fn GetCurrentDesktopNumber() -> i32 {
    get_current_desktop_number().map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn GetDesktopCount() -> i32 {
    get_desktop_count().map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn GetDesktopIdByNumber(number: i32) -> DesktopID {
    if number < 0 {
        return DesktopID::default();
    }
    get_desktop_by_index(number as usize).map_or(DesktopID::default(), |desktop| desktop.get_id())
}

#[no_mangle]
pub extern "C" fn GetDesktopNumber() -> i32 {
    get_current_desktop_number().map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn GetDesktopNumberById(desktop_id: DesktopID) -> i32 {
    get_desktop_by_guid(&desktop_id).map_or(-1, |x| x.get_index().map_or(-1, |y| y as i32))
}

#[no_mangle]
pub extern "C" fn GetWindowDesktopId(hwnd: HWND) -> DesktopID {
    get_desktop_by_window(hwnd as u32).map_or(DesktopID::default(), |x| x.get_id())
}

#[no_mangle]
pub extern "C" fn GetWindowDesktopNumber(hwnd: HWND) -> i32 {
    get_desktop_by_window(hwnd as u32).map_or(-1, |x| x.get_index().map_or(-1, |y| y as i32))
}

#[no_mangle]
pub extern "C" fn IsWindowOnCurrentVirtualDesktop(hwnd: HWND) -> i32 {
    is_window_on_current_desktop(hwnd as u32).map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn MoveWindowToDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32 {
    move_window_to_desktop_number(hwnd as u32, desktop_number as usize).map_or(-1, |_| 1)
}

#[no_mangle]
pub extern "C" fn GoToDesktopNumber(desktop_number: i32) {
    go_to_desktop_number(desktop_number as usize).unwrap_or_default()
}

#[no_mangle]
pub extern "C" fn RegisterPostMessageHook(listener_hwnd: HWND, message_offset: i32) {
    todo!()
}

#[no_mangle]
pub extern "C" fn UnregisterPostMessageHook(listener_hwnd: HWND) {
    todo!()
}
#[no_mangle]
pub extern "C" fn IsPinnedWindow(hwnd: HWND) -> i32 {
    is_pinned_window(hwnd as u32).map_or(-1, |x| x as i32)
}
#[no_mangle]
pub extern "C" fn PinWindow(hwnd: HWND) {
    pin_window(hwnd as u32).unwrap_or_default()
}
#[no_mangle]
pub extern "C" fn UnPinWindow(hwnd: HWND) {
    unpin_window(hwnd as u32).unwrap_or_default()
}
#[no_mangle]
pub extern "C" fn IsPinnedApp(hwnd: HWND) -> i32 {
    is_pinned_app(hwnd as u32).map_or(-1, |x| x as i32)
}
#[no_mangle]
pub extern "C" fn PinApp(hwnd: HWND) {
    pin_app(hwnd as u32).unwrap_or_default()
}
#[no_mangle]
pub extern "C" fn UnPinApp(hwnd: HWND) {
    unpin_app(hwnd as u32).unwrap_or_default()
}
#[no_mangle]
pub extern "C" fn IsWindowOnDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32 {
    get_desktop_by_index(desktop_number as usize).map_or(-1, |x| {
        is_window_on_desktop(hwnd as u32, &x).map_or(-1, |b| b as i32)
    })
}
#[no_mangle]
pub extern "C" fn RestartVirtualDesktopAccessor() {
    // ?
}

#[no_mangle]
pub extern "C" fn lib_test() {
    println!("Hello from the library!");
}
/*
* int GetCurrentDesktopNumber()
* int GetDesktopCount()
* GUID GetDesktopIdByNumber(int number) // Returns zeroed GUID with invalid number found
* int GetDesktopNumber(IVirtualDesktop *pDesktop)
* int GetDesktopNumberById(GUID desktopId)
* GUID GetWindowDesktopId(HWND window)
* int GetWindowDesktopNumber(HWND window)
* int IsWindowOnCurrentVirtualDesktop(HWND window)
* BOOL MoveWindowToDesktopNumber(HWND window, int number)
* void GoToDesktopNumber(int number)
* void RegisterPostMessageHook(HWND listener, int messageOffset)
* void UnregisterPostMessageHook(HWND hwnd)
* int IsPinnedWindow(HWND hwnd) // Returns 1 if pinned, 0 if not pinned, -1 if not valid
* void PinWindow(HWND hwnd)
* void UnPinWindow(HWND hwnd)
* int IsPinnedApp(HWND hwnd) // Returns 1 if pinned, 0 if not pinned, -1 if not valid
* void PinApp(HWND hwnd)
* void UnPinApp(HWND hwnd)
* int IsWindowOnDesktopNumber(HWND window, int number) /
* void RestartVirtualDesktopAccessor() // Call this during taskbar created message

*/
