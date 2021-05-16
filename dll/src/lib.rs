use winapi::shared::{guiddef::GUID, windef::HWND};
use winvd::{
    get_current_desktop, get_desktop_by_window, helpers::*, is_window_on_current_desktop,
    move_window_to_desktop,
};

#[no_mangle]
pub extern "C" fn lib_test() {
    println!("Hello from the library!");
}

#[no_mangle]
pub extern "C" fn GetCurrentDesktopNumber() -> i32 {
    get_current_desktop_number().unwrap() as i32
}

#[no_mangle]
pub extern "C" fn GetDesktopCount() -> i32 {
    get_desktop_count().unwrap() as i32
}

#[no_mangle]
pub extern "C" fn GetDesktopIdByNumber(number: i32) -> GUID {
    // let desk = get_current_desktop().unwrap();
    // desk
    todo!()
}

#[no_mangle]
pub extern "C" fn GetDesktopNumber() -> i32 {
    get_current_desktop_number().unwrap() as i32
}

#[no_mangle]
pub extern "C" fn GetDesktopNumberById(desktop_id: GUID) -> i32 {
    todo!()
}

#[no_mangle]
pub extern "C" fn GetWindowDesktopId(hwnd: HWND) -> GUID {
    // let desk = get_desktop_by_window(hwnd as u32).unwrap();
    todo!()
}

#[no_mangle]
pub extern "C" fn GetWindowDesktopNumber(hwnd: HWND) -> u32 {
    let desk = get_desktop_by_window(hwnd as u32).unwrap();
    desk.get_index().unwrap() as u32
}

#[no_mangle]
pub extern "C" fn IsWindowOnCurrentVirtualDesktop(hwnd: HWND) -> bool {
    is_window_on_current_desktop(hwnd as u32).unwrap()
}

#[no_mangle]
pub extern "C" fn MoveWindowToDesktopNumber(hwnd: HWND, desktop_number: u32) -> u32 {
    move_window_to_desktop_number(hwnd as u32, desktop_number as usize).unwrap();
    1
}

#[no_mangle]
pub extern "C" fn GoToDesktopNumber(desktop_number: u32) {
    go_to_desktop_number(desktop_number as usize).unwrap()
}

#[no_mangle]
pub extern "C" fn RegisterPostMessageHook(listener_hwnd: HWND, message_offset: u32) {
    todo!()
}

#[no_mangle]
pub extern "C" fn UnregisterPostMessageHook(listener_hwnd: HWND) {
    todo!()
}
#[no_mangle]
pub extern "C" fn IsPinnedWindow(hwnd: HWND) -> bool {
    todo!()
}
#[no_mangle]
pub extern "C" fn PinWindow(hwnd: HWND) {
    todo!()
}
#[no_mangle]
pub extern "C" fn UnPinWindow(hwnd: HWND) {
    todo!()
}
#[no_mangle]
pub extern "C" fn IsPinnedApp(hwnd: HWND) -> bool {
    todo!()
}
#[no_mangle]
pub extern "C" fn PinApp(hwnd: HWND) {
    todo!()
}
#[no_mangle]
pub extern "C" fn UnPinApp(hwnd: HWND) {
    todo!()
}
#[no_mangle]
pub extern "C" fn IsWindowOnDesktopNumber(window: HWND, desktop_number: u32) -> u32 {
    todo!()
}
#[no_mangle]
pub extern "C" fn RestartVirtualDesktopAccessor() {
    todo!()
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
