# VirtualDesktopAccessor.dll

DLL for accessing Windows 11 (22H2 Os Build 22621.1105) Virtual Desktop features from e.g. AutoHotkey. MIT Licensed, see [LICENSE](LICENSE.txt) &copy; Jari Pennanen, 2015-2023

This repository also contains [Rust library `winvd`](./README-crate.md) for accessing the Virtual Desktop via Rust bindings.

## AutoHotkey example here:

* [AutoHotkey V1 example.ahk ⬅️](./example.ahk)
* [AutoHotkey V2 example.ah2 ⬅️](./example.ah2)

## Download from releases:

[Download the DLL from releases ⬇️](https://github.com/Ciantic/VirtualDesktopAccessor/releases/)

## Reference of exported DLL functions

All functions return -1 in case of error.

```rust
fn GetCurrentDesktopNumber() -> i32
fn GetDesktopCount() -> i32
fn GetDesktopIdByNumber(number: i32) -> GUID // Untested
fn GetDesktopNumberById(desktop_id: GUID) -> i32 // Untested
fn GetWindowDesktopId(hwnd: HWND) -> GUID
fn GetWindowDesktopNumber(hwnd: HWND) -> i32
fn IsWindowOnCurrentVirtualDesktop(hwnd: HWND) -> i32
fn MoveWindowToDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32
fn GoToDesktopNumber(desktop_number: i32) -> i32
fn SetDesktopName(desktop_number: i32, in_name_ptr: *const i8) -> i32  // Win11 only
fn GetDesktopName(desktop_number: i32, out_utf8_ptr: *mut u8, out_utf8_len: usize) -> i32 // Win11 only
fn RegisterPostMessageHook(listener_hwnd: HWND, message_offset: u32) -> i32
fn UnregisterPostMessageHook(listener_hwnd: HWND) -> i32
fn IsPinnedWindow(hwnd: HWND) -> i32
fn PinWindow(hwnd: HWND) -> i32
fn UnPinWindow(hwnd: HWND) -> i32
fn IsPinnedApp(hwnd: HWND) -> i32
fn PinApp(hwnd: HWND) -> i32
fn UnPinApp(hwnd: HWND) -> i32 
fn IsWindowOnDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32
fn CreateDesktop() -> i32 // Win11 only
fn RemoveDesktop(remove_desktop_number: i32, fallback_desktop_number: i32) -> i32 // Win11 only
```