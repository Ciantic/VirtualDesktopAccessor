# VirtualDesktopAccessor.dll

DLL for accessing Windows 11 (22H2 Os Build 22621.1105) Virtual Desktop features from e.g. AutoHotkey. 

MIT Licensed, see [LICENSE](LICENSE.txt) &copy; Jari Pennanen, 2015-2023

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

---- 


## winvd - Windows virtual desktop bindings library for Rust

https://crates.io/crates/winvd
https://github.com/ciantic/VirtualDesktopAccessor/tree/rust/

The implementation abstracts the annoying COM API to a simple functions.
Accessing these functions should be thread-safe.

### Example

You may want to use `helpers` sub module in this crate, it is most stable API at
the moment. It contains almost all the wanted features but with numbered
helpers.

```rust
use winvd::helpers::{get_desktop_count, go_to_desktop_number};
use winvd::{get_event_receiver, VirtualDesktopEvent};

fn main() {
    // Desktop count
    println!("Desktops: {:?}", get_desktop_count());

    // Go to second desktop, index = 1
    go_to_desktop_number(1);

    // Listen on interesting events
    // TODO: Document

}
```

See more examples from the [testbin sources 🢅](https://github.com/Ciantic/VirtualDesktopAccessor/blob/rust/testbin/src/main.rs).

### Notes

```
cargo clean
cargo build --release --workspace
cargo build --features debug --workspace
```