# VirtualDesktopAccessor.dll

DLL for accessing Windows 11 (22H2 Os Build 22621.1105) Virtual Desktop features from e.g. AutoHotkey. 

MIT Licensed, see [LICENSE](LICENSE.txt) &copy; Jari Pennanen, 2015-2023

## AutoHotkey example here:

* [AutoHotkey V1 example.ahk ⬅️](./example.ahk)
* [AutoHotkey V2 example.ah2 ⬅️](./example.ah2)

## Download from releases:

[Download the DLL from releases ⬇️](https://github.com/Ciantic/VirtualDesktopAccessor/releases/)

## Reference of exported DLL functions

```rust
fn GetCurrentDesktopNumber() -> i32
fn GetDesktopCount() -> i32
fn GetDesktopIdByNumber(number: i32) -> DesktopID // Untested
fn GetDesktopNumber() -> i32
fn GetDesktopNumberById(desktop_id: DesktopID) -> i32
fn GetWindowDesktopId(hwnd: HWND) -> DesktopID
fn GetWindowDesktopNumber(hwnd: HWND) -> i32
fn IsWindowOnCurrentVirtualDesktop(hwnd: HWND) -> i32
fn MoveWindowToDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32
fn GoToDesktopNumber(desktop_number: i32)
fn SetDesktopName(desktop_number: i32, in_name_ptr: *const i8) -> i32  // Win11 only
fn GetDesktopName(desktop_number: i32, out_utf8_ptr: *mut u8, out_utf8_len: usize) -> i32 // Win11 only
fn RegisterPostMessageHook(listener_hwnd: HWND, message_offset: u32)
fn UnregisterPostMessageHook(listener_hwnd: HWND)
fn IsPinnedWindow(hwnd: HWND) -> i32
fn PinWindow(hwnd: HWND)
fn UnPinWindow(hwnd: HWND)
fn IsPinnedApp(hwnd: HWND) -> i32
fn PinApp(hwnd: HWND)
fn UnPinApp(hwnd: HWND)
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