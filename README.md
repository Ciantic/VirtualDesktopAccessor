# VirtualDesktopAccessor.dll

DLL for accessing Windows 11 (22H2 Os Build 22621.963) Virtual Desktop features from e.g. AutoHotkey. MIT Licensed, see LICENSE.txt (c) Jari Pennanen, 2015-2023

## AutoHotkey example here:

* [AutoHotkey V1 example.ahk ⬅️](./example.ahk)
* [AutoHotkey V2 example.ah2 ⬅️](./example.ah2)

## Download from releases:

[Download the DLL from releases ⬇️](https://github.com/Ciantic/VirtualDesktopAccessor/releases/)

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
    go_to_desktop_number(1).unwrap();

    // Listen on interesting events
    // TODO: Document

}
```

See more examples from the [testbin sources 🢅](https://github.com/Ciantic/VirtualDesktopAccessor/blob/rust/testbin/src/main.rs).

### Notes

```
cargo clean
cargo build --release --workspace
```