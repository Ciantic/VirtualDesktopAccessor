# winvd - Windows virtual desktop bindings library for Rust

https://crates.io/crates/winvd
https://github.com/ciantic/VirtualDesktopAccessor/tree/rust/

The implementation abstracts the annoying COM API to a simple functions. Accessing these functions should be thread-safe.

## Example

```rust
use winvd::{get_desktop_count, go_to_desktop, get_event_receiver};

fn main() {
    // Desktop count
    println!("Desktops: {:?}", get_desktop_count());

    // Go to second desktop, index = 1
    go_to_desktop(1).unwrap();

    // Listen on interesting events
    std::thread::spawn(|| {
        get_event_receiver().iter().for_each(|msg| match msg {
            VirtualDesktopEvent::DesktopChanged(old, new) => {
                println!(
                    "<- Desktop changed from {:?} to {:?}",
                    old,
                    new
                );
            }
            VirtualDesktopEvent::DesktopCreated(desk) => {
                println!("<- New desktop created {:?}", desk);
            }
            VirtualDesktopEvent::DesktopDestroyed(desk) => {
                println!("<- Desktop destroyed {:?}", desk);
            }
            VirtualDesktopEvent::WindowChanged(hwnd) => {
                println!("<- Window changed {:?}", hwnd);
            }
        })
    });

}
```

See more examples from the [testbin sources 🢅](https://github.com/Ciantic/VirtualDesktopAccessor/blob/rust/testbin/src/main.rs).

## When explorer.exe restarts

In case you want a robust event listener, you need to notify when the
explorer.exe restarts. Listen on window message loop [for `TaskbarCreated`
messages 🢅](https://docs.microsoft.com/en-us/windows/win32/shell/taskbar#taskbar-creation-notification), and call the `notify_explorer_restarted` to recreate the underlying sender loop.

## Other

This might deprecate CPP implementation, once I get a DLL also done with Rust.
