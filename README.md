# winvd - Windows virtual desktop bindings library for Rust

First version available now in: https://crates.io/crates/winvd

## Example

```rust
use winvd::VirtualDesktopService;

fn main() {
    let service = VirtualDesktopService::create_with_com().unwrap();

    // Show all desktops
    let desktops = service.get_desktops().unwrap();
    println!("Desktops {:?}", desktops);

    // Listen on desktop changes
    service.on_desktop_change(Box::new(|old, new| {
        println!("Desktop changed from {:?} to {:?}", old, new);
    }));

    // Go to second desktop, index = 1
    let second_desktop = desktops.get(1).unwrap();
    service.go_to_desktop(second_desktop).unwrap();

    // See more examples from the testbin
    // https://github.com/Ciantic/VirtualDesktopAccessor/blob/rust/testbin/src/main.rs
}
```

## Other

This might deprecate CPP implementation, once I get a DLL also done with Rust.
