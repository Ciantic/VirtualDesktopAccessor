# winvd - Windows virtual desktop bindings library for Rust

First version available now in: https://crates.io/crates/winvd

## Example

```rust
use winvd::VirtualDesktopService;

fn main() {
    let service = VirtualDesktopService::create_with_com().unwrap();
    println!("Desktops {:?}", service.get_desktops().unwrap());
}
```

## Other

This might deprecate CPP implementation, once I get a DLL also done with Rust.
