# IVirtualDesktopNotification

In COM guidance the caller increases the reference count with `AddRef()` and the
callee should use `Release()`. This assumption doesn't seem to be true with
`IVirtualDesktopNotification`.

Notice that `IVirtualDesktopNotification` is an interface this code implements,
we don't call the methods on it, but Windows shell calls those, this means that
all values we get should have been added with a reference by the Windows shell
for us.

What is bizarre is that only the correctly functioning interface uses
`ManuallyDrop<T>` like this:

```rust
pub unsafe trait IVirtualDesktopNotification: IUnknown {
    pub unsafe fn virtual_desktop_created(
        &self,
        monitors: ManuallyDrop<IObjectArray>,
        desktop: ManuallyDrop<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_destroy_begin(
        &self,
        monitors: ManuallyDrop<IObjectArray>,
        desktop_destroyed: ManuallyDrop<IVirtualDesktop>,
        desktop_fallback: ManuallyDrop<IVirtualDesktop>,
    ) -> HRESULT;
    // ...
}
```

And we must never call `Release()` ourselves on the values.

Naively I would think it should be this:

```rust
pub unsafe trait IVirtualDesktopNotification: IUnknown {
    pub unsafe fn virtual_desktop_created(
        &self,
        monitors: IObjectArray,
        desktop: IVirtualDesktop,
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_destroy_begin(
        &self,
        monitors: IObjectArray,
        desktop_destroyed: IVirtualDesktop,
        desktop_fallback: IVirtualDesktop,
    ) -> HRESULT;
    // ...
}
```

Because `windows-rs` COM objects call `Release()` during `drop`. If I allow
calls to `drop` the crash occurs but after switching desktops repeatedly for
~2000-3000 times. This is not an easy bug to make happen.

The only way I get this work reliably is not to call the `Release()` in the
implementation's methods. It's as if Windows shell is calling the
`IVirtualDesktopNotification` methods with an assumption it's giving us a
reference and not a value.