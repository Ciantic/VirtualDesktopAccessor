# COM Guidance by David Risney:

> 1. When a COM object is passed from caller to callee as an input parameter to
>    a method, the caller is expected to keep a reference on the object for the
>    duration of the method call. The callee shouldn't need to call `AddRef` or
>    `Release` for the synchronous duration of that method call. For example, if
>    you're the callee implementing one of the async completed handler or event
>    handler `Invoke` methods, the WebView2 code will call your Invoke method
>    and pass in a COM object and you don't need to call `AddRef` or `Release`
>    on it.

> 2. When a COM object is passed from callee to caller as an out parameter from
>    a method the object is provided to the caller with a reference already
>    taken and the caller owns the reference. Which is to say, it is the
>    caller's responsibility to call `Release` when they're done with the
>    object. For example, if you call
>    `ICoreWebView2::get_Settings(ICoreWebView2Settings**)`, the `get_Settings`
>    code will call `AddRef` on the `ICoreWebView2Settings` it hands back to you
>    and its up to you only to call `Release` when you're done.

> 3. When making a copy of a COM object pointer you need to call `AddRef` and
>    `Release`. The `AddRef` must be called before you call `Release` on the
>    original COM object pointer. If for example you have an async method call
>    completion handler method that receives a COM object as an in parameter but
>    you need to deal with that COM object asynchronously later, you'll need to
>    make a copy of the COM object pointer and call `AddRef` during the
>    completion handler and then `Release` later after you finish your async
>    work with the object.

https://github.com/MicrosoftEdge/WebView2Feedback/issues/2133

## IVirtualDesktopNotification

To sum it up according to COM guidance the caller increases the reference count
with `AddRef()` and the callee should use `Release()` when getting the value
through `Out` parameter. However, when it's an `In` parameter the caller ensures
the value is held alive during the call and you must not call `Release()`.

Notice that `IVirtualDesktopNotification` is an interface this code implements,
we don't call the methods on it, but Windows shell calls those, this means that
all values we get should have been added with a reference by the Windows shell
for us.

This means the interface **must** be using `ManuallyDrop<T>`:

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

Naively one would think it can be done like this:

```rust
pub unsafe trait IVirtualDesktopNotification: IUnknown {
    pub unsafe fn virtual_desktop_created(
        &self,
        monitors: IObjectArray, // This is wrong, should be ManuallyDrop<IObjectArray>
        desktop: IVirtualDesktop, // This is wrong, should be ManuallyDrop<IVirtualDesktop>
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_destroy_begin(
        &self,
        monitors: IObjectArray, // This is wrong, should be ManuallyDrop<IObjectArray>
        desktop_destroyed: IVirtualDesktop, // This is wrong, should be ManuallyDrop<IVirtualDesktop>
        desktop_fallback: IVirtualDesktop, // This is wrong, should be ManuallyDrop<IVirtualDesktop>
    ) -> HRESULT;
    // ...
}
```

Because `windows-rs` COM objects call `Release()` during `drop` it will cause a
subtle bug. If I allow calls to `drop` the crash occurs but after switching
desktops repeatedly for ~2000-3000 times. This is not an easy bug to make
happen.


## Note about explorer.exe crashing and reusing cookies

I've observed that IVirtualDesktopNotification::register reuses cookies if
explorer.exe crashes. This means that unregistering must be done before a new
one is created.

Here is what I encountered:

1. Registered notification with cookie 24
2. Explorer.exe crashed
3. Explorer.exe restarted
4. Registered notification with cookie 24

If you were to now unregister the old one, it would unregister the new one. This
means we have to unregister the old value before registering a new one.