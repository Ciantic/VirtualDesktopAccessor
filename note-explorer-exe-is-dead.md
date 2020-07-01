If you didn't start the explorer exe in the first place you will now get: `Err(ComClassNotRegistered)` from the `VirtualDesktopService::create_with_com()`.

If you kill the explorer.exe, or restart explorer.exe, the COM API starts to give following results:

```
Desktop of notepad: Err(ComError(HRESULT(0x800706BA), "IVirtualDesktopManager.get_desktop_by_window")), hwnd: 4131168
Notepad is on current desktop: Err(ComError(HRESULT(0x800706BA), "IVirtualDesktopManager.is_window_on_current_desktop"))
Is notepad on desktop: DesktopID(5F693E2A-AF18-4ACF-86E2-CBEFA84E8033), true or false: Err(ComError(HRESULT(0x800706BA), "IVirtualDesktopManager.get_desktop_by_window"))
Try to move non existant window... Err(ComError(HRESULT(0x800706BA), "IVirtualDesktopManagerInternal.get_desktops"))
Move notepad to first desktop for three seconds, and then return it...
Move to first... Err(ComError(HRESULT(0x800706BA), "IVirtualDesktopManagerInternal.get_desktops"))
Wait three seconds...
Move back to this desktop Err(ComError(HRESULT(0x800706BA), "IVirtualDesktopManagerInternal.get_desktops"))
Pin the notepad window Err(ComError(HRESULT(0x800706BA), "IApplicationView.get_view_for_hwnd"))
Switch between desktops 1 and this one...
Move to first... Err(ComError(HRESULT(0x800706BA), "IVirtualDesktopManagerInternal.get_desktops"))
Wait three seconds...
Move back to this desktop Err(ComError(HRESULT(0x800706BA), "IVirtualDesktopManagerInternal.get_desktops"))
Unpin the notepad window Err(ComError(HRESULT(0x800706BA), "IApplicationView.get_view_for_hwnd"))
```

This means `0x800706BA` is good indication I have to restart the service if I get `0x800706BA`
