# IVirtualDesktopNotification and apartments

`IVirtualDesktopNotificationService::Register` inherits the apartment from the calling thread, this means that before calling it ensure your thread is in multithreaded apartment, otherwise it's a bit difficult to receive the events.

Do this in whole new thread:

```C
// This initializes MTA = multi threaded apartment
CoInitializeEx(0, COINIT_MULTITHREADED);
IVirtualDesktopNotificationService::register(yourinstance, cookie);
// ...
```

However, for normal calls it's **important to use single threaded apartments**, otherwise you will encounter random crashes when doing quick switching of desktops etc.

```C
// COINIT_APARTMENTTHREADED is actually STA = Single threaded apartment, regardless of the name:
CoInitializeEx(0, COINIT_APARTMENTTHREADED); 
// Switch desktop
// Iterate desktops ...
// Get name of desktop ...
CoUninitialize();
```

When the Windows shell calls your listener methods, for example `IVirtualDesktopNotification::current_virtual_desktop_changed` it seems to be initializing the new thread as multithreaded apartment automatically, same as you used when calling the `register`. If you try to call single threaded apartment intialization inside these methods it will fail with: `HRESULT(0x80010106) "Cannot change thread mode after it is set."`

This also means that in order to make robust listener, which does not crash, you need to use channels or other communication from listener methods to another thread which is initialized as a single threaded apartment `COINIT_APARTMENTTHREADED`.

## Notes

* I tried to replace `CoInitializeEx(0, COINIT_MULTITHREADED);` with `CoIncrementMTAUsage() -> Cookie` and corresponding `CoDecrementMTAUsage(cookie)` but during drop of thread the `CoDecrementMTAUsage(cookie)` fails.