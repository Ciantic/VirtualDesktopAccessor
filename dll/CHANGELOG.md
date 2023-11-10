# VirtualDesktopAccessor.dll change log

## Windows 11, Eight binary - IVirtualDesktopNotification changes (2023-11-10)

My interface definition was missing `virtual_desktop_switched` and
`remote_virtual_desktop_connected` this caused a memory corruption if you were
lucky.

## Windows 11, Seventh binary - New interface definitions (2023-08-31)

Interface definitions changed in the latest Windows 11 build 22621, this is not
backward compatible with old Windows 11 binaries! 

## Windows 11, Sixth binary (not released)

Now that I have re-learned the COM lifetime rules the hard way, I could get this working with multithreaded apartments.

The change is not causing any visible changes to DLL API, just removing old code about single-threaded apartments.


## Windows 11, Fifth binary (2023-02-22)

Earlier Windows 11 DLLs had so many bugs that I can't even list them. 

**NOTE** Some DLL functions changed signature, be cautious.

But let me summarize:

* Apparently, all my assumptions about COM object methods `AddRef` and `Release` were totally wrong
* My `HWND` type was 32-bit, but it should be 64-bit, the same as pointers
* This means the previous versions worked out of sheer luck

How the heck did the earlier versions even work? 