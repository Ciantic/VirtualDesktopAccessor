# VirtualDesktopAccessor.dll change log

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