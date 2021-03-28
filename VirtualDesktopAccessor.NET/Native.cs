using System;
using System.Runtime.InteropServices;

namespace VirtualDesktopAccessor
{
    using GUID = System.Guid;
    using HWND = System.IntPtr;
    using BOOL = System.Int32;
    using HRESULT = System.UInt32;
    using UINT = System.UInt32;
    using ULONGLONG = System.UInt64;


    public static class Native
    {
        [DllImport("VirtualDesktopAccessor")]
        public extern static void EnableKeepMinimized();
        [DllImport("VirtualDesktopAccessor")]
        public extern static void RestoreMinimized();
        [DllImport("VirtualDesktopAccessor")]
        public extern static int GetDesktopCount();
        [DllImport("VirtualDesktopAccessor")]
        public extern static int GetDesktopNumberById(GUID desktopId);
        [DllImport("VirtualDesktopAccessor")]
        public extern static GUID GetWindowDesktopId(HWND window);
        [DllImport("VirtualDesktopAccessor")]
        public extern static int GetWindowDesktopNumber(HWND window);
        [DllImport("VirtualDesktopAccessor")]
        public extern static int IsWindowOnCurrentVirtualDesktop(HWND window);
        [DllImport("VirtualDesktopAccessor")]
        public extern static GUID GetDesktopIdByNumber(int number);
        [DllImport("VirtualDesktopAccessor")]
        public extern static int IsWindowOnDesktopNumber(HWND window, int number);
        [DllImport("VirtualDesktopAccessor")]
        public extern static BOOL MoveWindowToDesktopNumber(HWND window, int number);
        [DllImport("VirtualDesktopAccessor")]
        public extern static int GetCurrentDesktopNumber();
        [DllImport("VirtualDesktopAccessor")]
        public extern static void GoToDesktopNumber(int number);
        [DllImport("VirtualDesktopAccessor")]
        public extern static int IsPinnedWindow(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static void PinWindow(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static void UnPinWindow(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static int IsPinnedApp(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static void PinApp(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static void UnPinApp(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static int ViewIsShownInSwitchers(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static int ViewIsVisible(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static HWND ViewGetThumbnailHwnd(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static HRESULT ViewSetFocus(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static HWND ViewGetFocused();
        [DllImport("VirtualDesktopAccessor")]
        public extern static HRESULT ViewSwitchTo(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static UINT ViewGetByZOrder(HWND[] windows, UINT count, BOOL onlySwitcherWindows, BOOL onlyCurrentDesktop);
        [DllImport("VirtualDesktopAccessor")]
        public extern static UINT ViewGetByLastActivationOrder(HWND[] windows, UINT count, BOOL onlySwitcherWindows, BOOL onlyCurrentDesktop);
        [DllImport("VirtualDesktopAccessor")]
        public extern static ULONGLONG ViewGetLastActivationTimestamp(HWND hwnd);
        [DllImport("VirtualDesktopAccessor")]
        public extern static void RestartVirtualDesktopAccessor();
        [DllImport("VirtualDesktopAccessor")]
        public extern static void RegisterPostMessageHook(HWND listener, int messageOffset);
        [DllImport("VirtualDesktopAccessor")]
        public extern static void UnregisterPostMessageHook(HWND hwnd);
    }
}
