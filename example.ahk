hVirtualDesktopAccessor := DllCall("LoadLibrary", Str, "target\release\VirtualDesktopAccessor.dll", "Ptr")

GoToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GoToDesktopNumber", "Ptr")
GetCurrentDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetCurrentDesktopNumber", "Ptr")
IsWindowOnCurrentVirtualDesktopProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsWindowOnCurrentVirtualDesktop", "Ptr")
IsWindowOnDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsWindowOnDesktopNumber", "Ptr")
MoveWindowToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "MoveWindowToDesktopNumber", "Ptr")
RegisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "RegisterPostMessageHook", "Ptr")
UnregisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "UnregisterPostMessageHook", "Ptr")
RestoreMinimizedProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "RestoreMinimized", "Ptr")
EnableKeepMinimizedProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "EnableKeepMinimized", "Ptr")
IsPinnedWindowProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsPinnedWindow", "Ptr")

activeWindowByDesktop := {}

MoveCurrentWindowToDesktop(number) {
    global MoveWindowToDesktopNumberProc, GoToDesktopNumberProc, activeWindowByDesktop
    WinGet, activeHwnd, ID, A
    activeWindowByDesktop[number] := 0 ; Do not activate
    DllCall(MoveWindowToDesktopNumberProc, UInt, activeHwnd, UInt, number)
    DllCall(GoToDesktopNumberProc, UInt, number)
}

GoToPrevDesktop() {
    global GetCurrentDesktopNumberProc, GoToDesktopNumberProc
    current := DllCall(GetCurrentDesktopNumberProc, UInt)
    if (current = 0) {
        GoToDesktopNumber(7)
    } else {
        GoToDesktopNumber(current - 1)
    }
    return
}

GoToNextDesktop() {
    global GetCurrentDesktopNumberProc, GoToDesktopNumberProc
    current := DllCall(GetCurrentDesktopNumberProc, UInt)
    if (current = 7) {
        GoToDesktopNumber(0)
    } else {
        GoToDesktopNumber(current + 1)
    }
    return
}

GoToDesktopNumber(num) {
    global GetCurrentDesktopNumberProc, GoToDesktopNumberProc, IsPinnedWindowProc, activeWindowByDesktop

    ; Store the active window of old desktop, if it is not pinned
    WinGet, activeHwnd, ID, A
    current := DllCall(GetCurrentDesktopNumberProc, UInt)
    isPinned := DllCall(IsPinnedWindowProc, UInt, activeHwnd)
    if (isPinned == 0) {
        activeWindowByDesktop[current] := activeHwnd
    }

    ; Try to avoid flashing task bar buttons, deactivate the current window if it is not pinned
    if (isPinned != 1) {
        WinActivate, ahk_class Shell_TrayWnd
    }

    ; Change desktop
    DllCall(GoToDesktopNumberProc, Int, num)
    return
}
MoveOrGotoDesktopNumber(num) {
    ; If user is holding down Mouse left button, move the current window also
    if (GetKeyState("LButton")) {
        MoveCurrentWindowToDesktop(num)
    } else {
        GoToDesktopNumber(num)
    }
    return
}
#+1::MoveOrGotoDesktopNumber(0)
#+2::MoveOrGotoDesktopNumber(1)
#+3::MoveOrGotoDesktopNumber(2)
#+4::MoveOrGotoDesktopNumber(3)
#+5::MoveOrGotoDesktopNumber(4)