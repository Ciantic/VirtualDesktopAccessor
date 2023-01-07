; Path to the DLL, relative to the script
VDA_PATH := A_ScriptDir . "\target\release\VirtualDesktopAccessor.dll"
hVirtualDesktopAccessor := DllCall("LoadLibrary", Str, VDA_PATH, "Ptr")

GoToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GoToDesktopNumber", "Ptr")
GetCurrentDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetCurrentDesktopNumber", "Ptr")
IsWindowOnCurrentVirtualDesktopProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsWindowOnCurrentVirtualDesktop", "Ptr")
IsWindowOnDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsWindowOnDesktopNumber", "Ptr")
MoveWindowToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "MoveWindowToDesktopNumber", "Ptr")
RestoreMinimizedProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "RestoreMinimized", "Ptr")
EnableKeepMinimizedProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "EnableKeepMinimized", "Ptr")
IsPinnedWindowProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsPinnedWindow", "Ptr")

; On change listeners
RegisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "RegisterPostMessageHook", "Ptr")
UnregisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "UnregisterPostMessageHook", "Ptr")

MoveCurrentWindowToDesktop(number) {
    global MoveWindowToDesktopNumberProc, GoToDesktopNumberProc
    WinGet, activeHwnd, ID, A
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
    global GoToDesktopNumberProc
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

; DllCall(RegisterPostMessageHookProc, Int, hwnd, Int, 0x1400 + 30)
; OnMessage(0x1400 + 30, "OnChangeDesktop")
; OnChangeDesktop(wParam, lParam, msg, hwnd) {
;     desktopNumber := lParam + 1
;     MsgBox, desktopNumber
; }

#+1::MoveOrGotoDesktopNumber(0)
#+2::MoveOrGotoDesktopNumber(1)
#+3::MoveOrGotoDesktopNumber(2)
#+4::MoveOrGotoDesktopNumber(3)
#+5::MoveOrGotoDesktopNumber(4)