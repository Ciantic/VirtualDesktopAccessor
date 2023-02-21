; AutoHotkey v1 script

; Get hwnd of AutoHotkey window, for listener
DetectHiddenWindows, On
ahkWindowHwnd := WinExist("ahk_pid " . DllCall("GetCurrentProcessId", "Uint"))
ahkWindowHwnd += 0x1000 << 32

; Path to the DLL, relative to the script
VDA_PATH := A_ScriptDir . "\target\release\VirtualDesktopAccessor.dll"
hVirtualDesktopAccessor := DllCall("LoadLibrary", Str, VDA_PATH, "Ptr")

GetDesktopCountProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetDesktopCount", "Ptr")
GoToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GoToDesktopNumber", "Ptr")
GetCurrentDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetCurrentDesktopNumber", "Ptr")
IsWindowOnCurrentVirtualDesktopProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsWindowOnCurrentVirtualDesktop", "Ptr")
IsWindowOnDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsWindowOnDesktopNumber", "Ptr")
MoveWindowToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "MoveWindowToDesktopNumber", "Ptr")
IsPinnedWindowProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "IsPinnedWindow", "Ptr")
GetDesktopNameProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GetDesktopName", "Ptr")
SetDesktopNameProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "SetDesktopName", "Ptr")
CreateDesktopProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "CreateDesktop", "Ptr")
RemoveDesktopProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "RemoveDesktop", "Ptr")

; On change listeners
RegisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "RegisterPostMessageHook", "Ptr")
UnregisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "UnregisterPostMessageHook", "Ptr")

GetDesktopCount() {
    global GetDesktopCountProc
    count := DllCall(GetDesktopCountProc, UInt)
    return count
}

MoveCurrentWindowToDesktop(number) {
    global MoveWindowToDesktopNumberProc, GoToDesktopNumberProc
    WinGet, activeHwnd, ID, A
    DllCall(MoveWindowToDesktopNumberProc, UInt, activeHwnd, UInt, number)
    DllCall(GoToDesktopNumberProc, UInt, number)
}

GoToPrevDesktop() {
    global GetCurrentDesktopNumberProc, GoToDesktopNumberProc
    current := DllCall(GetCurrentDesktopNumberProc, UInt)
    last_desktop := GetDesktopCount() - 1
    ; If current desktop is 0, go to last desktop
    if (current = 0) {
        MoveOrGotoDesktopNumber(last_desktop)
    } else {
        MoveOrGotoDesktopNumber(current - 1)
    }
    return
}

GoToNextDesktop() {
    global GetCurrentDesktopNumberProc, GoToDesktopNumberProc
    current := DllCall(GetCurrentDesktopNumberProc, UInt)
    last_desktop := GetDesktopCount() - 1
    ; If current desktop is last, go to first desktop
    if (current = last_desktop) {
        MoveOrGotoDesktopNumber(0)
    } else {
        MoveOrGotoDesktopNumber(current + 1)
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
GetDesktopName(num) {
    global GetDesktopNameProc
    utf8_buffer := ""
    utf8_buffer_len := VarSetCapacity(utf8_buffer, 1024, 0)
    ran := DllCall(GetDesktopNameProc, UInt, num, Ptr, &utf8_buffer, UInt, utf8_buffer_len)
    name := StrGet(&utf8_buffer, 1024, "UTF-8")
    return name
}
SetDesktopName(num, name) {
    ; NOTICE! For UTF-8 to work AHK file must be saved with UTF-8 with BOM

    global SetDesktopNameProc
    VarSetCapacity(name_utf8, 1024, 0)
    StrPut(name, &name_utf8, "UTF-8")
    ran := DllCall(SetDesktopNameProc, UInt, num, Ptr, &name_utf8)
    return ran
}
CreateDesktop() {
    global CreateDesktopProc
    ran := DllCall(CreateDesktopProc)
    return ran
}
RemoveDesktop(remove_desktop_number, fallback_desktop_number) {
    global RemoveDesktopProc
    ran := DllCall(RemoveDesktopProc, "UInt", remove_desktop_number, "UInt", fallback_desktop_number)
    return ran
}

; SetDesktopName(0, "It works! 🐱")

; How to listen to desktop changes
DllCall(RegisterPostMessageHookProc, Int, ahkWindowHwnd, Int, 0x1400 + 30)
OnMessage(0x1400 + 30, "OnChangeDesktop")
OnChangeDesktop(wParam, lParam, msg, hwnd) {
    Critical, 100
    OldDesktop := wParam + 1
    NewDesktop := lParam + 1
    Name := GetDesktopName(NewDesktop - 1)

    ; Use Dbgview.exe to checkout the output debug logs
    OutputDebug % "Desktop changed to " Name " from " OldDesktop " to " NewDesktop
}

#+1:: MoveOrGotoDesktopNumber(0)
#+2:: MoveOrGotoDesktopNumber(1)
#+3:: MoveOrGotoDesktopNumber(2)
#+4:: MoveOrGotoDesktopNumber(3)
#+5:: MoveOrGotoDesktopNumber(4)

F13 & 1:: MoveOrGotoDesktopNumber(0)
F13 & 2:: MoveOrGotoDesktopNumber(1)
F13 & 3:: MoveOrGotoDesktopNumber(2)
F13 & 4:: MoveOrGotoDesktopNumber(3)
F13 & 5:: MoveOrGotoDesktopNumber(4)
F14 UP:: GoToPrevDesktop()
F15 UP:: GoToNextDesktop()