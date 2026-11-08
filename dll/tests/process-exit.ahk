#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
#ErrorStdOut

; Keep the DLL loaded until process exit. FreeLibrary before ExitApp would
; avoid the shutdown ordering that this regression test needs to exercise.
; Suppress Windows crash dialogs only in this disposable test process.
DllCall("SetErrorMode", "UInt", 3, "UInt")

try {
    if A_Args.Length != 2
        throw Error("Expected absolute DLL path and mode (load-only or count)")
    library := DllCall("LoadLibraryW", "Str", A_Args[1], "Ptr")
    if !library
        throw Error("LoadLibraryW failed", , A_LastError)
    mode := A_Args[2]
    if mode = "count" {
        getCount := DllCall("GetProcAddress", "Ptr", library, "AStr", "GetDesktopCount", "Ptr")
        if !getCount
            throw Error("GetDesktopCount export not found")
        count := DllCall(getCount, "Int")
        if count < 1
            throw Error("GetDesktopCount failed; an interactive Windows desktop is required", , count)
        FileAppend("desktop_count=" count "`n", "*")
    } else if mode != "load-only" {
        throw Error("Unknown mode", , mode)
    }
    FileAppend("VDA_EXIT_READY`n", "*")
    ExitApp(0)
} catch as err {
    FileAppend("PROBE_ERROR: " err.Message " " err.Extra "`n", "**")
    ExitApp(2)
}
