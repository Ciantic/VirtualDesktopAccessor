
VirtualDesktopAccessor.dll
==========================

DLL for accessing Windows 10 Virtual Desktop features from e.g. AutoHotkey

AutoHotkey script, and examples:

	DetectHiddenWindows, On
	hwnd:=WinExist("ahk_pid " . DllCall("GetCurrentProcessId","Uint"))
	
	hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", "C:\Source\CandCPP\VirtualDesktopAccessor\x64\Release\VirtualDesktopAccessor.dll", "Ptr") 
	; Debug the problem with this
	; if !hVirtualDesktopAccessor
	; {
	   ; MsgBox Failed (error %ErrorLevel%; %A_LastError%).
	; }
	GoToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GoToDesktopNumber", "Ptr")
	; Debug the problem with this
	; if !GoToDesktopNumberProc
	; {
	   ; MsgBox Failed (error %ErrorLevel%; %A_LastError%).
	; }

	; Switch to desktop number 1
	DllCall(GoToDesktopNumberProc, Int, 0)

	; Switch to desktop number 2 ...
	DllCall(GoToDesktopNumberProc, Int, 1)


Jari Pennanen, 2015