
VirtualDesktopAccessor.dll
==========================

DLL for accessing Windows 10 Virtual Desktop features from e.g. AutoHotkey

Download the VirtualDesktopAccessor.dll from directory x64\Release\VirtualDesktopAccessor.dll in the repository. This DLL works only on 64 bit Windows 10.

You probably first need the [VS 2015 runtimes vc_redist.x64.exe and/or vc_redist.x86.exe](https://www.microsoft.com/en-us/download/details.aspx?id=48145), if they are not installed already. I've built the DLL using VS 2015, and Microsoft is not providing those runtimes (who knows why) with Windows 10 yet.

AutoHotkey script, and examples:

	DetectHiddenWindows, On
	hwnd:=WinExist("ahk_pid " . DllCall("GetCurrentProcessId","Uint"))
	hwnd+=0x1000<<32

	hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", "C:\Source\CandCPP\VirtualDesktopAccessor\x64\Release\VirtualDesktopAccessor.dll", "Ptr") 
	GoToDesktopNumberProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "GoToDesktopNumber", "Ptr")
	RegisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "RegisterPostMessageHook", "Ptr")
	UnregisterPostMessageHookProc := DllCall("GetProcAddress", Ptr, hVirtualDesktopAccessor, AStr, "UnregisterPostMessageHook", "Ptr")

	DllCall(RegisterPostMessageHookProc, Int, hwnd, Int, 0x1400 + 30)
	OnMessage(0x1400 + 30, "VWMess")
	VWMess(wParam, lParam, msg, hwnd) {
		
		; When switching to desktop 1, set background pluto.jpg
		if (lParam == 0) {
			DllCall("SystemParametersInfo", UInt, 0x14, UInt, 0, Str, "C:\Users\Jarppa\Pictures\Backgrounds\pluto.jpg", UInt, 1)
		; When switching to desktop 2, set background DeskGmail.png
		} else if (lParam == 1) {
			DllCall("SystemParametersInfo", UInt, 0x14, UInt, 0, Str, "C:\Users\Jarppa\Pictures\Backgrounds\DeskGmail.png", UInt, 1)
		; When switching to desktop 7 or 8, set background DeskMisc.png
		} else if (lParam == 6 || lParam == 7) {
			DllCall("SystemParametersInfo", UInt, 0x14, UInt, 0, Str, "C:\Users\Jarppa\Pictures\Backgrounds\DeskMisc.png", UInt, 1)
		; Other desktops, set background to DeskWork.png
		} else {
			DllCall("SystemParametersInfo", UInt, 0x14, UInt, 0, Str, "C:\Users\Jarppa\Pictures\Backgrounds\DeskWork.png", UInt, 1)
		}
	}
	; Win + Ctrl + 1 = Switch to desktop 1
	*#1::DllCall(GoToDesktopNumberProc, Int, 0)

	; Win + Ctrl + 2 = Switch to desktop 2
	*#2::DllCall(GoToDesktopNumberProc, Int, 1)

	; ...



Jari Pennanen, 2015