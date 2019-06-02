// TestVirtualDesktopAccessorWin32.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "TestVirtualDesktopAccessorWin32.h"
#include "dllmain.h"

WCHAR szClassName[] = L"TestVirtualDesktopAccesorWin32";

#define MESSAGE_OFFSET WM_USER + 60

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{

	switch (message) {
		case WM_DESTROY:
			PostQuitMessage(0);
			break;
		case MESSAGE_OFFSET + VDA_CurrentVirtualDesktopChanged:
			std::wcout << L"CurrentVirtualDesktopChanged old: " << wParam << " new:" << lParam << "\r\n";
			break;
		case MESSAGE_OFFSET + VDA_ViewVirtualDesktopChanged:
			std::wcout << L"CurrentVirtualDesktopChanged old: " << wParam << " new:" << lParam << "\r\n";
			break;
		case MESSAGE_OFFSET + VDA_VirtualDesktopCreated:
			std::wcout << L"CurrentVirtualDesktopChanged old: " << wParam << " new:" << lParam << "\r\n";
			break;
		case MESSAGE_OFFSET + VDA_VirtualDesktopDestroyBegin:
			std::wcout << L"CurrentVirtualDesktopChanged old: " << wParam << " new:" << lParam << "\r\n";
			break;
		case MESSAGE_OFFSET + VDA_VirtualDesktopDestroyed:
			std::wcout << L"CurrentVirtualDesktopChanged old: " << wParam << " new:" << lParam << "\r\n";
			break;
		case MESSAGE_OFFSET + VDA_VirtualDesktopDestroyFailed:
			std::wcout << L"CurrentVirtualDesktopChanged old: " << wParam << " new:" << lParam << "\r\n";
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
					 _In_opt_ HINSTANCE hPrevInstance,
					 _In_ LPWSTR    lpCmdLine,
					 _In_ int       nCmdShow)
{
	HWND hwnd;
	MSG messages;
	WNDCLASSEX wincl;

	wincl.hInstance = hInstance;
	wincl.lpszClassName = szClassName;
	wincl.lpfnWndProc = WndProc;
	wincl.style = CS_DBLCLKS;
	wincl.cbSize = sizeof(WNDCLASSEX);
	wincl.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wincl.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
	wincl.hCursor = LoadCursor(NULL, IDC_ARROW);
	wincl.lpszMenuName = NULL;
	wincl.cbClsExtra = 0;
	wincl.cbWndExtra = 0;
	wincl.hbrBackground = (HBRUSH)COLOR_BACKGROUND;

	if (!RegisterClassEx(&wincl))
		return 0;

	hwnd = CreateWindowEx(0, 
		szClassName, 
		L"TestVirtualDesktopAccesorWin32",
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		544,          
		375,          
		HWND_DESKTOP, 
		NULL,         
		hInstance,    
		NULL          
		);

	ShowWindow(hwnd, SW_HIDE);
	_OpenDllWindow(hInstance);

	RegisterPostMessageHook(hwnd, MESSAGE_OFFSET);
	std::wcout << "Desktops: " << GetDesktopCount() << "\r\n";
	std::wcout << "Console Window's Desktop Number: " << GetWindowDesktopNumber(GetConsoleWindow()) << std::endl;
	std::wcout << "Current Desktop Number: " << GetCurrentDesktopNumber() << "\r\n";

	HWND arr[1024];
	UINT numberOfWindows = ViewGetByLastActivationOrder((HWND*)&arr, 1024, 1, 0);
	std::wcout << "Number of windows: " << numberOfWindows << std::endl;


	HWND notepad = FindWindow(_T("Notepad"), NULL);
	if (notepad != 0) {
		int number = GetCurrentDesktopNumber();

		std::wcout << "Is notepad on this desktop: " << IsWindowOnDesktopNumber(notepad, number) << std::endl;
		std::wcout << "Is notepad on current desktop: " << IsWindowOnCurrentVirtualDesktop(notepad) << std::endl;
		std::wcout << "Is notepad shown on alt tab switcher: " << ViewIsShownInSwitchers(notepad) << std::endl;
		std::wcout << "Is notepad visible: " << ViewIsVisible(notepad) << std::endl;
		std::wcout << "Where is thumbnail of notepad: " << ViewGetThumbnailHwnd(notepad) << std::endl;

		// Test pinning it
		std::wcout << "Try pinning the notepad (the window)." << std::endl;
		PinWindow(notepad);
		std::wcout << "Is notepad pinned?" << IsPinnedWindow(notepad) << std::endl;
		GoToDesktopNumber(number + 1);
		Sleep(2000);
		GoToDesktopNumber(number);
		UnPinWindow(notepad);
		std::wcout << "Is notepad pinned?" << IsPinnedWindow(notepad) << std::endl;

		Sleep(2000);

		// Test pinning the whole app!
		std::wcout << "Try pinning the notepad (the app executable)." << std::endl;
		PinApp(notepad);
		std::wcout << "Is notepad pinned app?" << IsPinnedApp(notepad) << std::endl;
		GoToDesktopNumber(number + 1);
		Sleep(2000);
		GoToDesktopNumber(number);
		UnPinApp(notepad);
		std::wcout << "Is notepad pinned app?" << IsPinnedApp(notepad) << std::endl;

		// Test moving it
		if (!IsWindowOnCurrentVirtualDesktop(notepad)) {
			MoveWindowToDesktopNumber(notepad, number);
		}
		std::wcout << "Notepad is on on the current desktop. " << std::endl;
		Sleep(1000);
		MoveWindowToDesktopNumber(notepad, number + 1);
		std::wcout << "Notepad should now have been moved to the next desktop." << std::endl;
		Sleep(1000);
		MoveWindowToDesktopNumber(notepad, number);		
	}
	else {
		std::wcout << "Start notepad to try moving a window, or pinning it" << std::endl;
	}
	

	GUID g = GetDesktopIdByNumber(GetCurrentDesktopNumber());
	WCHAR text[255];
	StringFromGUID2(g, &text[0], 255);
	std::wcout << "Current Desktop GUID: " << text << std::endl;

	GUID g2 = GetWindowDesktopId(GetConsoleWindow());
	WCHAR text2[255];
	StringFromGUID2(g2, &text2[0], 255);
	std::wcout << "Console Window's Desktop GUID: " << text2 << std::endl;
	while (GetMessage(&messages, NULL, 0, 0))
	{
		TranslateMessage(&messages);
		DispatchMessage(&messages);
	}

	UnregisterPostMessageHook(hwnd);
	
	return messages.wParam;
}


int main() {
	return wWinMain(GetModuleHandle(NULL), NULL, GetCommandLine(), SW_SHOW);
}
