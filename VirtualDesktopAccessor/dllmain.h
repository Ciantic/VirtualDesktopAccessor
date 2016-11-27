#pragma once
#include "stdafx.h"
#include "Win10Desktops.h"

#define DllExport   __declspec( dllexport )

#define VDA_VirtualDesktopCreated 5
#define VDA_VirtualDesktopDestroyBegin 4
#define VDA_VirtualDesktopDestroyFailed 3
#define VDA_VirtualDesktopDestroyed 2
#define VDA_ViewVirtualDesktopChanged 1
#define VDA_CurrentVirtualDesktopChanged 0

#define VDA_IS_NORMAL 1
#define VDA_IS_MINIMIZED 2
#define VDA_IS_MAXIMIZED 3

std::map<HWND, int> listeners;
IServiceProvider* pServiceProvider = nullptr;
IVirtualDesktopManagerInternal *pDesktopManagerInternal = nullptr;
IVirtualDesktopManager *pDesktopManager = nullptr;
IApplicationViewCollection *viewCollection = nullptr;
IVirtualDesktopPinnedApps *pinnedApps = nullptr;
IVirtualDesktopNotificationService* pDesktopNotificationService = nullptr;
BOOL registeredForNotifications = FALSE;

std::map<HWND, int> oldWindowCmds;
std::vector<HWND> hideWindowes;
std::vector<HWND> showWindowes;
DWORD idNotificationService = 0;
BOOL _changingDesktop = false;
BOOL _keepMinimized = false;

struct ChangeDesktopAction {
	GUID newDesktopGuid;
	GUID oldDesktopGuid;
};

void _PostMessageToListeners(int msgOffset, WPARAM wParam, LPARAM lParam) {
	for each (std::pair<HWND, int> listener in listeners) {
		PostMessage(listener.first, listener.second + msgOffset, wParam, lParam);
	}
}

void _RegisterService(BOOL force = FALSE) {
	if (force) {
		pServiceProvider = nullptr;
		pDesktopManagerInternal = nullptr;
		pDesktopManager = nullptr;
		viewCollection = nullptr;
		pinnedApps = nullptr;
		pDesktopNotificationService = nullptr;
		registeredForNotifications = FALSE;
	}

	if (pServiceProvider != nullptr) {
		return;
	}
	::CoInitialize(NULL);
	::CoCreateInstance(
		CLSID_ImmersiveShell, NULL, CLSCTX_LOCAL_SERVER,
		__uuidof(IServiceProvider), (PVOID*)&pServiceProvider);

	if (pServiceProvider == nullptr) {
		std::wcout << L"FATAL ERROR: pServiceProvider is null";
		return;
	}
	pServiceProvider->QueryService(__uuidof(IApplicationViewCollection), &viewCollection);

	pServiceProvider->QueryService(__uuidof(IVirtualDesktopManager), &pDesktopManager);

	pServiceProvider->QueryService(
		CLSID_VirtualDesktopPinnedApps,
		__uuidof(IVirtualDesktopPinnedApps), (PVOID*)&pinnedApps);

	pServiceProvider->QueryService(
		CLSID_VirtualDesktopManagerInternal,
		__uuidof(IVirtualDesktopManagerInternal), (PVOID*)&pDesktopManagerInternal);

	if (viewCollection == nullptr) {
		std::wcout << L"FATAL ERROR: viewCollection is null";
		return;
	}

	if (pDesktopManagerInternal == nullptr) {
		std::wcout << L"FATAL ERROR: pDesktopManagerInternal is null";
		return;
	}

	// Notification service
	HRESULT hrNotificationService = pServiceProvider->QueryService(
		CLSID_IVirtualNotificationService,
		__uuidof(IVirtualDesktopNotificationService),
		(PVOID*)&pDesktopNotificationService);
}


IApplicationView* _GetApplicationViewForHwnd(HWND hwnd) {
	if (hwnd == 0)
		return nullptr;
	IApplicationView* app = nullptr;
	viewCollection->GetViewForHwnd(hwnd, &app);
	return app;
}

__inline BOOL CALLBACK
_EnumWindowProc_ChangeDesktop(HWND hwnd, LPARAM lParam)
{
	ChangeDesktopAction* act = (ChangeDesktopAction*)lParam;
	if (act == nullptr) {
		return TRUE;
	}
	IApplicationView* view = _GetApplicationViewForHwnd(hwnd);
	if (view != nullptr) {
		GUID winDesktopGuid;
		if (SUCCEEDED(view->GetVirtualDesktopId(&winDesktopGuid))) {
			if (winDesktopGuid == act->oldDesktopGuid) {
				std::wcout << "Old desktop's window " << hwnd << "\r\n";

				int style;
				if ((style = GetWindowLong(hwnd, GWL_STYLE)) & WS_CHILD) // Ignore all windows with child flag set
					return TRUE;

				if (style & WS_MINIMIZE) {
					oldWindowCmds[hwnd] = VDA_IS_MINIMIZED;
				}
				else if (style & WS_MAXIMIZE) {
					oldWindowCmds[hwnd] = VDA_IS_MAXIMIZED;
				}
				else {
					oldWindowCmds[hwnd] = VDA_IS_NORMAL;
				}
				hideWindowes.push_back(hwnd);
			}
			else if (winDesktopGuid == act->newDesktopGuid) {
				showWindowes.insert(showWindowes.begin(), hwnd);
			}
		}
		view->Release();
	}
	return TRUE;
}

void _ChangeDesktop_HideOld(BOOL async = false) {
	for each (HWND win in hideWindowes)
	{
		if (async) {
			ShowWindowAsync(win, SW_SHOWMINNOACTIVE);
		}
		else {
			ShowWindow(win, SW_SHOWMINNOACTIVE);
		}
	}

}

void _ChangeDesktop_ShowNew(BOOL async = false) {
	HWND active = GetForegroundWindow();
	for each (HWND win in showWindowes)
	{
		if ((oldWindowCmds[win] == VDA_IS_MAXIMIZED || oldWindowCmds[win] == VDA_IS_NORMAL) && IsIconic(win)) {

			if (async) {
				ShowWindowAsync(win, SW_SHOWNOACTIVATE);
			}
			else {
				ShowWindow(win, SW_SHOWNOACTIVATE);
			}
			
		}
	}
}

void _ChangeDesktop_ListWins(ChangeDesktopAction act) {
	hideWindowes.clear();
	showWindowes.clear();
	EnumWindows(_EnumWindowProc_ChangeDesktop, (LPARAM)&act);

}

void DllExport EnableKeepMinimized() {
	_keepMinimized = true;
}

void DllExport RestoreMinimized() {
	std::map<HWND, int>::iterator it;
	for (it = oldWindowCmds.begin(); it != oldWindowCmds.end(); it++)
	{
		if ((it->second == VDA_IS_MAXIMIZED || it->second == VDA_IS_NORMAL) && IsIconic(it->first)) {
			ShowWindowAsync(it->first, SW_SHOWNOACTIVATE);
		}
	}
}

int DllExport GetDesktopCount()
{
	_RegisterService();

	IObjectArray *pObjectArray = nullptr;
	HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);

	if (SUCCEEDED(hr))
	{
		UINT count;
		hr = pObjectArray->GetCount(&count);
		pObjectArray->Release();
		return count;
	}

	return -1;
}

int DllExport GetDesktopNumberById(GUID desktopId) {
	_RegisterService();

	IObjectArray *pObjectArray = nullptr;
	HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);
	int found = -1;

	if (SUCCEEDED(hr))
	{
		UINT count;
		hr = pObjectArray->GetCount(&count);

		if (SUCCEEDED(hr))
		{
			for (UINT i = 0; i < count; i++)
			{
				IVirtualDesktop *pDesktop = nullptr;

				if (FAILED(pObjectArray->GetAt(i, __uuidof(IVirtualDesktop), (void**)&pDesktop)))
					continue;

				GUID id = { 0 };
				if (SUCCEEDED(pDesktop->GetID(&id)) && id == desktopId)
				{
					found = i;
					pDesktop->Release();
					break;
				}

				pDesktop->Release();
			}
		}

		pObjectArray->Release();
	}

	return found;
}

IVirtualDesktop* _GetDesktopByNumber(int number) {
	_RegisterService();

	IObjectArray *pObjectArray = nullptr;
	HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);
	IVirtualDesktop* found = nullptr;

	if (SUCCEEDED(hr))
	{
		UINT count;
		hr = pObjectArray->GetCount(&count);
		pObjectArray->GetAt(number, __uuidof(IVirtualDesktop), (void**)&found);
		pObjectArray->Release();
	}

	return found;
}

GUID DllExport GetWindowDesktopId(HWND window) {
	_RegisterService();

	GUID pDesktopId = {};
	pDesktopManager->GetWindowDesktopId(window, &pDesktopId);

	return pDesktopId;
}

int DllExport GetWindowDesktopNumber(HWND window) {
	_RegisterService();

	GUID* pDesktopId = new GUID({ 0 });
	if (SUCCEEDED(pDesktopManager->GetWindowDesktopId(window, pDesktopId))) {
		return GetDesktopNumberById(*pDesktopId);
	}

	return -1;
}

int DllExport IsWindowOnCurrentVirtualDesktop(HWND window) {
	_RegisterService();

	BOOL b;
	if (SUCCEEDED(pDesktopManager->IsWindowOnCurrentVirtualDesktop(window, &b))) {
		return b;
	}

	return -1;
}

GUID DllExport GetDesktopIdByNumber(int number) {
	GUID id;
	IVirtualDesktop* pDesktop = _GetDesktopByNumber(number);
	if (pDesktop != nullptr) {
		pDesktop->GetID(&id);
	}
	return id;
}


int DllExport IsWindowOnDesktopNumber(HWND window, int number) {
	_RegisterService();
	IApplicationView* app = nullptr;
	if (window == 0) {
		return -1;
	}
	viewCollection->GetViewForHwnd(window, &app);
	GUID desktopId = { 0 };
	app->GetVirtualDesktopId(&desktopId);
	GUID desktopCheckId = GetDesktopIdByNumber(number);
	app->Release();
	if (desktopCheckId == GUID_NULL || desktopId == GUID_NULL) {
		return -1;
	}

	if (GetDesktopIdByNumber(number) == desktopId) {
		return 1;
	}
	else {
		return 0;
	}
	
	return -1;
}

BOOL DllExport MoveWindowToDesktopNumber(HWND window, int number) {
	_RegisterService();
	IVirtualDesktop* pDesktop = _GetDesktopByNumber(number);
	if (pDesktopManager == nullptr) {
		std::wcout << L"ARRGH?";
		return false;
	}
	if (window == 0) {
		return false;
	}
	if (pDesktop != nullptr) {
		GUID id = { 0 };
		if (SUCCEEDED(pDesktop->GetID(&id))) {
			IApplicationView* app = nullptr;
			viewCollection->GetViewForHwnd(window, &app);
			if (app != nullptr) {
				pDesktopManagerInternal->MoveViewToDesktop(app, pDesktop);
				return true;
			}
		}
	}
	return false;
}

int DllExport GetDesktopNumber(IVirtualDesktop *pDesktop) {
	_RegisterService();

	if (pDesktop == nullptr) {
		return -1;
	}

	GUID guid;

	if (SUCCEEDED(pDesktop->GetID(&guid))) {
		return GetDesktopNumberById(guid);
	}

	return -1;
}
IVirtualDesktop* GetCurrentDesktop() {
	_RegisterService();

	if (pDesktopManagerInternal == nullptr) {
		return nullptr;
	}
	IVirtualDesktop* found = nullptr;
	pDesktopManagerInternal->GetCurrentDesktop(&found);
	return found;
}

int DllExport GetCurrentDesktopNumber() {
	IVirtualDesktop* virtualDesktop = GetCurrentDesktop();
	int number = GetDesktopNumber(virtualDesktop);
	virtualDesktop->Release();
	return number;
}

void DllExport GoToDesktopNumber(int number) {
	_RegisterService();

	if (pDesktopManagerInternal == nullptr) {
		return;
	}

	IVirtualDesktop* oldDesktop = GetCurrentDesktop();
	GUID oldId = { 0 };
	oldDesktop->GetID(&oldId);
	oldDesktop->Release();

	IObjectArray *pObjectArray = nullptr;
	HRESULT hr = pDesktopManagerInternal->GetDesktops(&pObjectArray);
	int found = -1;
	ChangeDesktopAction act;

	if (SUCCEEDED(hr))
	{
		UINT count;
		hr = pObjectArray->GetCount(&count);

		if (SUCCEEDED(hr))
		{
			for (UINT i = 0; i < count; i++)
			{
				IVirtualDesktop *pDesktop = nullptr;

				if (FAILED(pObjectArray->GetAt(i, __uuidof(IVirtualDesktop), (void**)&pDesktop)))
					continue;

				GUID id = { 0 };
				pDesktop->GetID(&id);
				if (i == number) {
					if (_keepMinimized) {
						_changingDesktop = true;
						act.oldDesktopGuid = oldId;
						act.newDesktopGuid = id;
						_ChangeDesktop_ListWins(act);
						_ChangeDesktop_ShowNew();
					}
					pDesktopManagerInternal->SwitchDesktop(pDesktop);
				}

				pDesktop->Release();
			}
		}
		pObjectArray->Release();
	}
}

struct ShowWindowOnDesktopAction {
	int desktopNumber;
	int cmdShow;
};
//
//__inline BOOL CALLBACK
//_EnumWindowProc_ShowWindowOnDesktopAsync(HWND hwnd, LPARAM lParam)
//{
//	ShowWindowOnDesktopAction *act = (ShowWindowOnDesktopAction *) lParam;
//	if (act == nullptr) {
//		return TRUE;
//	}
//	IApplicationView* view = GetApplicationViewForHwnd(hwnd);
//	if (view != nullptr) {
//		GUID desktopId;
//		if (SUCCEEDED(view->GetVirtualDesktopId(&desktopId))) {
//			int deskNum = GetDesktopNumberById(desktopId);
//			std::wcout << "Window " << hwnd << " desk" << deskNum << " hide " << act->desktopNumber << "\r\n";
//			if (deskNum == act->desktopNumber) {
//				ShowWindowAsync(hwnd, act->cmdShow);
//			}
//		}
//	}
//	return TRUE;
//}



//void ShowWindowOnDesktopAsync(int desktopNumber, int cmdShow) {
//	ShowWindowOnDesktopAction act;
//	act.cmdShow = cmdShow;
//	act.desktopNumber = desktopNumber;
//	EnumWindows(_EnumWindowProc_ShowWindowOnDesktopAsync, (LPARAM) &act);
//}
//
//void _temp_enumAll() {
//	ShowWindowOnDesktopAsync(6, SW_MINIMIZE);
//}
//
//void _temp_GetViews() {
//	_RegisterService();
//	IObjectArray *array;
//	UINT count;
//	if (SUCCEEDED(viewCollection->GetViewsByZOrder(&array))) {
//		if (SUCCEEDED(array->GetCount(&count))) {
//			for (int i = 0; i < count; i++)
//			{
//				IApplicationView *view;
//				if (SUCCEEDED(array->GetAt(i, IID_IApplicationView, (void**)&view))) {
//					GUID desktopId;
//					// BOOL isTray;
//					// view->IsTray(&isTray);
//					if (SUCCEEDED(view->GetVirtualDesktopId(&desktopId))) {
//						int deskNum = GetDesktopNumberById(desktopId);
//						std::wcout << "On Desktop: " << deskNum << "\r\n";
//						if (deskNum == 6) {
//							UINT state;
//							int vis;
//							view->GetViewState(&state);
//							view->GetVisibility(&vis);
//
//							std::wcout << "State: " << state << " " << vis << "\r\n";
//						}
//					}
//				}
//			}
//		}
//	}
//
//}

LPWSTR _GetApplicationIdForHwnd(HWND hwnd) {
	if (hwnd == 0)
		return nullptr;
	IApplicationView* app = _GetApplicationViewForHwnd(hwnd);
	if (app != nullptr) {
		LPWSTR appId = new TCHAR[1024];
		app->GetAppUserModelId(&appId);
		app->Release();
		return appId;
	}
	return nullptr;
}

int DllExport IsPinnedWindow(HWND hwnd) {
	if (hwnd == 0)
		return -1;
	_RegisterService();
	IApplicationView* pView = _GetApplicationViewForHwnd(hwnd);
	BOOL isPinned = false;
	if (pView != nullptr) {
		pinnedApps->IsViewPinned(pView, &isPinned);
		pView->Release();
		if (isPinned) {
			return 1;
		}
		else {
			return 0;
		}
	}

	return -1;
}

void DllExport PinWindow(HWND hwnd) {
	if (hwnd == 0)
		return;
	_RegisterService();
	IApplicationView* pView = _GetApplicationViewForHwnd(hwnd);
	if (pView != nullptr) {
		pinnedApps->PinView(pView);
		pView->Release();
	}
}

void DllExport UnPinWindow(HWND hwnd) {
	if (hwnd == 0)
		return;
	_RegisterService();
	IApplicationView* pView = _GetApplicationViewForHwnd(hwnd);
	if (pView != nullptr) {
		pinnedApps->UnpinView(pView);
		pView->Release();
	}
}

int DllExport IsPinnedApp(HWND hwnd) {
	if (hwnd == 0)
		return -1;
	_RegisterService();
	LPWSTR appId = _GetApplicationIdForHwnd(hwnd);
	if (appId != nullptr) {
		BOOL isPinned = false;
		pinnedApps->IsAppIdPinned(appId, &isPinned);
		if (isPinned) {
			return 1;
		}
		else {
			return 0;
		}
	}
	return -1;
}

void DllExport PinApp(HWND hwnd) {
	if (hwnd == 0)
		return;
	_RegisterService();
	LPWSTR appId = _GetApplicationIdForHwnd(hwnd);
	if (appId != nullptr) {
		pinnedApps->PinAppID(appId);
	}
}

void DllExport UnPinApp(HWND hwnd) {
	if (hwnd == 0)
		return;
	_RegisterService();
	LPWSTR appId = _GetApplicationIdForHwnd(hwnd);
	if (appId != nullptr) {
		pinnedApps->UnpinAppID(appId);
	}
}

class _Notifications : public IVirtualDesktopNotification {
private:
	ULONG _referenceCount;
public:
	// Inherited via IVirtualDesktopNotification
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void ** ppvObject) override
	{
		// Always set out parameter to NULL, validating it first.
		if (!ppvObject)
			return E_INVALIDARG;
		*ppvObject = NULL;

		if (riid == IID_IUnknown || riid == IID_IVirtualDesktopNotification)
		{
			// Increment the reference count and return the pointer.
			*ppvObject = (LPVOID)this;
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}
	virtual ULONG STDMETHODCALLTYPE AddRef() override
	{
		return InterlockedIncrement(&_referenceCount);
	}

	virtual ULONG STDMETHODCALLTYPE Release() override
	{
		ULONG result = InterlockedDecrement(&_referenceCount);
		if (result == 0)
		{
			delete this;
		}
		return 0;
	}
	virtual HRESULT STDMETHODCALLTYPE VirtualDesktopCreated(IVirtualDesktop * pDesktop) override
	{
		_PostMessageToListeners(VDA_VirtualDesktopCreated, GetDesktopNumber(pDesktop), 0);
		return S_OK;
	}
	virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyBegin(IVirtualDesktop * pDesktopDestroyed, IVirtualDesktop * pDesktopFallback) override
	{
		_PostMessageToListeners(VDA_VirtualDesktopDestroyBegin, GetDesktopNumber(pDesktopDestroyed), GetDesktopNumber(pDesktopFallback));
		return S_OK;
	}
	virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyFailed(IVirtualDesktop * pDesktopDestroyed, IVirtualDesktop * pDesktopFallback) override
	{
		_PostMessageToListeners(VDA_VirtualDesktopDestroyFailed, GetDesktopNumber(pDesktopDestroyed), GetDesktopNumber(pDesktopFallback));
		return S_OK;
	}
	virtual HRESULT STDMETHODCALLTYPE VirtualDesktopDestroyed(IVirtualDesktop * pDesktopDestroyed, IVirtualDesktop * pDesktopFallback) override
	{
		_PostMessageToListeners(VDA_VirtualDesktopDestroyed, GetDesktopNumber(pDesktopDestroyed), GetDesktopNumber(pDesktopFallback));
		return S_OK;
	}
	virtual HRESULT STDMETHODCALLTYPE ViewVirtualDesktopChanged(IApplicationView * pView) override
	{
		_PostMessageToListeners(VDA_ViewVirtualDesktopChanged, 0, 0);
		return S_OK;
	}
	virtual HRESULT STDMETHODCALLTYPE CurrentVirtualDesktopChanged(
		IVirtualDesktop *pDesktopOld,
		IVirtualDesktop *pDesktopNew) override
	{
		viewCollection->RefreshCollection();
		ChangeDesktopAction act;
		if (pDesktopOld != nullptr) {
			pDesktopOld->GetID(&act.oldDesktopGuid);
		}
		if (pDesktopNew != nullptr) {
			pDesktopNew->GetID(&act.newDesktopGuid);
		}

		// This happens at times
		if (act.oldDesktopGuid != act.newDesktopGuid && _keepMinimized) {
			if (!_changingDesktop) {
				_ChangeDesktop_ListWins(act);
				_ChangeDesktop_HideOld();
				_ChangeDesktop_ShowNew();
			}
			else {
				_ChangeDesktop_HideOld(true);
				_changingDesktop = false;
			}
		}
		_PostMessageToListeners(VDA_CurrentVirtualDesktopChanged, GetDesktopNumberById(act.oldDesktopGuid), GetDesktopNumberById(act.newDesktopGuid));
		return S_OK;
	}
};

void _RegisterDesktopNotifications() {
	_RegisterService();
	if (pDesktopNotificationService == nullptr) {
		return;
	}
	if (registeredForNotifications) {
		return;
	}
	_Notifications *nf = new _Notifications();
	HRESULT res = pDesktopNotificationService->Register(nf, &idNotificationService);
	if (SUCCEEDED(res)) {
		registeredForNotifications = TRUE;
	}
}

void DllExport RestartVirtualDesktopAccessor() {
	_RegisterService(TRUE);
	_RegisterDesktopNotifications();
}

void DllExport RegisterPostMessageHook(HWND listener, int messageOffset) {
	_RegisterService();

	listeners.insert(std::pair<HWND, int>(listener, messageOffset));
	if (listeners.size() != 1) {
		return;
	}
	_RegisterDesktopNotifications();
}

void DllExport UnregisterPostMessageHook(HWND hwnd) {
	_RegisterService();

	listeners.erase(hwnd);
	if (listeners.size() != 0) {
		return;
	}

	if (pDesktopNotificationService == nullptr) {
		return;
	}

	if (idNotificationService > 0) {
		registeredForNotifications = TRUE;
		pDesktopNotificationService->Unregister(idNotificationService);
	}
}

//HINSTANCE  _wndProdHModule;
//LRESULT CALLBACK _DllWindowProc(HWND, UINT, WPARAM, LPARAM);
//UINT _wmTaskbarCreated;
//
//// Register our windows Class
//BOOL _RegisterDLLWindowClass(wchar_t szClassName[])
//{
//	WNDCLASSEX wc;
//	wc.hInstance = _wndProdHModule;
//	wc.lpszClassName = (LPCWSTR)L"InjectedDLLWindowClass";
//	wc.lpszClassName = (LPCWSTR)szClassName;
//	wc.lpfnWndProc = _DllWindowProc;
//	wc.style = CS_DBLCLKS;
//	wc.cbSize = sizeof(WNDCLASSEX);
//	wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
//	wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);
//	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
//	wc.lpszMenuName = NULL;
//	wc.cbClsExtra = 0;
//	wc.cbWndExtra = 0;
//	wc.hbrBackground = (HBRUSH)COLOR_BACKGROUND;
//	if (!RegisterClassEx(&wc))
//		return 0;
//}
//
//// The new thread
//DWORD WINAPI _DllThreadProc(LPVOID lpParam)
//{
//	MSG messages;
//	_RegisterDLLWindowClass(L"VirtualDesktopAccessorListener");
//	HWND hwnd = CreateWindowEx(0, L"VirtualDesktopAccessorListener", NULL, WS_EX_TOOLWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 400, 300, NULL, NULL, _wndProdHModule, NULL);
//	if (hwnd == NULL) {
//		wchar_t buf[256];
//		FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM, NULL, GetLastError(),
//			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), buf, 256, NULL);
//		MessageBox(NULL, buf, _T("Error"), MB_OK);
//	}
//	while (GetMessage(&messages, NULL, 0, 0))
//	{
//		TranslateMessage(&messages);
//		DispatchMessage(&messages);
//	}
//	return 1;
//}
//
//LRESULT CALLBACK _DllWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
//{
//	switch (message)
//	{
//	case WM_CREATE:
//		_wmTaskbarCreated = RegisterWindowMessage(_T("TaskbarCreated"));
//		MessageBox(NULL, _T("Created"), _T(""), MB_OK);
//	case WM_COMMAND:
//		break;
//	case WM_DESTROY:
//		PostQuitMessage(0);
//		break;
//	default:
//		if (message == _wmTaskbarCreated) {
//			MessageBox(NULL, _T("Taskbar Created"), _T(""), MB_OK);
//			RestartVirtualDesktopAccessor();
//		}
//		return DefWindowProc(hwnd, message, wParam, lParam);
//	}
//	return 0;
//}
//
VOID _OpenDllWindow(HINSTANCE injModule) {
	//_wndProdHModule = injModule;
	//CreateThread(0, NULL, _DllThreadProc, NULL, NULL, NULL);
}