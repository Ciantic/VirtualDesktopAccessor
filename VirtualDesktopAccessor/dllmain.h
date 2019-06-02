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

DWORD idNotificationService = 0;

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

void DllExport EnableKeepMinimized() {
	//_keepMinimized = true;
}

void DllExport RestoreMinimized() {
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

	GUID pDesktopId = {};
	if (SUCCEEDED(pDesktopManager->GetWindowDesktopId(window, &pDesktopId))) {
		return GetDesktopNumberById(pDesktopId);
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
		pDesktop->Release();
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

LPWSTR _GetApplicationIdForHwnd(HWND hwnd) {
	// TODO: This should not return a pointer, it should take in a pointer, or return either wstring or std::string

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

int DllExport ViewIsShownInSwitchers(HWND hwnd) {

	//// Iterate views for fun
	//IObjectArray* arr = nullptr;
	//UINT count;
	//viewCollection->GetViews(&arr);
	//arr->GetCount(&count);

	//for (int i = 0; i < count; i++)
	//{
	//	IApplicationView* app2;
	//	HRESULT getAtResult = arr->GetAt(i, IID_IApplicationView, (void**)&app2);
	//	if (app2 != nullptr && getAtResult == S_OK) {
	//		PWSTR modelId;
	//		app2->GetAppUserModelId(&modelId);

	//		BOOL showInSwitchers = 0;
	//		app2->GetShowInSwitchers(&showInSwitchers);

	//		BOOL isVisible = 0;
	//		app2->GetVisibility(&isVisible);

	//		int unknown1 = 0;
	//		HRESULT res1 = app2->Unknown1(&unknown1);
	//		int unknown2 = 0;
	//		HRESULT res2 = app2->Unknown2(&unknown2);
	//		int unknown3 = 0;
	//		HRESULT res3 = app2->Unknown3(&unknown3);
	//		int unknown5 = 0;
	//		HRESULT res5 = app2->Unknown5(&unknown5);
	//		int unknown8 = 0;
	//		HRESULT res8 = app2->Unknown8(&unknown8);

	//		/* E_NOTIMPL
	//		BOOL isInHighZOrderBand = 0;
	//		HRESULT zres = app2->IsInHighZOrderBand(&isInHighZOrderBand);
	//		*/

	//		/* Access violation
	//		BOOL isTray = 0;
	//		HRESULT isTrayRes = app2->IsTray(&isTray);
	//		*/

	//		wprintf(L"modelId: %s switcher: %d visible: %d  %d %d %d %d %d\n", modelId, showInSwitchers, isVisible, unknown1, unknown2, unknown3, unknown5, unknown8);

	//		/* Seems to be always nullptr
	//		HSTRING className;
	//		app2->GetRuntimeClassName(&className);
	//		*/

	//		/*
	//		Seems to be always 0xcccccccc00000000
	//		IApplicationView* app2Owner;
	//		if (app2->GetRootSwitchableOwner(&app2Owner) == S_OK && app2Owner != (IApplicationView*) 0xcccccccc00000000) {
	//			PWSTR modelIdOwner;
	//			app2Owner->GetAppUserModelId(&modelIdOwner);
	//			wprintf(L"modelId owner: %s \n", modelIdOwner);
	//			app2Owner->Release();
	//		}
	//		*/
	//	}
	//	app2->Release();
	//}


	_RegisterService();
	IApplicationView* view = _GetApplicationViewForHwnd(hwnd);
	int result = -1;
	if (view != nullptr) {
		BOOL show = 0;
		if (view->GetShowInSwitchers(&show) == S_OK) {
			result = show;
		}
		view->Release();
	}
	return result;
}

int DllExport ViewIsVisible(HWND hwnd) {
	_RegisterService();
	IApplicationView* view = _GetApplicationViewForHwnd(hwnd);
	int result = -1;
	if (view != nullptr) {
		int show = 0;
		if (view->GetVisibility(&show) == S_OK) {
			result = show;
		}
		view->Release();
	}
	return result;
}

HWND DllExport ViewGetThumbnailHwnd(HWND hwnd) {
	_RegisterService();
	IApplicationView* view = _GetApplicationViewForHwnd(hwnd);
	HWND result = 0;
	if (view != nullptr) {
		if (view->GetThumbnailWindow(&result) != S_OK) {
			result = 0;
		}
		view->Release();
	}
	return result;
}

HRESULT DllExport ViewSetFocus(HWND hwnd) {
	_RegisterService();
	IApplicationView* view = _GetApplicationViewForHwnd(hwnd);
	HRESULT result = -1;
	if (view != nullptr) {
		result = view->SetFocus();
		view->Release();
	}
	return result;
}

HWND DllExport ViewGetFocused() {
	_RegisterService();
	IApplicationView* view;
	HRESULT getAtResult = viewCollection->GetViewInFocus(&view);
	HWND ret = 0;
	if (view != nullptr && getAtResult == S_OK) {
		HWND wnd = 0;
		if (view->GetThumbnailWindow(&wnd) == S_OK && wnd != 0) {
			ret = wnd;
		}
		view->Release();
	}
	return ret;
}

HRESULT DllExport ViewSwitchTo(HWND hwnd) {
	_RegisterService();
	IApplicationView* view = _GetApplicationViewForHwnd(hwnd);
	HRESULT result = -1;
	if (view != nullptr) {
		result = view->SwitchTo();
		view->Release();
	}
	return result;
}

UINT DllExport ViewGetByZOrder(HWND *windows, UINT count, BOOL onlySwitcherWindows, BOOL onlyCurrentDesktop) {
	_RegisterService();
	IObjectArray* arr = nullptr;
	UINT countViews;
	IApplicationView* view;
	int fill = 0;
	if (viewCollection->GetViewsByZOrder(&arr) != S_OK) {
		return 0;
	}
	arr->GetCount(&countViews);
	if (countViews > count) {
		arr->Release();
		return 0;
	}

	for (UINT i = 0; i < count; i++)
	{
		HRESULT getAtResult = arr->GetAt(i - 1, IID_IApplicationView, (void**)&view);
		
		if (view != nullptr && getAtResult == S_OK) {
			HWND wnd = 0;
			BOOL showInSwitchers = false;
			BOOL isOnCurrentDesktop = false;
			if (onlySwitcherWindows && (view->GetShowInSwitchers(&showInSwitchers) != S_OK || !showInSwitchers)) {
				view->Release();
				continue;
			}
			if (view->GetThumbnailWindow(&wnd) != S_OK || wnd == 0) {
				view->Release();
				continue;
			}
			if (onlyCurrentDesktop && (pDesktopManager->IsWindowOnCurrentVirtualDesktop(wnd, &isOnCurrentDesktop) != S_OK || !isOnCurrentDesktop)) {
				view->Release();
				continue;
			}
			windows[fill] = wnd;
			fill++;
			view->Release();
		}
	}
	arr->Release();
	return fill;
}

struct TempWindowEntry {
	HWND hwnd;
	ULONGLONG lastActivationTimestamp;
};

UINT DllExport ViewGetByLastActivationOrder(HWND *windows, UINT count, BOOL onlySwitcherWindows, BOOL onlyCurrentDesktop) {
	_RegisterService();
	IObjectArray* arr = nullptr;
	UINT countViews;
	IApplicationView* view;
	if (viewCollection->GetViews(&arr) != S_OK) {
		return 0;
	}
	arr->GetCount(&countViews);
	if (countViews > count) {
		arr->Release();
		return 0;
	}

	std::vector<TempWindowEntry> unsorted;
	for (UINT i = 0; i < count; i++)
	{
		HRESULT getAtResult = arr->GetAt(i - 1, IID_IApplicationView, (void**)&view);
		if (view != nullptr && getAtResult == S_OK) {
			HWND wnd = 0;
			ULONGLONG lastActivationTimestamp = 0;
			BOOL showInSwitchers = false;
			BOOL isOnCurrentDesktop = false;

			if (onlySwitcherWindows && (view->GetShowInSwitchers(&showInSwitchers) != S_OK || !showInSwitchers)) {
				view->Release();
				continue;
			}
			if (view->GetThumbnailWindow(&wnd) != S_OK || wnd == 0) {
				view->Release();
				continue;
			}

			if (onlyCurrentDesktop && (pDesktopManager->IsWindowOnCurrentVirtualDesktop(wnd, &isOnCurrentDesktop) != S_OK || !isOnCurrentDesktop)) {
				view->Release();
				continue;
			}

			if (view->GetLastActivationTimestamp(&lastActivationTimestamp) != S_OK) {
				view->Release();
				continue;
			}
			TempWindowEntry entry;
			entry.hwnd = wnd;
			entry.lastActivationTimestamp = lastActivationTimestamp;
			unsorted.push_back(entry);
			view->Release();
		}
	}
	arr->Release();

	std::sort(unsorted.begin(), unsorted.end(), [](auto const& lhs, auto const& rhs) {
		return lhs.lastActivationTimestamp > rhs.lastActivationTimestamp;
	});

	UINT i = 0;
	for (auto entry : unsorted) {
		windows[i] = entry.hwnd;
		i++;
	}
	
	return i;
}

ULONGLONG DllExport ViewGetLastActivationTimestamp(HWND hwnd) {
	_RegisterService();
	IApplicationView* view = _GetApplicationViewForHwnd(hwnd);
	ULONGLONG result = 0;
	if (view != nullptr) {
		if (view->GetLastActivationTimestamp(&result) != S_OK) {
			result = 0;
		}
		view->Release();
	}
	return result;
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

	// TODO: This is never deleted
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

VOID _OpenDllWindow(HINSTANCE injModule) {
}