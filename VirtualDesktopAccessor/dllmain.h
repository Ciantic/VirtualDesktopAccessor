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

std::map<HWND, int> listeners;
IServiceProvider* pServiceProvider = nullptr;
IVirtualDesktopManagerInternal *pDesktopManagerInternal = nullptr;
IVirtualDesktopManager *pDesktopManager = nullptr;
IApplicationViewCollection *viewCollection = nullptr;
IVirtualDesktopPinnedApps *pinnedApps = nullptr;

DWORD idNotificationService = 0;
IVirtualDesktopNotificationService* pDesktopNotificationService = nullptr;

void _PostMessageToListeners(int msgOffset, WPARAM wParam, LPARAM lParam) {
	for each (std::pair<HWND, int> listener in listeners) {
		PostMessage(listener.first, listener.second + msgOffset, wParam, lParam);
	}
}

void _RegisterService() {
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

IVirtualDesktop* GetDesktopByNumber(int number) {
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

BOOL DllExport MoveWindowToDesktopNumber(HWND window, int number) {
	_RegisterService();
	IVirtualDesktop* pDesktop = GetDesktopByNumber(number);
	if (pDesktopManager == nullptr) {
		std::wcout << L"ARRGH?";
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

GUID DllExport GetDesktopIdByNumber(int number) {
	GUID id;
	IVirtualDesktop* pDesktop = GetDesktopByNumber(number);
	if (pDesktop != nullptr) {
		pDesktop->GetID(&id);
	}
	return id;
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
	return GetDesktopNumber(GetCurrentDesktop());
}

void DllExport GoToDesktopNumber(int number) {
	_RegisterService();

	if (pDesktopManagerInternal == nullptr) {
		return;
	}

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
				if (i == number) {
					pDesktopManagerInternal->SwitchDesktop(pDesktop);
				}

				pDesktop->Release();
			}
		}
		pObjectArray->Release();
	}
}

IApplicationView* GetApplicationViewForHwnd(HWND hwnd) {
	IApplicationView* app = nullptr;
	viewCollection->GetViewForHwnd(hwnd, &app);
	return app;
}

LPWSTR GetApplicationIdForHwnd(HWND hwnd) {
	IApplicationView* app = GetApplicationViewForHwnd(hwnd);
	if (app != nullptr) {
		std::wcout << "App view" << app << std::endl;
		LPWSTR appId = new TCHAR[1024];
		app->GetAppUserModelId(&appId);
		std::wcout << "App view ID" << appId << std::endl;
		return appId;
	}
	return nullptr;
}

int DllExport IsPinnedWindow(HWND hwnd) {
	_RegisterService();
	IApplicationView* pView = GetApplicationViewForHwnd(hwnd);
	BOOL isPinned = false;
	if (pView != nullptr) {
		pinnedApps->IsViewPinned(pView, &isPinned);
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
	_RegisterService();
	IApplicationView* pView = GetApplicationViewForHwnd(hwnd);
	if (pView != nullptr) {
		pinnedApps->PinView(pView);
	}
}

void DllExport UnPinWindow(HWND hwnd) {
	_RegisterService();
	IApplicationView* pView = GetApplicationViewForHwnd(hwnd);
	if (pView != nullptr) {
		pinnedApps->UnpinView(pView);
	}
}

/*
void DllExport IsPinnedApp(HWND hwnd) {
	_RegisterService();

}

void DllExport PinApp(HWND hwnd) {
	_RegisterService();
	LPWSTR appId = GetApplicationIdForHwnd(hwnd);
	std::wcout << "Pin app" << appId << std::endl;
	if (appId != nullptr) {
		pinnedApps->PinAppID(appId);
	}
}

void DllExport UnPinApp(HWND hwnd) {

}
*/
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
		_PostMessageToListeners(VDA_CurrentVirtualDesktopChanged, GetDesktopNumber(pDesktopOld), GetDesktopNumber(pDesktopNew));
		return S_OK;
	}
};

void DllExport RegisterPostMessageHook(HWND listener, int messageOffset) {
	_RegisterService();

	listeners.insert(std::pair<HWND, int>(listener, messageOffset));
	if (listeners.size() != 1) {
		return;
	}

	if (pDesktopNotificationService == nullptr) {
		return;
	}

	_Notifications *nf = new _Notifications();
	pDesktopNotificationService->Register(nf, &idNotificationService);
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
		pDesktopNotificationService->Unregister(idNotificationService);
	}
}