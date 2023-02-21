use crate::log::log_output;

/// Purpose of this module is to provide helpers to access functions in interfaces module, not for direct consumption
///
/// All functions here either take in a reference to an interface or initializes a com interace.
use super::interfaces::*;
use super::Result;
use std::convert::TryFrom;
use std::mem::ManuallyDrop;
use std::rc::Rc;
use std::{cell::RefCell, ffi::c_void};
use windows::core::HRESULT;
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::System::Com::CoUninitialize;
use windows::Win32::System::Com::CLSCTX_LOCAL_SERVER;
use windows::Win32::System::Com::COINIT_APARTMENTTHREADED;
use windows::Win32::System::Threading::GetCurrentThread;
use windows::Win32::System::Threading::SetThreadPriority;
use windows::Win32::System::Threading::THREAD_PRIORITY_TIME_CRITICAL;
use windows::{
    core::{Interface, Vtable, GUID, HSTRING},
    Win32::{System::Com::CoCreateInstance, UI::Shell::Common::IObjectArray},
};

type WCHAR = u16;
type APPIDPWSTR = *const WCHAR;

#[derive(Debug, PartialEq, Clone)]
pub enum Error {
    /// Window is not found
    WindowNotFound,

    /// Desktop with given ID is not found
    DesktopNotFound,

    /// Creationg of desktop failed
    CreateDesktopFailed,

    /// Remove desktop failed
    RemoveDesktopFailed,

    /// Unable to create service, ensure that explorer.exe is running
    ClassNotRegistered,

    /// Unable to connect to service
    RpcServerNotAvailable,

    /// Com object not connected
    ComObjectNotConnected,

    /// Generic element not found
    ComElementNotFound,

    /// Some unhandled COM error
    ComError(HRESULT),

    /// This should not happen, this means that successful COM call allocated a
    /// null pointer, in this case it is an error in the COM service, or it's
    /// usage.
    ComAllocatedNullPtr,

    /// Sender error
    SenderError,

    /// Receiver Error
    ReceiverError,

    /// Listener thread not created
    ListenerThreadIdNotCreated,
}

trait HRESULTHelpers {
    fn as_error(&self) -> Error;
    fn as_result(&self) -> Result<()>;
}

impl HRESULTHelpers for ::windows::core::HRESULT {
    fn as_error(&self) -> Error {
        if self.0 == -2147221164 {
            // 0x80040154
            return Error::ClassNotRegistered;
        }
        if self.0 == -2147023174 {
            // 0x800706BA
            return Error::RpcServerNotAvailable;
        }
        if self.0 == -2147220995 {
            // 0x800401FD
            return Error::ComObjectNotConnected;
        }
        if self.0 == -2147319765 {
            // 0x8002802B
            return Error::ComElementNotFound;
        }
        Error::ComError(self.clone())
    }

    fn as_result(&self) -> Result<()> {
        if self.is_ok() {
            return Ok(());
        }
        Err(self.as_error())
    }
}

impl From<::windows::core::Error> for Error {
    fn from(r: ::windows::core::Error) -> Self {
        r.code().as_error()
    }
}

// From SendError for Error
impl From<std::sync::mpsc::SendError<ComFn>> for Error {
    fn from(_: std::sync::mpsc::SendError<ComFn>) -> Self {
        Error::SenderError
    }
}

// From std::sync::mpsc::RecvError for Error
impl From<std::sync::mpsc::RecvError> for Error {
    fn from(_: std::sync::mpsc::RecvError) -> Self {
        Error::ReceiverError
    }
}

struct ComSta();
impl ComSta {
    fn new() -> Self {
        #[cfg(debug_assertions)]
        log_output("CoInitializeEx COINIT_APARTMENTTHREADED");

        let _ = unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) };
        ComSta()
    }
}
impl Drop for ComSta {
    fn drop(&mut self) {
        #[cfg(debug_assertions)]
        log_output("CoUninitialize");

        unsafe { CoUninitialize() };
    }
}

type ComFn = Box<dyn Fn(&ComObjects) + Send + 'static>;

static WORKER_CHANNEL: once_cell::sync::Lazy<(
    std::sync::mpsc::SyncSender<ComFn>,
    std::thread::JoinHandle<()>,
)> = once_cell::sync::Lazy::new(|| {
    // TODO: Is rendezvous channel correct here? (0 = rendezvous channel)
    let (sender, receiver) = std::sync::mpsc::sync_channel::<ComFn>(0);
    (
        sender,
        std::thread::spawn(move || {
            // Set thread priority to time critical, explorer.exe really hates if your listener thread is slow
            unsafe { SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL) };

            let com = ComObjects::new();
            for f in receiver {
                f(&com);
            }
        }),
    )
});

/// This is a helper function to initialize and run COM related functions in a known thread
///
/// Virtual Desktop COM Objects don't like to being called from different threads rapidly, something goes wrong. This function ensures that all COM calls are done in a single thread.
pub fn with_com_objects<F, T>(f: F) -> Result<T>
where
    F: Fn(&ComObjects) -> Result<T> + Send + 'static,
    T: Send + 'static,
{
    // Oneshot channel
    let (sender, receiver) = std::sync::mpsc::channel();

    WORKER_CHANNEL.0.send(Box::new(move |c| {
        // Retry the function up to 5 times if it gives an error
        let mut r = f(c);
        for _ in 0..5 {
            match &r {
                Err(er)
                    if er == &Error::ClassNotRegistered
                        || er == &Error::RpcServerNotAvailable
                        || er == &Error::ComObjectNotConnected
                        || er == &Error::ComAllocatedNullPtr =>
                {
                    #[cfg(debug_assertions)]
                    log_output(&format!("Retry the function after {:?}", er));

                    // Explorer.exe has mostlikely crashed, retry the function
                    c.drop_services();
                    r = f(c);
                    continue;
                }
                other => {
                    // Show the error
                    #[cfg(debug_assertions)]
                    if let Err(er) = &other {
                        log_output(&format!("with_com_objects failed with {:?}", er));
                    }

                    // Return the Result
                    break;
                }
            }
        }
        let send_result = sender.send(r);
        if let Err(e) = send_result {
            #[cfg(debug_assertions)]
            log_output(&format!("with_com_objects failed to send result {:?}", e));
        }
    }))?;

    receiver.recv()?
}

#[derive(Copy, Clone, Debug)]
pub enum DesktopInternal {
    Index(u32),
    Guid(GUID),
    IndexGuid(u32, GUID),
}

// Impl from u32 to DesktopTest
impl From<u32> for DesktopInternal {
    fn from(index: u32) -> Self {
        DesktopInternal::Index(index)
    }
}

// Impl from i32 to DesktopTest
impl From<i32> for DesktopInternal {
    fn from(index: i32) -> Self {
        DesktopInternal::Index(index as u32)
    }
}

// Impl from GUID to DesktopTest
impl From<GUID> for DesktopInternal {
    fn from(guid: GUID) -> Self {
        DesktopInternal::Guid(guid)
    }
}

// Impl from &GUID to DesktopTest
impl From<&GUID> for DesktopInternal {
    fn from(guid: &GUID) -> Self {
        DesktopInternal::Guid(*guid)
    }
}

impl<'a> TryFrom<&'a IVirtualDesktop> for DesktopInternal {
    type Error = Error;

    fn try_from(desktop: &'a IVirtualDesktop) -> Result<Self> {
        let mut guid = GUID::default();
        unsafe { desktop.get_id(&mut guid).as_result()? }
        Ok(DesktopInternal::Guid(guid))
    }
}
impl<'a> TryFrom<&'a ManuallyDrop<IVirtualDesktop>> for DesktopInternal {
    type Error = Error;

    fn try_from(desktop: &'a ManuallyDrop<IVirtualDesktop>) -> Result<Self> {
        let mut guid = GUID::default();
        unsafe { desktop.get_id(&mut guid).as_result()? }
        Ok(DesktopInternal::Guid(guid))
    }
}

pub struct ComObjects {
    provider: RefCell<Option<Rc<IServiceProvider>>>,
    manager: RefCell<Option<Rc<IVirtualDesktopManager>>>,
    manager_internal: RefCell<Option<Rc<IVirtualDesktopManagerInternal>>>,

    #[allow(dead_code)]
    notification_service: RefCell<Option<Rc<IVirtualDesktopNotificationService>>>,
    pinned_apps: RefCell<Option<Rc<IVirtualDesktopPinnedApps>>>,
    view_collection: RefCell<Option<Rc<IApplicationViewCollection>>>,

    // Order is important, this must be dropped last
    #[allow(dead_code)]
    com_sta: ComSta,
}

impl ComObjects {
    pub fn new() -> Self {
        Self {
            provider: RefCell::new(None),
            manager: RefCell::new(None),
            manager_internal: RefCell::new(None),
            notification_service: RefCell::new(None),
            pinned_apps: RefCell::new(None),
            view_collection: RefCell::new(None),
            com_sta: ComSta::new(),
        }
    }

    fn get_provider(&self) -> Result<Rc<IServiceProvider>> {
        let mut provider = self.provider.borrow_mut();
        if provider.is_none() {
            let new_provider = Rc::new(unsafe {
                CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_LOCAL_SERVER)?
            });
            *provider = Some(new_provider);
        }

        provider
            .as_ref()
            .map(|v| Rc::clone(&v))
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_manager(&self) -> Result<Rc<IVirtualDesktopManager>> {
        let mut manager = self.manager.borrow_mut();
        if manager.is_none() {
            let provider = self.get_provider()?;
            let mut obj = std::ptr::null_mut::<c_void>();
            unsafe {
                provider
                    .query_service(
                        &IVirtualDesktopManager::IID,
                        &IVirtualDesktopManager::IID,
                        &mut obj,
                    )
                    .as_result()?;
            }
            assert_eq!(obj.is_null(), false);
            *manager = Some(Rc::new(unsafe { IVirtualDesktopManager::from_raw(obj) }));
        }
        manager
            .as_ref()
            .map(|v| Rc::clone(&v))
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_manager_internal(&self) -> Result<Rc<IVirtualDesktopManagerInternal>> {
        let mut manager_internal = self.manager_internal.borrow_mut();
        if manager_internal.is_none() {
            let provider = self.get_provider()?;
            let mut obj = std::ptr::null_mut::<c_void>();
            unsafe {
                provider
                    .query_service(
                        &CLSID_VirtualDesktopManagerInternal,
                        &IVirtualDesktopManagerInternal::IID,
                        &mut obj,
                    )
                    .as_result()?;
            }
            assert_eq!(obj.is_null(), false);
            *manager_internal = Some(Rc::new(unsafe {
                IVirtualDesktopManagerInternal::from_raw(obj)
            }));
        }
        manager_internal
            .as_ref()
            .map(|v| Rc::clone(&v))
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_notification_service(&self) -> Result<Rc<IVirtualDesktopNotificationService>> {
        let mut notification_service = self.notification_service.borrow_mut();
        if notification_service.is_none() {
            let provider = self.get_provider()?;
            let mut obj = std::ptr::null_mut::<c_void>();
            unsafe {
                provider
                    .query_service(
                        &CLSID_IVirtualNotificationService,
                        &IVirtualDesktopNotificationService::IID,
                        &mut obj,
                    )
                    .as_result()?;
            }
            assert_eq!(obj.is_null(), false);
            *notification_service = Some(Rc::new(unsafe {
                IVirtualDesktopNotificationService::from_raw(obj)
            }));
        }
        notification_service
            .as_ref()
            .map(|v| Rc::clone(&v))
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_pinned_apps(&self) -> Result<Rc<IVirtualDesktopPinnedApps>> {
        let mut pinned_apps = self.pinned_apps.borrow_mut();
        if pinned_apps.is_none() {
            let provider = self.get_provider()?;
            let mut obj = std::ptr::null_mut::<c_void>();
            unsafe {
                provider
                    .query_service(
                        &CLSID_VirtualDesktopPinnedApps,
                        &IVirtualDesktopPinnedApps::IID,
                        &mut obj,
                    )
                    .as_result()?;
            }
            assert_eq!(obj.is_null(), false);
            *pinned_apps = Some(Rc::new(unsafe { IVirtualDesktopPinnedApps::from_raw(obj) }));
        }
        pinned_apps
            .as_ref()
            .ok_or(Error::ComAllocatedNullPtr)
            .map(|a| Rc::clone(a))
    }

    fn get_view_collection(&self) -> Result<Rc<IApplicationViewCollection>> {
        let mut view_collection = self.view_collection.borrow_mut();
        if view_collection.is_none() {
            let provider = self.get_provider()?;
            let mut obj = std::ptr::null_mut::<c_void>();
            unsafe {
                provider
                    .query_service(
                        &IApplicationViewCollection::IID,
                        &IApplicationViewCollection::IID,
                        &mut obj,
                    )
                    .as_result()?;
            }
            assert_eq!(obj.is_null(), false);
            *view_collection = Some(Rc::new(unsafe {
                IApplicationViewCollection::from_raw(obj)
            }));
        }
        view_collection
            .as_ref()
            .map(|v| Rc::clone(&v))
            .ok_or(Error::ComAllocatedNullPtr)
    }

    pub fn drop_services(&self) {
        self.provider.borrow_mut().take();
        self.manager.borrow_mut().take();
        self.manager_internal.borrow_mut().take();
        self.notification_service.borrow_mut().take();
        self.pinned_apps.borrow_mut().take();
        self.view_collection.borrow_mut().take();
    }

    pub(crate) fn is_connected(&self) -> bool {
        // TODO: What is a best way to check if service is connected?

        // Calling any method yields an error if service is not connected.
        //
        // I call get_count method, if it's well implemented it should be just
        // like returning a value, not allocating anything.
        match self.get_manager_internal() {
            Ok(manager_internal) => {
                let mut out_count = 0;
                let res = unsafe {
                    manager_internal
                        .get_desktop_count(0, &mut out_count)
                        .as_result()
                };

                #[cfg(debug_assertions)]
                if let Err(er) = &res {
                    log_output(&format!("is connected error: {:?} {}", er, out_count));
                }

                if out_count == 0 {
                    return false;
                }
                match res {
                    Ok(_) => true,
                    Err(_) => false,
                }
            }
            Err(_) => false,
        }
    }

    fn get_idesktops_array(&self) -> Result<IObjectArray> {
        let mut desktops = None;
        unsafe {
            self.get_manager_internal()?
                .get_desktops(0, &mut desktops)
                .as_result()?
        }
        desktops.ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_desktop_index_by_guid(&self, id: &GUID) -> Result<u32> {
        let desktops = self.get_idesktops_array()?;
        let count = unsafe { desktops.GetCount()? };
        for i in 0..count {
            let desktop_id: GUID = get_idesktop_guid(&unsafe { desktops.GetAt(i)? })?;
            if desktop_id == *id {
                return Ok(i);
            }
        }
        Err(Error::DesktopNotFound)
    }

    fn get_desktop_guid_by_index(&self, id: u32) -> Result<GUID> {
        let desktops = self.get_idesktops_array()?;
        let count = unsafe { desktops.GetCount()? };
        if id >= count {
            return Err(Error::DesktopNotFound);
        }
        get_idesktop_guid(&unsafe { desktops.GetAt(id)? })
    }

    fn get_idesktop(&self, desktop: &DesktopInternal) -> Result<IVirtualDesktop> {
        match desktop {
            DesktopInternal::Index(id) => {
                let desktops = self.get_idesktops_array()?;
                let count = unsafe { desktops.GetCount()? };
                if *id >= count {
                    return Err(Error::DesktopNotFound);
                }
                Ok(unsafe { desktops.GetAt(*id)? })
            }
            DesktopInternal::Guid(id) => {
                let manager = self.get_manager_internal()?;
                let mut desktop = None;
                unsafe {
                    manager.find_desktop(id, &mut desktop).as_result()?;
                }
                desktop.ok_or(Error::DesktopNotFound)
            }
            DesktopInternal::IndexGuid(_, id) => {
                let manager = self.get_manager_internal()?;
                let mut desktop = None;
                unsafe {
                    manager.find_desktop(id, &mut desktop).as_result()?;
                }
                desktop.ok_or(Error::DesktopNotFound)
            }
        }
    }

    fn move_view_to_desktop(
        &self,
        view: &IApplicationView,
        desktop: &DesktopInternal,
    ) -> Result<()> {
        let desktop = self.get_idesktop(desktop)?;
        unsafe {
            self.get_manager_internal()?
                .move_view_to_desktop(ComIn::new(&view), ComIn::new(&desktop))
                .as_result()
                .map_err(|e| {
                    if e == Error::ComElementNotFound {
                        Error::DesktopNotFound
                    } else {
                        e
                    }
                })?
        }
        Ok(())
    }

    fn get_iapplication_view_for_hwnd(&self, hwnd: &HWND) -> Result<IApplicationView> {
        let mut view = None;
        unsafe {
            self.get_view_collection()?
                .get_view_for_hwnd(hwnd.clone(), &mut view)
                .as_result()
                .map_err(|er| {
                    if er == Error::ComElementNotFound {
                        Error::WindowNotFound
                    } else {
                        er
                    }
                })?
        }
        view.ok_or(Error::WindowNotFound)
    }

    pub fn get_desktop_index(&self, id: &DesktopInternal) -> Result<u32> {
        match id {
            DesktopInternal::Index(id) => Ok(*id),
            DesktopInternal::Guid(guid) => self.get_desktop_index_by_guid(guid),
            DesktopInternal::IndexGuid(id, _) => Ok(*id),
        }
    }

    pub fn get_desktop_id(&self, desktop: &DesktopInternal) -> Result<GUID> {
        match desktop {
            DesktopInternal::Index(id) => self.get_desktop_guid_by_index(*id),
            DesktopInternal::Guid(guid) => Ok(*guid),
            DesktopInternal::IndexGuid(_, guid) => Ok(*guid),
        }
    }

    pub fn get_desktops(&self) -> Result<Vec<DesktopInternal>> {
        let desktops = self.get_idesktops_array()?;
        let count = unsafe { desktops.GetCount()? };
        let mut result = Vec::with_capacity(count as usize);
        for i in 0..count {
            let desktop = unsafe { desktops.GetAt(i)? };
            let id = get_idesktop_guid(&desktop)?;
            result.push(DesktopInternal::IndexGuid(i, id));
        }
        Ok(result)
    }

    pub fn register_for_notifications(
        &self,
        notification: &IVirtualDesktopNotification,
    ) -> Result<u32> {
        let notification_service = self.get_notification_service()?;

        unsafe {
            let mut cookie = 0;
            notification_service
                .register(notification.as_raw(), &mut cookie)
                .as_result()
                .map(|_| cookie)
        }
    }

    pub fn unregister_for_notifications(&self, cookie: u32) -> Result<()> {
        let notification_service = self.get_notification_service()?;
        unsafe { notification_service.unregister(cookie).as_result() }
    }

    pub fn switch_desktop(&self, desktop: &DesktopInternal) -> Result<()> {
        let desktop = self.get_idesktop(&desktop)?;
        unsafe {
            self.get_manager_internal()?
                .switch_desktop(0, ComIn::new(&desktop))
                .as_result()?
        }
        Ok(())
    }

    pub fn create_desktop(&self) -> Result<DesktopInternal> {
        let mut desktop = None;
        unsafe {
            self.get_manager_internal()?
                .create_desktop(0, &mut desktop)
                .as_result()?
        }
        let desktop = desktop.ok_or(Error::ComAllocatedNullPtr)?;
        let id = get_idesktop_guid(&desktop)?;
        let index = self.get_desktop_index_by_guid(&id)?;
        Ok(DesktopInternal::IndexGuid(index, id))
    }

    pub fn remove_desktop(
        &self,
        desktop: &DesktopInternal,
        fallback_desktop: &DesktopInternal,
    ) -> Result<()> {
        let desktop = self.get_idesktop(desktop)?;
        let fb_desktop = self.get_idesktop(fallback_desktop)?;
        unsafe {
            self.get_manager_internal()?
                .remove_desktop(ComIn::new(&desktop), ComIn::new(&fb_desktop))
                .as_result()?
        }
        Ok(())
    }

    pub fn is_window_on_desktop(&self, window: &HWND, desktop: &DesktopInternal) -> Result<bool> {
        let desktop_win = self.get_desktop_by_window(window)?;
        Ok(self.get_desktop_id(&desktop_win)? == self.get_desktop_id(&*desktop)?)
    }

    pub fn is_window_on_current_desktop(&self, window: &HWND) -> Result<bool> {
        unsafe {
            let mut value = false;
            self.get_manager()?
                .is_window_on_current_desktop(window.clone(), &mut value)
                .as_result()
                .map_err(|er| match er {
                    // Window does not exist
                    Error::ComElementNotFound => Error::WindowNotFound,
                    _ => er,
                })?;
            Ok(value)
        }
    }

    pub fn move_window_to_desktop(&self, window: &HWND, desktop: &DesktopInternal) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        self.move_view_to_desktop(&view, desktop)
    }

    pub fn get_desktop_count(&self) -> Result<u32> {
        let manager = self.get_manager_internal()?;
        let mut count = 0;
        unsafe {
            manager.get_desktop_count(0, &mut count).as_result()?;
        };
        Ok(count)
    }

    pub fn get_desktop_by_window(&self, window: &HWND) -> Result<DesktopInternal> {
        let mut desktop = GUID::default();
        unsafe {
            self.get_manager()?
                .get_desktop_by_window(window.clone(), &mut desktop)
                .as_result()
                .map_err(|er| match er {
                    // Window does not exist
                    Error::ComElementNotFound => Error::WindowNotFound,
                    _ => er,
                })?
        };
        if desktop == GUID::default() {
            return Err(Error::WindowNotFound);
        }
        Ok(DesktopInternal::Guid(desktop))
    }

    pub fn get_current_desktop(&self) -> Result<DesktopInternal> {
        let mut desktop = None;
        unsafe {
            self.get_manager_internal()?
                .get_current_desktop(0, &mut desktop)
                .as_result()?
        }
        let desktop = desktop.ok_or(Error::ComAllocatedNullPtr)?;
        let id = get_idesktop_guid(&desktop)?;
        Ok(DesktopInternal::Guid(id))
    }

    pub fn is_pinned_window(&self, window: &HWND) -> Result<bool> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        unsafe {
            let mut value = false;
            self.get_pinned_apps()?
                .is_view_pinned(ComIn::new(&view), &mut value)
                .as_result()?;
            Ok(value)
        }
    }

    pub fn pin_window(&self, window: &HWND) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        unsafe {
            self.get_pinned_apps()?
                .pin_view(ComIn::new(&view))
                .as_result()?;
        }
        Ok(())
    }

    pub fn unpin_window(&self, window: &HWND) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        unsafe {
            self.get_pinned_apps()?
                .unpin_view(ComIn::new(&view))
                .as_result()?;
        }
        Ok(())
    }

    fn get_iapplication_id_for_view(&self, view: &IApplicationView) -> Result<APPIDPWSTR> {
        let mut app_id: APPIDPWSTR = std::ptr::null_mut();
        unsafe {
            view.get_app_user_model_id(&mut app_id as *mut _ as *mut _)
                .as_result()?
        }
        Ok(app_id)
    }

    pub fn is_pinned_app(&self, window: &HWND) -> Result<bool> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        let app_id = self.get_iapplication_id_for_view(&view)?;
        unsafe {
            let mut value = false;
            self.get_pinned_apps()?
                .is_app_pinned(app_id, &mut value)
                .as_result()?;
            Ok(value)
        }
    }

    pub fn pin_app(&self, window: &HWND) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        let app_id = self.get_iapplication_id_for_view(&view)?;
        unsafe {
            self.get_pinned_apps()?.pin_app(app_id).as_result()?;
        }
        Ok(())
    }

    pub fn unpin_app(&self, window: &HWND) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        let app_id = self.get_iapplication_id_for_view(&view)?;
        unsafe {
            self.get_pinned_apps()?.unpin_app(app_id).as_result()?;
        }
        Ok(())
    }

    pub fn get_desktop_name(&self, desktop: &DesktopInternal) -> Result<String> {
        let desktop = self.get_idesktop(&desktop)?;
        let mut name = HSTRING::default();
        unsafe {
            desktop.get_name(&mut name).as_result()?;
        }
        Ok(name.to_string())
    }

    pub fn set_desktop_name(&self, desktop: &DesktopInternal, name: &str) -> Result<()> {
        let desktop = self.get_idesktop(&desktop)?;
        let manager_internal = self.get_manager_internal()?;

        unsafe {
            manager_internal
                .set_name(ComIn::new(&desktop), HSTRING::from(name))
                .as_result()
        }
    }

    pub fn get_desktop_wallpaper(&self, desktop: &DesktopInternal) -> Result<String> {
        let desktop = self.get_idesktop(&desktop)?;
        let mut path = HSTRING::default();
        unsafe {
            desktop.get_wallpaper(&mut path).as_result()?;
        }
        Ok(path.to_string())
    }

    pub fn set_desktop_wallpaper(&self, desktop: &DesktopInternal, path: &str) -> Result<()> {
        let manager_internal = self.get_manager_internal()?;
        let desktop = self.get_idesktop(&desktop)?;
        unsafe {
            manager_internal
                .set_wallpaper(ComIn::new(&desktop), HSTRING::from(path))
                .as_result()
        }
    }
}

pub fn get_idesktop_guid(desktop: &IVirtualDesktop) -> Result<GUID> {
    let mut guid = GUID::default();
    unsafe { desktop.get_id(&mut guid).as_result()? }
    Ok(guid)
}
