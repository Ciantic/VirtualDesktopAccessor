use crate::hresult::HRESULT;
use crate::log_output;

/// Purpose of this module is to provide helpers to access functions in interfaces module, not for direct consumption
///
/// All functions here either take in a reference to an interface or initializes a com interace.
use super::interfaces::*;
use super::Result;
use std::convert::TryFrom;
use std::rc::Rc;
use std::rc::Weak;
use std::{cell::RefCell, ffi::c_void};
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::CoIncrementMTAUsage;
use windows::Win32::System::Com::CoInitializeEx;
use windows::Win32::System::Com::CoUninitialize;
use windows::Win32::System::Com::CLSCTX_LOCAL_SERVER;
use windows::Win32::System::Com::COINIT;
use windows::Win32::System::Com::COINIT_APARTMENTTHREADED;
use windows::Win32::System::Com::COINIT_MULTITHREADED;
use windows::Win32::System::Com::CO_MTA_USAGE_COOKIE;
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
    ServiceNotCreated,

    ServiceNotConnected,

    /// Some unhandled COM error
    ComError(HRESULT),

    /// This should not happen, this means that successful COM call allocated a
    /// null pointer, in this case it is an error in the COM service, or it's
    /// usage.
    ComAllocatedNullPtr,

    // Sender error
    SenderError,
}

// impl From<HRESULT> for Error {
//     fn from(hr: HRESULT) -> Self {
//         if hr == HRESULT(0x800706BA) {
//             // Explorer.exe has mostlikely crashed
//             return Error::ServiceNotConnected;
//         }

//         Error::ComError(hr)
//     }
// }
// impl From<HRESULT> for Result<()> {
//     fn from(item: HRESULT) -> Self {
//         if !item.failed() {
//             Ok(())
//         } else {
//             Err(item.into())
//         }
//     }
// }

fn map_win_err(er: ::windows::core::Error) -> Error {
    Error::ComError(HRESULT(er.code().0 as u32))
}

struct ComSta();
impl ComSta {
    fn new() -> Self {
        #[cfg(debug_assertions)]
        log_output("CoInitializeEx COINIT_APARTMENTTHREADED");

        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap() };
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
    crossbeam_channel::Sender<ComFn>,
    std::thread::JoinHandle<()>,
)> = once_cell::sync::Lazy::new(|| {
    let (sender, receiver) = crossbeam_channel::unbounded::<ComFn>();
    (
        sender,
        std::thread::spawn(move || {
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

    WORKER_CHANNEL
        .0
        .send(Box::new(move |c| {
            for _ in 0..5 {
                let r = f(c);
                if let Err(Error::ServiceNotConnected) = r {
                    // Explorer.exe has mostlikely crashed, retry the function
                    c.drop_services();
                    continue;
                }

                sender.send(r).unwrap();
                return;
            }
        }))
        .unwrap();
    receiver.recv().unwrap()

    // Naive implementation that causes illegal memory access on rapid threading test
    //
    // std::thread::spawn(|| {
    //     let com = ComObjects::new();
    //     f(&com)
    // })
    // .join()
    // .unwrap()
}

pub trait ComObjectsAsResult {
    fn as_result(&self) -> Result<Rc<ComObjects>>;
}

impl ComObjectsAsResult for Weak<ComObjects> {
    fn as_result(&self) -> Result<Rc<ComObjects>> {
        self.upgrade().ok_or(Error::ServiceNotCreated)
    }
}

#[derive(Copy, Clone, Debug)]
pub enum DesktopInternal {
    Index(u32),
    Guid(GUID),
    IndexGuid(u32, GUID),
}
unsafe impl Send for DesktopInternal {}
unsafe impl Sync for DesktopInternal {}

// Impl equality for DesktopInternal
impl DesktopInternal {
    pub fn try_eq(&self, other: &Self) -> Result<bool> {
        match (self, other) {
            (DesktopInternal::Index(a), DesktopInternal::Index(b)) => Ok(a == b),
            (DesktopInternal::Guid(a), DesktopInternal::Guid(b)) => Ok(a == b),
            (DesktopInternal::IndexGuid(a, b), DesktopInternal::IndexGuid(c, d)) => {
                Ok(a == c && b == d)
            }
            (DesktopInternal::Index(a), DesktopInternal::IndexGuid(b, _)) => Ok(a == b),
            (DesktopInternal::IndexGuid(a, _), DesktopInternal::Index(b)) => Ok(a == b),
            (DesktopInternal::Guid(a), DesktopInternal::IndexGuid(_, b)) => Ok(a == b),
            (DesktopInternal::IndexGuid(_, a), DesktopInternal::Guid(b)) => Ok(a == b),
            _ => {
                let self_ = self.clone();
                let other_ = other.clone();
                with_com_objects(move |f| {
                    Ok(f.get_desktop_id(&self_)? == f.get_desktop_id(&other_)?)
                })
            }
        }
    }
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

impl<'a> TryFrom<ComIn<'a, IVirtualDesktop>> for DesktopInternal {
    type Error = Error;

    fn try_from(desktop: ComIn<IVirtualDesktop>) -> Result<Self> {
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
                CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_LOCAL_SERVER)
                    .map_err(map_win_err)?
            });
            *provider = Some(new_provider);
        }

        provider
            .as_ref()
            .map(|v| Rc::clone(&v))
            .ok_or(Error::ServiceNotCreated)
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
            .ok_or(Error::ServiceNotCreated)
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
            .ok_or(Error::ServiceNotCreated)
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
            .ok_or(Error::ServiceNotCreated)
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
            .ok_or(Error::ServiceNotCreated)
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
            .ok_or(Error::ServiceNotCreated)
    }

    pub fn drop_services(&self) {
        self.provider.borrow_mut().take();
        self.manager.borrow_mut().take();
        self.manager_internal.borrow_mut().take();
        self.notification_service.borrow_mut().take();
        self.pinned_apps.borrow_mut().take();
        self.view_collection.borrow_mut().take();
    }

    fn get_idesktops_array(&self) -> Result<IObjectArray> {
        let mut desktops = None;
        unsafe {
            self.get_manager_internal()?
                .get_desktops(0, &mut desktops)
                .as_result()?
        }
        Ok(desktops.unwrap())
    }

    fn get_desktop_index_by_guid(&self, id: &GUID) -> Result<u32> {
        let desktops = self.get_idesktops_array()?;
        let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
        for i in 0..count {
            let desktop_id: GUID =
                get_idesktop_guid(&unsafe { desktops.GetAt(i).map_err(map_win_err)? })?;
            if desktop_id == *id {
                return Ok(i);
            }
        }
        Err(Error::DesktopNotFound)
    }

    fn get_desktop_guid_by_index(&self, id: u32) -> Result<GUID> {
        let desktops = self.get_idesktops_array()?;
        let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
        if id >= count {
            return Err(Error::DesktopNotFound);
        }
        get_idesktop_guid(&unsafe { desktops.GetAt(id).map_err(map_win_err)? })
    }

    fn get_idesktop(&self, desktop: &DesktopInternal) -> Result<IVirtualDesktop> {
        match desktop {
            DesktopInternal::Index(id) => {
                let desktops = self.get_idesktops_array()?;
                let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
                if *id >= count {
                    return Err(Error::DesktopNotFound);
                }
                Ok(unsafe { desktops.GetAt(*id).map_err(map_win_err)? })
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
                    if e == Error::ComError(HRESULT(0x8002802B)) {
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
                    if er == Error::ComError(HRESULT(0x8002802B)) {
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
        let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
        let mut result = Vec::with_capacity(count as usize);
        for i in 0..count {
            let desktop = unsafe { desktops.GetAt(i).map_err(map_win_err)? };
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
        let desktop = desktop.unwrap();
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
        self.get_desktop_by_window(window)
            .map(|id| id.try_eq(&desktop))?
            .or(Ok(false))
    }

    pub fn is_window_on_current_desktop(&self, window: &HWND) -> Result<bool> {
        unsafe {
            let mut value = false;
            self.get_manager()?
                .is_window_on_current_desktop(window.clone(), &mut value)
                .as_result()
                .map_err(|er| match er {
                    // Window does not exist
                    Error::ComError(HRESULT(0x8002802B)) => Error::WindowNotFound,
                    _ => er,
                })?;
            Ok(value)
        }
    }

    pub fn move_window_to_desktop(&self, window: &HWND, desktop: &DesktopInternal) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        self.move_view_to_desktop(&view, desktop)
    }

    // pub fn get_desktop<T>(&self, desktop: T) -> Desktop
    // where
    //     T: Into<Desktop>,
    // {
    //     desktop.into()
    // }

    pub fn get_desktop_count(&self) -> Result<u32> {
        let desktops = self.get_idesktops_array()?;
        let count = unsafe { desktops.GetCount().map_err(map_win_err)? };
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
                    Error::ComError(HRESULT(0x8002802B)) => Error::WindowNotFound,
                    _ => er,
                })?
        };
        if desktop == GUID::default() {
            return Err(Error::DesktopNotFound);
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
        let desktop = desktop.unwrap();
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

#[cfg(test)]
mod test {
    use super::*;

    #[test]
    fn test_com_objects_non_thread_local() {
        let com_objects = super::ComObjects::new();
        let _provider = com_objects.get_provider().unwrap();
        let _manager = com_objects.get_manager().unwrap();
        let _manager_internal = com_objects.get_manager_internal().unwrap();
        let _notification_service = com_objects.get_notification_service().unwrap();
        let _pinned_apps = com_objects.get_pinned_apps().unwrap();
        let _view_collection = com_objects.get_view_collection().unwrap();
    }
}
