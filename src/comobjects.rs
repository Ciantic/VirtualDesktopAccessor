//! This module contains COM object for accessing the Windows Virtual Desktop API
#![allow(clippy::upper_case_acronyms)]

use super::interfaces_multi::*;
use super::Result;
use std::convert::TryFrom;
use std::rc::Rc;
use std::{cell::RefCell, ffi::c_void};
use windows::core::HRESULT;
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Com::CoIncrementMTAUsage;
use windows::Win32::System::Com::CLSCTX_LOCAL_SERVER;
use windows::{
    core::{Interface, GUID, HSTRING},
    Win32::{System::Com::CoCreateInstance, UI::Shell::Common::IObjectArray},
};

#[cfg(debug_assertions)]
use crate::log::log_output;

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

    /// Com is not initialized, call CoInitializeEx or CoIncrementMTAUsage
    ComNotInitialized,

    /// Com object not connected
    ComObjectNotConnected,

    /// Generic element not found
    ComElementNotFound,

    /// A requested interface was not supported. Windows might have changed the
    /// definitions of some unstable virtual desktop interfaces used by this
    /// library.
    ComNoInterface,

    /// Not implemented. When supporting multiple Windows versions this is
    /// returned for methods that don't exist in the current version.
    ComNotImplemented,

    /// Some unhandled COM error
    ComError(HRESULT),

    /// This should not happen, this means that successful COM call allocated a
    /// null pointer, in this case it is an error in the COM service, or it's
    /// usage.
    ComAllocatedNullPtr,

    /// Borrow error
    InternalBorrowError,
}

pub(crate) trait HRESULTHelpers {
    fn as_error(&self) -> Error;
    fn as_result(&self) -> Result<()>;
}

impl HRESULTHelpers for ::windows::core::HRESULT {
    fn as_error(&self) -> Error {
        match self.0 {
            -2147221164 => {
                // 0x80040154
                Error::ClassNotRegistered
            }
            -2147023174 => {
                // 0x800706BA
                Error::RpcServerNotAvailable
            }
            -2147220995 => {
                // 0x800401FD
                Error::ComObjectNotConnected
            }
            -2147319765 => {
                // 0x8002802B
                Error::ComElementNotFound
            }
            -2147221008 => {
                // 0x800401F0
                Error::ComNotInitialized
            }
            -2147467262 => {
                // 0x80004002
                Error::ComNoInterface
            }
            -2147467263 => {
                // 0x80004001
                Error::ComNotImplemented
            }
            _ => Error::ComError(*self),
        }
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
impl<'a> TryFrom<&'a ComIn<'a, IVirtualDesktop>> for DesktopInternal {
    type Error = Error;

    fn try_from(desktop: &'a ComIn<'a, IVirtualDesktop>) -> Result<Self> {
        let mut guid = GUID::default();
        unsafe { desktop.get_id(&mut guid).as_result()? }
        Ok(DesktopInternal::Guid(guid))
    }
}

pub struct ComObjects {
    provider: RefCell<Option<Rc<IServiceProvider>>>,
    manager: RefCell<Option<Rc<IVirtualDesktopManager>>>,
    manager_internal: RefCell<Option<Rc<IVirtualDesktopManagerInternal>>>,
    notification_service: RefCell<Option<Rc<IVirtualDesktopNotificationService>>>,
    pinned_apps: RefCell<Option<Rc<IVirtualDesktopPinnedApps>>>,
    view_collection: RefCell<Option<Rc<IApplicationViewCollection>>>,
}

fn retry_function<F, R>(com_objects: &ComObjects, f: F, _fn_name: &str) -> Result<R>
where
    F: Fn() -> Result<R>,
{
    let mut value = f();
    for _ in 0..3 {
        match &value {
            Err(er)
                if er == &Error::ClassNotRegistered
                    || er == &Error::RpcServerNotAvailable
                    || er == &Error::ComObjectNotConnected
                    || er == &Error::ComAllocatedNullPtr
                    || er == &Error::ComNotInitialized =>
            {
                log_format!("Retry the function \"{_fn_name}\" after {:?}", er);

                if er == &Error::ComNotInitialized {
                    let _ = unsafe { CoIncrementMTAUsage() };
                }

                drop(value);

                // If private function is decorated, then this drop_services call
                // will cause borrow issues.
                com_objects.drop_services();

                value = f();
            }
            _ => {
                break;
            }
        }
    }

    #[cfg(debug_assertions)]
    if let Err(er) = &value {
        log_format!(
            "Com_objects function \"{_fn_name}\" failed with {:?}",
            er
        );
    }

    value
}

/// Safely reruns the function if it returns one of the recoverable errors
///
/// This should be applied to only public functions in ComObjects struct, having
/// it in private functions is not necessary. Decorating private functions will
/// also cause borrowing issues.
macro_rules! retry_function {(
    $( #[$attr:meta] )*
    $pub:vis
    fn $fname:ident (
        &$self_:ident $(,)? $( $arg_name:ident : $ArgTy:ty ),* $(,)?
    ) -> $RetTy:ty
    $body:block
) => (
    $( #[$attr] )*
    #[allow(unused_parens)]
    $pub
    fn $fname (
        &$self_, $( $arg_name : $ArgTy ),*
    ) -> $RetTy
    {
        retry_function(&$self_, || -> $RetTy {
            $body
        }, stringify!($fname))
    }
)}

impl ComObjects {
    pub fn new() -> Self {
        Self {
            provider: RefCell::new(None),
            manager: RefCell::new(None),
            manager_internal: RefCell::new(None),
            notification_service: RefCell::new(None),
            pinned_apps: RefCell::new(None),
            view_collection: RefCell::new(None),
        }
    }

    fn get_provider(&self) -> Result<Rc<IServiceProvider>> {
        let mut provider = self
            .provider
            .try_borrow_mut()
            .map_err(|_| Error::InternalBorrowError)?;
        if provider.is_none() {
            let new_provider = Rc::new(unsafe {
                CoCreateInstance(&CLSID_ImmersiveShell, None, CLSCTX_LOCAL_SERVER)?
            });
            *provider = Some(new_provider);
        }

        provider
            .as_ref()
            .map(Rc::clone)
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_manager(&self) -> Result<Rc<IVirtualDesktopManager>> {
        let mut manager = self
            .manager
            .try_borrow_mut()
            .map_err(|_| Error::InternalBorrowError)?;
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
            .map(Rc::clone)
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_manager_internal(&self) -> Result<Rc<IVirtualDesktopManagerInternal>> {
        let mut manager_internal = self
            .manager_internal
            .try_borrow_mut()
            .map_err(|_| Error::InternalBorrowError)?;
        if manager_internal.is_none() {
            let provider = self.get_provider()?;
            *manager_internal = Some(Rc::new(unsafe {
                IVirtualDesktopManagerInternal::query_service(&provider)?
            }));
        }
        manager_internal
            .as_ref()
            .map(Rc::clone)
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_notification_service(&self) -> Result<Rc<IVirtualDesktopNotificationService>> {
        let mut notification_service = self
            .notification_service
            .try_borrow_mut()
            .map_err(|_| Error::InternalBorrowError)?;
        if notification_service.is_none() {
            let provider = self.get_provider()?;
            *notification_service = Some(Rc::new(unsafe {
                IVirtualDesktopNotificationService::query_service(&provider)?
            }));
        }
        notification_service
            .as_ref()
            .map(Rc::clone)
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_pinned_apps(&self) -> Result<Rc<IVirtualDesktopPinnedApps>> {
        let mut pinned_apps = self
            .pinned_apps
            .try_borrow_mut()
            .map_err(|_| Error::InternalBorrowError)?;
        if pinned_apps.is_none() {
            let provider = self.get_provider()?;
            *pinned_apps = Some(Rc::new(unsafe { IVirtualDesktopPinnedApps::query_service(&provider)? }));
        }
        pinned_apps
            .as_ref()
            .ok_or(Error::ComAllocatedNullPtr)
            .map(Rc::clone)
    }

    fn get_view_collection(&self) -> Result<Rc<IApplicationViewCollection>> {
        let mut view_collection = self
            .view_collection
            .try_borrow_mut()
            .map_err(|_| Error::InternalBorrowError)?;
        if view_collection.is_none() {
            let provider = self.get_provider()?;
            *view_collection = Some(Rc::new(unsafe {
                IApplicationViewCollection::query_service(&provider)?
            }));
        }
        view_collection
            .as_ref()
            .map(Rc::clone)
            .ok_or(Error::ComAllocatedNullPtr)
    }

    fn drop_services(&self) {
        // Current implementation would be safe drop like this, but in case I
        // ever refactor I don't use this:

        // drop(self.provider.take());
        // drop(self.manager.take());
        // drop(self.manager_internal.take());

        // Instead I use try_borrow_mut() and map() to drop services.
        let _ = self.provider.try_borrow_mut().map(|mut v| v.take());
        let _ = self.manager.try_borrow_mut().map(|mut v| v.take());
        let _ = self.manager_internal.try_borrow_mut().map(|mut v| v.take());
        let _ = self
            .notification_service
            .try_borrow_mut()
            .map(|mut v| v.take());
        let _ = self.pinned_apps.try_borrow_mut().map(|mut v| v.take());
        let _ = self.view_collection.try_borrow_mut().map(|mut v| v.take());
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
                        .get_desktop_count(&mut out_count)
                        .as_result()
                };

                #[cfg(debug_assertions)]
                if let Err(er) = &res {
                    log_output(&format!("is connected error: {:?} {}", er, out_count));
                }

                if out_count == 0 || res.is_err() {
                    return false;
                }
                return true;
            }
            Err(_) => false,
        }
    }

    fn get_idesktops_array(&self) -> Result<IObjectArray> {
        let mut desktops = None;
        unsafe {
            self.get_manager_internal()?
                .get_desktops(&mut desktops)
                .as_result()?
        }
        desktops.ok_or(Error::ComAllocatedNullPtr)
    }

    fn get_desktop_index_by_guid(&self, id: &GUID) -> Result<u32> {
        let desktops = self.get_idesktops_array()?;
        let count = unsafe { desktops.GetCount()? };
        for i in 0..count {
            let desktop_id: GUID = get_idesktop_guid(&unsafe { IObjectArrayGetAt(&desktops, i)? })?;
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
        get_idesktop_guid(&unsafe { IObjectArrayGetAt(&desktops, id)? })
    }

    fn get_idesktop(&self, desktop: &DesktopInternal) -> Result<IVirtualDesktop> {
        match desktop {
            DesktopInternal::Index(id) => {
                let desktops = self.get_idesktops_array()?;
                let count = unsafe { desktops.GetCount()? };
                if *id >= count {
                    return Err(Error::DesktopNotFound);
                }
                Ok(unsafe { IObjectArrayGetAt(&desktops, *id)? })
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
        view: ComIn<IApplicationView>,
        desktop: &DesktopInternal,
    ) -> Result<()> {
        let desktop = self.get_idesktop(desktop)?;
        unsafe {
            self.get_manager_internal()?
                .move_view_to_desktop(view, ComIn::new(&desktop))
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
                .get_view_for_hwnd(*hwnd, &mut view)
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

    #[apply(retry_function)]
    pub fn get_desktop_index(&self, id: &DesktopInternal) -> Result<u32> {
        match id {
            DesktopInternal::Index(id) => Ok(*id),
            DesktopInternal::Guid(guid) => self.get_desktop_index_by_guid(guid),
            DesktopInternal::IndexGuid(id, _) => Ok(*id),
        }
    }

    #[apply(retry_function)]
    pub fn get_desktop_id(&self, desktop: &DesktopInternal) -> Result<GUID> {
        match desktop {
            DesktopInternal::Index(id) => self.get_desktop_guid_by_index(*id),
            DesktopInternal::Guid(guid) => Ok(*guid),
            DesktopInternal::IndexGuid(_, guid) => Ok(*guid),
        }
    }

    #[apply(retry_function)]
    pub fn get_desktops(&self) -> Result<Vec<DesktopInternal>> {
        let desktops = self.get_idesktops_array()?;
        let count = unsafe { desktops.GetCount()? };
        let mut result = Vec::with_capacity(count as usize);
        for i in 0..count {
            let desktop = unsafe { IObjectArrayGetAt(&desktops, i)? };
            let id = get_idesktop_guid(&desktop)?;
            result.push(DesktopInternal::IndexGuid(i, id));
        }
        Ok(result)
    }

    #[apply(retry_function)]
    pub fn register_for_notifications(
        &self,
        // notification: &IVirtualDesktopNotification,
        notification: *mut c_void, // IVirtualDesktopNotification raw pointer
    ) -> Result<u32> {
        let notification_service = self.get_notification_service()?;

        unsafe {
            let mut cookie = 0;
            notification_service
                .register(notification, &mut cookie)
                .as_result()
                .map(|_| cookie)
        }
    }

    #[apply(retry_function)]
    pub fn unregister_for_notifications(&self, cookie: u32) -> Result<()> {
        let notification_service = self.get_notification_service()?;
        unsafe { notification_service.unregister(cookie).as_result() }
    }

    #[apply(retry_function)]
    pub fn switch_desktop(&self, desktop: &DesktopInternal) -> Result<()> {
        let desktop = self.get_idesktop(desktop)?;
        unsafe {
            self.get_manager_internal()?
                .switch_desktop(ComIn::new(&desktop))
                .as_result()?
        }
        Ok(())
    }

    #[apply(retry_function)]
    pub fn create_desktop(&self) -> Result<DesktopInternal> {
        let mut desktop = None;
        unsafe {
            self.get_manager_internal()?
                .create_desktop(&mut desktop)
                .as_result()?
        }
        let desktop = desktop.ok_or(Error::ComAllocatedNullPtr)?;
        let id = get_idesktop_guid(&desktop)?;
        let index = self.get_desktop_index_by_guid(&id)?;
        Ok(DesktopInternal::IndexGuid(index, id))
    }

    #[apply(retry_function)]
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

    #[apply(retry_function)]
    pub fn is_window_on_desktop(&self, window: &HWND, desktop: &DesktopInternal) -> Result<bool> {
        let desktop_win = self.get_desktop_by_window(window)?;
        Ok(self.get_desktop_id(&desktop_win)? == self.get_desktop_id(desktop)?)
    }

    #[apply(retry_function)]
    pub fn is_window_on_current_desktop(&self, window: &HWND) -> Result<bool> {
        unsafe {
            let mut value = false;
            self.get_manager()?
                .is_window_on_current_desktop(*window, &mut value)
                .as_result()
                .map_err(|er| match er {
                    // Window does not exist
                    Error::ComElementNotFound => Error::WindowNotFound,
                    _ => er,
                })?;
            Ok(value)
        }
    }

    #[apply(retry_function)]
    pub fn move_window_to_desktop(&self, window: &HWND, desktop: &DesktopInternal) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        self.move_view_to_desktop(ComIn::new(&view), desktop)
    }

    #[apply(retry_function)]
    pub fn get_desktop_count(&self) -> Result<u32> {
        let manager = self.get_manager_internal()?;
        let mut count = 0;
        unsafe {
            manager.get_desktop_count(&mut count).as_result()?;
        };
        Ok(count)
    }

    #[apply(retry_function)]
    pub fn get_desktop_by_window(&self, window: &HWND) -> Result<DesktopInternal> {
        let mut desktop = GUID::default();
        unsafe {
            self.get_manager()?
                .get_desktop_by_window(*window, &mut desktop)
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

    #[apply(retry_function)]
    pub fn get_current_desktop(&self) -> Result<DesktopInternal> {
        let mut desktop = None;
        unsafe {
            self.get_manager_internal()?
                .get_current_desktop(&mut desktop)
                .as_result()?
        }
        let desktop = desktop.ok_or(Error::ComAllocatedNullPtr)?;
        let id = get_idesktop_guid(&desktop)?;
        Ok(DesktopInternal::Guid(id))
    }

    #[apply(retry_function)]
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

    #[apply(retry_function)]
    pub fn pin_window(&self, window: &HWND) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        unsafe {
            self.get_pinned_apps()?
                .pin_view(ComIn::new(&view))
                .as_result()?;
        }
        Ok(())
    }

    #[apply(retry_function)]
    pub fn unpin_window(&self, window: &HWND) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        unsafe {
            self.get_pinned_apps()?
                .unpin_view(ComIn::new(&view))
                .as_result()?;
        }
        Ok(())
    }

    #[apply(retry_function)]
    fn get_iapplication_id_for_view(&self, view: &IApplicationView) -> Result<APPIDPWSTR> {
        let mut app_id: APPIDPWSTR = std::ptr::null_mut();
        unsafe {
            view.get_app_user_model_id(&mut app_id as *mut _ as *mut _)
                .as_result()?
        }
        Ok(app_id)
    }

    #[apply(retry_function)]
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

    #[apply(retry_function)]
    pub fn pin_app(&self, window: &HWND) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        let app_id = self.get_iapplication_id_for_view(&view)?;
        unsafe {
            self.get_pinned_apps()?.pin_app(app_id).as_result()?;
        }
        Ok(())
    }

    #[apply(retry_function)]
    pub fn unpin_app(&self, window: &HWND) -> Result<()> {
        let view = self.get_iapplication_view_for_hwnd(window)?;
        let app_id = self.get_iapplication_id_for_view(&view)?;
        unsafe {
            self.get_pinned_apps()?.unpin_app(app_id).as_result()?;
        }
        Ok(())
    }

    #[apply(retry_function)]
    pub fn get_desktop_name(&self, desktop: &DesktopInternal) -> Result<String> {
        let desktop = self.get_idesktop(desktop)?;
        let mut name = HSTRING::default();
        unsafe {
            desktop.get_name(&mut name).as_result()?;
        }
        Ok(name.to_string())
    }

    #[apply(retry_function)]
    pub fn set_desktop_name(&self, desktop: &DesktopInternal, name: &str) -> Result<()> {
        let desktop = self.get_idesktop(desktop)?;
        let manager_internal = self.get_manager_internal()?;

        unsafe {
            manager_internal
                .set_name(ComIn::new(&desktop), HSTRING::from(name))
                .as_result()
        }
    }

    #[apply(retry_function)]
    pub fn get_desktop_wallpaper(&self, desktop: &DesktopInternal) -> Result<String> {
        let desktop = self.get_idesktop(desktop)?;
        let mut path = HSTRING::default();
        unsafe {
            desktop.get_wallpaper(&mut path).as_result()?;
        }
        Ok(path.to_string())
    }

    #[apply(retry_function)]
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

fn get_idesktop_guid(desktop: &IVirtualDesktop) -> Result<GUID> {
    let mut guid = GUID::default();
    unsafe { desktop.get_id(&mut guid).as_result()? }
    Ok(guid)
}

thread_local! {
    static COM_OBJECTS: ComObjects = ComObjects::new();
}

/// This is a helper function to initialize and run COM related functions in a
/// a single thread.
///
/// Virtual Desktop COM Objects don't like to being called from different
/// threads rapidly, something goes wrong. This function ensures that all COM
/// calls are done in a single thread.
pub fn with_com_objects<F, T>(f: F) -> Result<T>
where
    F: Fn(&ComObjects) -> Result<T> + 'static,
    T: 'static,
{
    // return std::thread::scope(|env| {
    //     let com2 = ComObjects::new();
    //     run_function_and_retry(&f, &com2)
    // });

    // return COM_OBJECTS.with(|c| run_function_and_retry(&f, &c));
    COM_OBJECTS.with(|c| f(c))
}
