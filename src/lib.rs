// The debug version
#[cfg(feature = "debug")]
macro_rules! debug_print {
    ($( $args:expr ),*) => { println!( $( $args ),* ); }
}

// Non-debug version
#[cfg(not(feature = "debug"))]
macro_rules! debug_print {
    ($( $args:expr ),*) => {};
}

mod changelistener;
mod comhelpers;
mod guid;
mod immersive;
mod interfaces;
// mod utils;
use com::runtime::{init_apartment, ApartmentType};
use com::{
    sys::{FAILED, HRESULT},
    ComInterface, ComRc,
};
pub use guid::DesktopID;
use winapi::shared::windef::HWND;

use changelistener::VirtualDesktopChangeListener;
use comhelpers::create_instance;
use immersive::{get_immersive_service, get_immersive_service_for_class};
use interfaces::{
    CLSID_IVirtualNotificationService, CLSID_ImmersiveShell, CLSID_VirtualDesktopManagerInternal,
    CLSID_VirtualDesktopPinnedApps, IApplicationView, IApplicationViewCollection,
    IApplicationViewVTable, IObjectArray, IObjectArrayVTable, IServiceProvider, IVirtualDesktop,
    IVirtualDesktopManager, IVirtualDesktopManagerInternal, IVirtualDesktopPinnedApps,
    IVirtualDesktopVTable,
};
use std::cell::Cell;

#[derive(Debug, Clone)]
pub enum Error {
    InitializationError(HRESULT),
    UnknownError,
    WindowNotFound,
    DesktopNotFound,
    ApartmentInitError(HRESULT),
    ComResultError(HRESULT, String),
}

/// Provides the stateful helper to accessing the Windows 10 Virtual Desktop
/// functions.
///
/// If you don't use other COM objects in your project, you have to use
/// `VirtualDesktopService::create_with_com()` constructor.
///
pub struct VirtualDesktopService {
    on_drop_deinit_apartment: Cell<bool>,
    #[allow(dead_code)]
    service_provider: ComRc<dyn IServiceProvider>,
    virtual_desktop_manager: ComRc<dyn IVirtualDesktopManager>,
    virtual_desktop_manager_internal: ComRc<dyn IVirtualDesktopManagerInternal>,
    app_view_collection: ComRc<dyn IApplicationViewCollection>,
    pinned_apps: ComRc<dyn IVirtualDesktopPinnedApps>,
    // virtual_desktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService>,
    events: Box<VirtualDesktopChangeListener>,
}

impl VirtualDesktopService {
    /// Initialize service and COM apartment. If you don't use other COM API's,
    /// you have to use this initialization.
    pub fn create_with_com() -> Result<VirtualDesktopService, Error> {
        // init_runtime().map_err(|o| Error::ApartmentInitError(o))?;
        init_apartment(ApartmentType::Multithreaded).map_err(|op| Error::ApartmentInitError(op))?;
        let service = VirtualDesktopService::create()?;
        service.on_drop_deinit_apartment.set(true);
        Ok(service)
    }

    /// Initialize only the service, must be-created on TaskbarCreated message
    pub fn create() -> Result<VirtualDesktopService, Error> {
        let service_provider = create_instance::<dyn IServiceProvider>(&CLSID_ImmersiveShell)
            .map_err(|hr| Error::InitializationError(hr))?;

        let virtual_desktop_manager = get_immersive_service::<dyn IVirtualDesktopManager>(
            &service_provider,
        )
        .map_err(|err| {
            Error::ComResultError(
                err,
                "IServiceProvider.query_service IVirtualDesktopManager".into(),
            )
        })?;

        let virtualdesktop_notification_service =
            get_immersive_service_for_class(&service_provider, CLSID_IVirtualNotificationService)
                .map_err(|err| {
                Error::ComResultError(
                    err,
                    "IServiceProvider.query_service IVirtualDesktopNotificationService".into(),
                )
            })?;

        let vd_manager_internal =
            get_immersive_service_for_class(&service_provider, CLSID_VirtualDesktopManagerInternal)
                .map_err(|err| {
                    Error::ComResultError(
                        err,
                        "IServiceProvider.query_service IVirtualDesktopManagerInternal".into(),
                    )
                })?;

        let app_view_collection = get_immersive_service(&service_provider).map_err(|err| {
            Error::ComResultError(
                err,
                "IServiceProvider.query_service IApplicationViewCollection".into(),
            )
        })?;

        let pinned_apps =
            get_immersive_service_for_class(&service_provider, CLSID_VirtualDesktopPinnedApps)
                .map_err(|err| {
                    Error::ComResultError(
                        err,
                        "IServiceProvider.query_service IVirtualDesktopPinnedApps".into(),
                    )
                })?;

        let listener = VirtualDesktopChangeListener::register(virtualdesktop_notification_service)
            .map_err(|err| {
                Error::ComResultError(
                    err,
                    "IServiceProvider.query_service VirtualDesktopChangeListener".into(),
                )
            })?;

        Ok(VirtualDesktopService {
            on_drop_deinit_apartment: Cell::new(false),
            virtual_desktop_manager: virtual_desktop_manager,
            service_provider: service_provider,
            events: listener,
            virtual_desktop_manager_internal: vd_manager_internal,
            app_view_collection: app_view_collection,
            pinned_apps: pinned_apps,
        })
    }

    /// Get raw desktop list
    fn _get_desktops(&self) -> Result<Vec<ComRc<dyn IVirtualDesktop>>, Error> {
        let ptr: *mut IObjectArrayVTable = std::ptr::null_mut();
        let res = unsafe { self.virtual_desktop_manager_internal.get_desktops(&ptr) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManagerInternal.get_desktops".into(),
            ));
        }

        let dc: ComRc<dyn IObjectArray> = unsafe { ComRc::from_raw(ptr as *mut *mut _) };
        let mut count = 0;
        let res = unsafe { dc.get_count(&mut count) };
        if FAILED(res) {
            return Err(Error::ComResultError(res, "IObjectArray.get_count".into()));
        }

        let mut desktops: Vec<ComRc<dyn IVirtualDesktop>> = vec![];

        for i in 0..(count - 1) {
            let ptr = std::ptr::null_mut();
            let res = unsafe { dc.get_at(i, &IVirtualDesktop::IID, &ptr) };
            if FAILED(res) {
                return Err(Error::ComResultError(res, "IObjectArray.get_at".into()));
            }
            // TODO: How long does the ptr is guarenteed to be alive? https://github.com/microsoft/com-rs/issues/141
            let desktop = unsafe { ComRc::from_raw(ptr as *mut _) };

            desktops.push(desktop);
        }
        Ok(desktops)
    }

    /// Get raw desktop by ID
    fn _get_desktop_by_id(&self, desktop: &DesktopID) -> Result<ComRc<dyn IVirtualDesktop>, Error> {
        // TODO: Is this safe? https://github.com/microsoft/com-rs/issues/141
        self._get_desktops()?
            .iter()
            .find(|v| {
                let mut id: DesktopID = Default::default();
                unsafe {
                    v.get_id(&mut id);
                }
                &id == desktop
            })
            .map(|v| v.clone())
            .ok_or(Error::DesktopNotFound)
    }

    /// Get application view for raw window
    fn _get_application_view_for_hwnd(
        &self,
        hwnd: HWND,
    ) -> Result<ComRc<dyn IApplicationView>, Error> {
        let ptr: *mut IApplicationViewVTable = std::ptr::null_mut();
        let res = unsafe { self.app_view_collection.get_view_for_hwnd(hwnd, &ptr) };
        if ptr.is_null() {
            return Err(Error::WindowNotFound);
        }
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IApplicationView.get_view_for_hwnd".into(),
            ));
        }
        return Ok(unsafe { ComRc::from_raw(ptr as *mut _) });
    }

    /// Get desktops (GUID's)
    pub fn get_desktops(&self) -> Result<Vec<DesktopID>, Error> {
        let ptr: *mut IObjectArrayVTable = std::ptr::null_mut();
        let res = unsafe { self.virtual_desktop_manager_internal.get_desktops(&ptr) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManagerInternal.get_desktops".into(),
            ));
        }

        let dc: ComRc<dyn IObjectArray> = unsafe { ComRc::from_raw(ptr as *mut _) };
        let mut count = 0;
        let res = unsafe { dc.get_count(&mut count) };
        if FAILED(res) {
            return Err(Error::ComResultError(res, "IObjectArray.get_count".into()));
        }

        let mut desktops: Vec<DesktopID> = vec![];

        for i in 0..(count - 1) {
            let ptr = std::ptr::null_mut();
            let res = unsafe { dc.get_at(i, &IVirtualDesktop::IID, &ptr) };
            if FAILED(res) {
                return Err(Error::ComResultError(res, "IObjectArray.get_at".into()));
            }
            let desktop: ComRc<dyn IVirtualDesktop> = unsafe { ComRc::from_raw(ptr as *mut _) };

            let mut desktopid = Default::default();
            let res = unsafe { desktop.get_id(&mut desktopid) };

            if FAILED(res) {
                return Err(Error::ComResultError(res, "IVirtualDesktop.get_id".into()));
            }
            desktops.push(desktopid);
        }

        Ok(desktops)
    }

    /// Get number of desktops
    pub fn get_desktop_count(&self) -> Result<u32, Error> {
        let ptr: *mut IObjectArrayVTable = std::ptr::null_mut();
        let res = unsafe { self.virtual_desktop_manager_internal.get_desktops(&ptr) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManagerInternal.get_desktops".into(),
            ));
        }

        let dc: ComRc<dyn IObjectArray> = unsafe { ComRc::from_raw(ptr as *mut _) };
        let mut count = 0;
        let res = unsafe { dc.get_count(&mut count) };
        if FAILED(res) {
            return Err(Error::ComResultError(res, "IObjectArray.get_count".into()));
        }
        Ok(count)
    }

    /// Get current desktop GUID
    pub fn get_current_desktop(&self) -> Result<DesktopID, Error> {
        let ptr: *mut IVirtualDesktopVTable = std::ptr::null_mut();

        let res = unsafe {
            self.virtual_desktop_manager_internal
                .get_current_desktop(&ptr)
        };
        if FAILED(res) {
            debug_print!("get_current_desktop failed {:?}", res);
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManagerInternal.get_current_desktop".into(),
            ));
        }

        let resultdc: ComRc<dyn IVirtualDesktop> = unsafe { ComRc::from_raw(ptr as *mut _) };
        let mut desktopid = Default::default();
        unsafe { resultdc.get_id(&mut desktopid) };

        Ok(desktopid)
    }

    /// Get window desktop ID
    pub fn get_desktop_by_window(&self, hwnd: HWND) -> Result<DesktopID, Error> {
        let mut desktop = Default::default();
        let res = unsafe {
            self.virtual_desktop_manager
                .get_desktop_by_window(hwnd, &mut desktop)
        };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManager.get_desktop_by_window".into(),
            ));
        }
        Ok(desktop)
    }

    /// Is window on current virtual desktop
    pub fn is_window_on_current_virtual_desktop(&self, hwnd: HWND) -> Result<bool, Error> {
        let mut isit = false;
        let res = unsafe {
            self.virtual_desktop_manager
                .is_window_on_current_desktop(hwnd, &mut isit)
        };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManager.is_window_on_current_desktop".into(),
            ));
        }
        Ok(isit)
    }

    /// Is window on desktop
    pub fn is_window_on_desktop(&self, hwnd: HWND, desktop: &DesktopID) -> Result<bool, Error> {
        let window_desktop = self.get_desktop_by_window(hwnd)?;
        Ok(&window_desktop == desktop)
    }

    /// Move window to desktop
    pub fn move_window_to_desktop(&self, hwnd: HWND, desktop: &DesktopID) -> Result<(), Error> {
        let desktop = self._get_desktop_by_id(desktop)?;
        let ptr = desktop
            .get_interface::<dyn IVirtualDesktop>()
            .ok_or(Error::DesktopNotFound)?;
        let view = self._get_application_view_for_hwnd(hwnd)?;
        let res = unsafe {
            self.virtual_desktop_manager_internal
                .move_view_to_desktop(view, ptr)
        };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManager.move_view_to_desktop".into(),
            ));
        }
        Ok(())
    }

    /// Go to desktop
    pub fn go_to_desktop(&self, desktop: &DesktopID) -> Result<(), Error> {
        let d = self._get_desktop_by_id(desktop)?;
        let res = unsafe { self.virtual_desktop_manager_internal.switch_desktop(d) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManagerInternal.switch_desktop".into(),
            ));
        }
        Ok(())
    }

    /// Is window pinned?
    pub fn is_pinned_window(&self, hwnd: HWND) -> Result<bool, Error> {
        let view = self._get_application_view_for_hwnd(hwnd)?;
        let mut test: bool = false;
        let res = unsafe { self.pinned_apps.is_view_pinned(view, &mut test) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopPinnedApps.is_view_pinned".into(),
            ));
        }
        Ok(test)
    }

    /// Pin window
    pub fn pin_window(&self, hwnd: HWND) -> Result<(), Error> {
        let view = self._get_application_view_for_hwnd(hwnd)?;
        let res = unsafe { self.pinned_apps.pin_view(view) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopPinnedApps.pin_view".into(),
            ));
        }
        Ok(())
    }

    /// Unpin window
    pub fn unpin_window(&self, hwnd: HWND) -> Result<(), Error> {
        let view = self._get_application_view_for_hwnd(hwnd)?;
        let res = unsafe { self.pinned_apps.unpin_view(view) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopPinnedApps.unpin_view".into(),
            ));
        }
        Ok(())
    }

    /// Callback for desktop change event, callback gets old desktop id, and new desktop id
    pub fn on_desktop_change(&self, callback: Box<dyn Fn(DesktopID, DesktopID) -> ()>) {
        self.events.on_desktop_change(callback);
    }

    /// Callback for desktop creation, callback recieves new desktop id
    pub fn on_desktop_created(&self, callback: Box<dyn Fn(DesktopID) -> ()>) {
        self.events.on_desktop_created(callback);
    }

    /// Callback for on desktop destroy event, callback recieves old desktop id
    pub fn on_desktop_destroyed(&self, callback: Box<dyn Fn(DesktopID) -> ()>) {
        self.events.on_desktop_destroyed(callback);
    }

    /// Callback for window changes, e.g. if window changes to different
    /// desktop, or window gets destroyed. Callback recieves HWND of thumbnail
    /// window (most likely top level window HWND).
    ///
    /// *Note* This can be a very chatty callback, and may have some false
    /// positives.
    pub fn on_window_change(&self, callback: Box<dyn Fn(HWND) -> ()>) {
        self.events.on_window_change(callback);
    }

    /*
    /// Is pinned app
    pub fn is_pinned_app(&self, hwnd: HWND) -> Result<(), Error> {
        Err(Error::UnknownError)
    }

    /// Pin app
    pub fn pin_app(&self, hwnd: HWND) -> Result<(), Error> {
        Err(Error::UnknownError)
    }

    /// Unpin app
    pub fn unpin_app(&self, hwnd: HWND) -> Result<(), Error> {
        Err(Error::UnknownError)
    }
    */

    /*
    /// Get desktop by desktop number
    pub fn get_desktop_by_number(&self) -> Result<GUID, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }

    /// Get current desktop number
    pub fn get_current_desktop_number() -> Result<u32, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }

    ///
    pub fn get_desktop_number_by_id(guid: &GUID) -> Result<i32, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn get_window_desktop_number(hwnd: i32) -> Result<i32, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn is_window_on_desktop_number(
        hwnd: i32,
        desktop_number: u32,
    ) -> Result<bool, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn move_window_to_desktop_number(
        hwnd: i32,
        desktop_number: u32,
    ) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn go_to_desktop_number(hwnd: i32, desktop_number: u32) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    */
}

impl Drop for VirtualDesktopService {
    fn drop(&mut self) {
        if self.on_drop_deinit_apartment.get() {
            // deinit_apartment() // TODO: This panics for me in tests, why?
        }
    }
}
