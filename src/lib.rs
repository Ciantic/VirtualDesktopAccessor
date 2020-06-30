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
use com::runtime::{
    deinit_apartment, get_class_object, init_apartment, init_runtime, ApartmentType,
};
use com::{
    co_class,
    interfaces::IUnknown,
    sys::{CoCreateInstance, CLSCTX_INPROC_SERVER, FAILED, HRESULT, S_OK},
    ComInterface, ComPtr, ComRc, IID,
};
pub use guid::DesktopID;
use std::iter::FilterMap;
use winapi::shared::windef::HWND;

use changelistener::VirtualDesktopChangeListener;
use comhelpers::create_instance;
use immersive::{get_immersive_service, get_immersive_service_for_class};
use interfaces::{
    CLSID_IVirtualNotificationService, CLSID_ImmersiveShell, CLSID_VirtualDesktopManagerInternal,
    CLSID_VirtualDesktopPinnedApps, IApplicationView, IApplicationViewCollection,
    IApplicationViewCollectionVTable, IApplicationViewVTable, IID_IVirtualDesktopNotification,
    IObjectArray, IObjectArrayVTable, IServiceProvider, IVirtualDesktop, IVirtualDesktopManager,
    IVirtualDesktopManagerInternal, IVirtualDesktopNotification,
    IVirtualDesktopNotificationService, IVirtualDesktopPinnedApps, IVirtualDesktopVTable,
};
use std::{cell::Cell, ffi::c_void, ptr::null_mut};

#[derive(Debug, Clone)]
pub enum Error {
    UnknownError,
    WindowNotFound,
    DesktopNotFound,
    ApartmentInitError(HRESULT),
    ComResultError(HRESULT, String),
}

pub struct VirtualDesktopService {
    on_drop_deinit_apartment: Cell<bool>,
    service_provider: ComRc<dyn IServiceProvider>,
    virtual_desktop_manager: ComRc<dyn IVirtualDesktopManager>,
    virtual_desktop_manager_internal: ComRc<dyn IVirtualDesktopManagerInternal>,
    app_view_collection: ComRc<dyn IApplicationViewCollection>,
    pinned_apps: ComRc<dyn IVirtualDesktopPinnedApps>,
    // virtual_desktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService>,
    pub events: Box<VirtualDesktopChangeListener>,
}

// TODO: Remove all unwraps!

impl VirtualDesktopService {
    /// Get raw desktop list
    fn _get_desktops(&self) -> Result<Vec<ComPtr<dyn IVirtualDesktop>>, Error> {
        let ptr: *mut IObjectArrayVTable = std::ptr::null_mut();
        let res = unsafe { self.virtual_desktop_manager_internal.get_desktops(&ptr) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopManagerInternal.get_desktops".into(),
            ));
        }

        let dc: ComRc<dyn IObjectArray> = ComRc::new(unsafe { ComPtr::new(ptr as *mut _) });
        let mut count = 0;
        let res = unsafe { dc.get_count(&mut count) };
        if FAILED(res) {
            return Err(Error::ComResultError(res, "IObjectArray.get_count".into()));
        }

        let mut desktops: Vec<ComPtr<dyn IVirtualDesktop>> = vec![];

        for i in 0..(count - 1) {
            let ptr = std::ptr::null_mut();
            let res = unsafe { dc.get_at(i, &IVirtualDesktop::IID, &ptr) };
            if FAILED(res) {
                return Err(Error::ComResultError(res, "IObjectArray.get_at".into()));
            }
            // TODO: How long does the ptr is guarenteed to be alive? https://github.com/microsoft/com-rs/issues/141
            let desktop: ComPtr<dyn IVirtualDesktop> = unsafe { ComPtr::new(ptr as *mut _) };

            desktops.push(desktop);
        }
        Ok(desktops)
    }

    /// Get raw desktop by ID
    fn _get_desktop_by_id(
        &self,
        desktop: &DesktopID,
    ) -> Result<ComPtr<dyn IVirtualDesktop>, Error> {
        // TODO: Is this safe? https://github.com/microsoft/com-rs/issues/141
        self._get_desktops()
            .unwrap()
            .iter()
            .find(|v| {
                let mut id: DesktopID = Default::default();
                unsafe { ComRc::new(v.clone().clone()).get_id(&mut id) };
                &id == desktop
            })
            .map(|v| v.clone())
            .ok_or(Error::DesktopNotFound)
    }

    /// Get application view for raw window
    fn _get_application_view_for_hwnd(
        &self,
        hwnd: HWND,
    ) -> Result<ComPtr<dyn IApplicationView>, Error> {
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
        return Ok(unsafe { ComPtr::new(ptr as *mut _) });
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

        let dc: ComRc<dyn IObjectArray> = ComRc::new(unsafe { ComPtr::new(ptr as *mut _) });
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
            let desktop: ComRc<dyn IVirtualDesktop> =
                ComRc::new(unsafe { ComPtr::new(ptr as *mut _) });

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

        let dc: ComRc<dyn IObjectArray> = ComRc::new(unsafe { ComPtr::new(ptr as *mut _) });
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

        let resultptr: ComPtr<dyn IVirtualDesktop> = unsafe { ComPtr::new(ptr as *mut _) };
        let resultdc = ComRc::new(resultptr);
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
        let view = self._get_application_view_for_hwnd(hwnd)?;
        let res = unsafe {
            self.virtual_desktop_manager_internal
                .move_view_to_desktop(view, desktop)
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
        let mut isIt: bool = false;
        let res = unsafe { self.pinned_apps.is_view_pinned(view, &mut isIt) };
        if FAILED(res) {
            return Err(Error::ComResultError(
                res,
                "IVirtualDesktopPinnedApps.is_view_pinned".into(),
            ));
        }
        Ok(isIt)
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

    /// Initialize service and COM apartment. If you don't use other COM API's,
    /// you may use this initialization.
    pub fn initialize() -> Result<VirtualDesktopService, Error> {
        // init_runtime().map_err(|o| Error::ApartmentInitError(o))?;
        init_apartment(ApartmentType::Multithreaded).map_err(|op| Error::ApartmentInitError(op))?;
        let service = VirtualDesktopService::initialize_only_service()?;
        service.on_drop_deinit_apartment.set(true);
        Ok(service)
    }

    /// Initialize only ImmersiveShell provider service, must be re-called on
    /// TaskbarCreated message
    pub fn initialize_only_service() -> Result<VirtualDesktopService, Error> {
        let service_provider =
            create_instance::<dyn IServiceProvider>(&CLSID_ImmersiveShell).unwrap();
        let virtual_desktop_manager =
            get_immersive_service::<dyn IVirtualDesktopManager>(&service_provider).unwrap();
        let virtualdesktop_notification_service =
            get_immersive_service_for_class::<dyn IVirtualDesktopNotificationService>(
                &service_provider,
                CLSID_IVirtualNotificationService,
            )
            .unwrap();
        let vd_manager_internal =
            get_immersive_service_for_class(&service_provider, CLSID_VirtualDesktopManagerInternal)
                .unwrap();
        let app_view_collection =
            get_immersive_service::<dyn IApplicationViewCollection>(&service_provider).unwrap();

        let pinned_apps = get_immersive_service_for_class::<dyn IVirtualDesktopPinnedApps>(
            &service_provider,
            CLSID_VirtualDesktopPinnedApps,
        )
        .unwrap();

        let listener =
            VirtualDesktopChangeListener::register(virtualdesktop_notification_service).unwrap();
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
}

impl Drop for VirtualDesktopService {
    fn drop(&mut self) {
        if self.on_drop_deinit_apartment.get() {
            // deinit_apartment() // TODO: This panics for me in tests, why?
        }
    }
}
