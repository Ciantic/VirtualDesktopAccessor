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
mod immersive;
mod interfaces;
use com::runtime::{
    deinit_apartment, get_class_object, init_apartment, init_runtime, ApartmentType,
};
use com::{
    co_class,
    interfaces::IUnknown,
    sys::{CoCreateInstance, CLSCTX_INPROC_SERVER, FAILED, GUID, HRESULT, S_OK},
    ComPtr, ComRc, IID,
};

use changelistener::VirtualDesktopChangeListener;
use comhelpers::create_instance;
use immersive::{get_immersive_service, get_immersive_service_for_class};
use interfaces::{
    CLSID_IVirtualNotificationService, CLSID_ImmersiveShell, CLSID_VirtualDesktopManagerInternal,
    IApplicationView, IID_IVirtualDesktopNotification, IServiceProvider, IVirtualDesktop,
    IVirtualDesktopManager, IVirtualDesktopManagerInternal, IVirtualDesktopNotification,
    IVirtualDesktopNotificationService, IVirtualDesktopVTable,
};
use std::cell::Cell;

#[derive(Debug)]
pub enum VirtualDesktopError {
    UnknownError,
    ApartmentInitializationError(HRESULT),
    ImmersiveShellError,
}

struct State {
    haa: u32,
    foo: Cell<Option<Box<dyn Fn() -> ()>>>,
}

pub struct VirtualDesktopService {
    service_provider: ComRc<dyn IServiceProvider>,
    virtual_desktop_manager: ComRc<dyn IVirtualDesktopManager>,
    virtual_desktop_manager_internal: ComRc<dyn IVirtualDesktopManagerInternal>,
    // virtual_desktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService>,
    virtual_desktop_notification_listener: Box<VirtualDesktopChangeListener>,
}

impl VirtualDesktopService {
    pub fn get_desktop_count() -> Result<i32, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn get_desktop_by_number() -> Result<GUID, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn get_current_desktop(&self) -> Result<GUID, VirtualDesktopError> {
        let ptr: *mut IVirtualDesktopVTable = std::ptr::null_mut();

        // debug_print!("ptr {:?}", ptrr as u32);
        // debug_print!("Current desktop ptr {:?}", ptr.as_raw());
        // let mut ptr: u32 = 0;

        let res = unsafe {
            self.virtual_desktop_manager_internal
                .get_current_desktop(&ptr)
        };
        let resultptr: ComPtr<dyn IVirtualDesktop> = unsafe { ComPtr::new(ptr as *mut _) };
        let resultdc = ComRc::new(resultptr);
        let mut desktopid = GUID {
            data1: 0,
            data2: 0,
            data3: 0,
            data4: [0, 0, 0, 0, 0, 0, 0, 0],
        };
        unsafe { resultdc.get_id(&mut desktopid) };
        debug_print!("Current desktop id {:?}", desktopid);

        // let desktop = ComRc::new(ptr);

        Err(VirtualDesktopError::UnknownError)
    }
    pub fn get_window_desktop_id(hwnd: i32) -> Result<GUID, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn is_window_on_current_virtual_desktop(hwnd: i32) -> Result<bool, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }

    // Numbered helpers
    pub fn get_current_desktop_number() -> Result<i32, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
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
    pub fn is_window_on_desktop(hwnd: i32, desktop: GUID) -> Result<bool, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn move_window_to_desktop(hwnd: i32, desktop: GUID) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn go_to_desktop(desktop: GUID) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn is_pinned_window(hwnd: i32) -> Result<bool, VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn pin_window(hwnd: i32) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn unpin_window(hwnd: i32) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn is_pinned_app(hwnd: i32) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn pin_app(hwnd: i32) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
    pub fn unpin_app(hwnd: i32) -> Result<(), VirtualDesktopError> {
        Err(VirtualDesktopError::UnknownError)
    }
}

/// Initialize COM apartment and ImmersiveShell provider service
pub fn initialize() -> Result<VirtualDesktopService, VirtualDesktopError> {
    init_apartment(ApartmentType::Multithreaded)
        .map_err(|op| VirtualDesktopError::ApartmentInitializationError(op))?;
    initialize_only_service()
}

/// Initialize only ImmersiveShell provider service, must be re-called on
/// TaskbarCreated message
pub fn initialize_only_service() -> Result<VirtualDesktopService, VirtualDesktopError> {
    let service_provider = create_instance::<dyn IServiceProvider>(&CLSID_ImmersiveShell).unwrap();
    let virtual_desktop_manager =
        get_immersive_service::<dyn IVirtualDesktopManager>(&service_provider).unwrap();
    let virtualdesktop_notification_service =
        get_immersive_service_for_class::<dyn IVirtualDesktopNotificationService>(
            &service_provider,
            CLSID_IVirtualNotificationService,
        )
        .unwrap();
    let vd_internale =
        get_immersive_service_for_class(&service_provider, CLSID_VirtualDesktopManagerInternal)
            .unwrap();

    let listener =
        VirtualDesktopChangeListener::register(virtualdesktop_notification_service).unwrap();
    Ok(VirtualDesktopService {
        virtual_desktop_manager: virtual_desktop_manager,
        service_provider: service_provider,
        virtual_desktop_notification_listener: listener,
        virtual_desktop_manager_internal: vd_internale,
    })
}
