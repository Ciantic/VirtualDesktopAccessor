mod changelistener;
mod comhelpers;
mod desktopid;
mod error;
mod hresult;
mod immersive;
mod interfaces;
mod service;
use com::runtime::{init_apartment, ApartmentType};
use crossbeam_channel::{unbounded, Receiver, Sender};

use service::VirtualDesktopService;
use std::cell::{Ref, RefCell};
use std::sync::{
    atomic::{AtomicBool, Ordering},
    Mutex,
};

pub use changelistener::VirtualDesktopEvent;
pub use desktopid::DesktopID;
pub use error::Error;
pub use hresult::HRESULT;
pub use interfaces::HWND;
use once_cell::sync::Lazy;

static SERVICE: Lazy<Mutex<RefCell<Result<Box<VirtualDesktopService>, Error>>>> =
    Lazy::new(|| Mutex::new(RefCell::new(Err(Error::ServiceNotCreated))));

static EVENTS: Lazy<(Sender<VirtualDesktopEvent>, Receiver<VirtualDesktopEvent>)> =
    Lazy::new(unbounded);

static HAS_LISTENERS: AtomicBool = AtomicBool::new(false);

fn error_side_effect(err: &Error) -> Result<bool, Error> {
    #[cfg(feature = "debug")]
    println!("{:?}", err);

    match err {
        Error::ComNotInitialized => {
            #[cfg(feature = "debug")]
            println!("Com initialize");
            // init_runtime().map_err(HRESULT::from_i32)?;
            init_apartment(ApartmentType::Multithreaded).map_err(HRESULT::from_i32)?;
            Ok(true)
        }
        Error::ServiceNotCreated | Error::ComRpcUnavailable | Error::ComClassNotRegistered => {
            Ok(true)
        }
        _ => Ok(false),
    }
}

fn with_service<T, F>(cb: F) -> Result<T, Error>
where
    F: Fn(&VirtualDesktopService) -> Result<T, Error>,
{
    match SERVICE.lock() {
        Ok(cell) => {
            for _ in 0..6 {
                let service_ref: Ref<Result<Box<VirtualDesktopService>, Error>> = (*cell).borrow();
                let result = service_ref.as_ref();
                match result {
                    Ok(v) => match cb(&v) {
                        Ok(r) => return Ok(r),
                        Err(err) => match error_side_effect(&err) {
                            Ok(false) => return Err(err),
                            Ok(true) => (),
                            Err(err) => return Err(err),
                        },
                    },
                    Err(err) => {
                        // Ignore
                        #[allow(unused_must_use)]
                        {
                            error_side_effect(&err);
                        }
                    }
                }
                drop(service_ref);
                #[cfg(feature = "debug")]
                println!("Try to create");
                let _ = (*cell).replace(VirtualDesktopService::create());
            }
            Err(Error::ServiceNotCreated)
        }
        Err(_) => {
            #[cfg(feature = "debug")]
            println!("Lock failed?");
            Err(Error::ServiceNotCreated)
        }
    }
}

/// Should be called when explorer is restarted
pub fn notify_explorer_restarted() -> Result<(), Error> {
    if let Ok(cell) = SERVICE.lock() {
        let _ = (*cell).replace(Ok(VirtualDesktopService::create()?));
        Ok(())
    } else {
        Ok(())
    }
}

/// Get event receiver
pub fn get_event_receiver() -> Receiver<VirtualDesktopEvent> {
    let _res = with_service(|s| {
        s.get_event_receiver()?;
        HAS_LISTENERS.store(true, Ordering::SeqCst);
        Ok(())
    });

    EVENTS.1.clone()
}

/// Get desktops
pub fn get_desktops() -> Result<Vec<DesktopID>, Error> {
    with_service(|s| s.get_desktops())
}

/// Get number of desktops
pub fn get_desktop_count() -> Result<u32, Error> {
    with_service(|s| s.get_desktop_count())
}

/// Get current desktop ID
pub fn get_current_desktop() -> Result<DesktopID, Error> {
    with_service(|s| s.get_current_desktop())
}

/// Get window desktop ID
pub fn get_desktop_by_window(hwnd: HWND) -> Result<DesktopID, Error> {
    with_service(|s| s.get_desktop_by_window(hwnd))
}

/// Is window on current virtual desktop
pub fn is_window_on_current_virtual_desktop(hwnd: HWND) -> Result<bool, Error> {
    with_service(|s| s.is_window_on_current_virtual_desktop(hwnd))
}

/// Is window on desktop
pub fn is_window_on_desktop(hwnd: HWND, desktop: &DesktopID) -> Result<bool, Error> {
    with_service(|s| s.is_window_on_desktop(hwnd, desktop))
}

/// Move window to desktop
pub fn move_window_to_desktop(hwnd: HWND, desktop: &DesktopID) -> Result<(), Error> {
    with_service(|s| s.move_window_to_desktop(hwnd, desktop))
}

/// Go to desktop
pub fn go_to_desktop(desktop: &DesktopID) -> Result<(), Error> {
    with_service(|s| s.go_to_desktop(desktop))
}

/// Is window pinned?
pub fn is_pinned_window(hwnd: HWND) -> Result<bool, Error> {
    with_service(|s| s.is_pinned_window(hwnd))
}

/// Pin window
pub fn pin_window(hwnd: HWND) -> Result<(), Error> {
    with_service(|s| s.pin_window(hwnd))
}

/// Unpin window
pub fn unpin_window(hwnd: HWND) -> Result<(), Error> {
    with_service(|s| s.unpin_window(hwnd))
}
