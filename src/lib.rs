mod changelistener;
mod comhelpers;
mod desktopid;
mod error;
mod hresult;
mod immersive;
mod interfaces;
mod service;
use changelistener::EventListener;
use com::runtime::{init_apartment, ApartmentType};
use com::{ComInterface, ComRc};
use crossbeam_channel::{unbounded, Receiver, Sender};

use service::VirtualDesktopService;
use std::cell::RefCell;
use std::{
    cell::Cell,
    rc::Rc,
    sync::{Arc, LockResult, Mutex, RwLock},
    thread,
};
use thread::JoinHandle;

pub use changelistener::VirtualDesktopEvent;
pub use desktopid::DesktopID;
pub use error::Error;
pub use hresult::HRESULT;
pub use interfaces::HWND;

// Notice that VirtualDesktopService, and all ComRc types are not thread-safe,
// so we must allocate VirtualDesktopService per thread.
thread_local! {
    static SERVICE: RefCell<Result<VirtualDesktopService, Error>> = RefCell::new(Err(Error::ServiceNotCreated));
}

fn errorhandler<T, F>(
    service: &RefCell<Result<VirtualDesktopService, Error>>,
    error: Error,
    cb: F,
    retry: u32,
) -> Result<T, Error>
where
    F: Fn(&VirtualDesktopService) -> Result<T, Error>,
{
    #[cfg(feature = "debug")]
    println!("Try to error correcting: {:?}", error);
    if retry == 0 {
        return Err(error);
    }
    match error {
        Error::ServiceNotCreated => {
            #[cfg(feature = "debug")]
            println!("Service is not created ...");

            #[allow(unused_must_use)]
            {
                service.replace(VirtualDesktopService::create());
            }
            recreate(cb, retry)
        }
        Error::ComNotInitialized => {
            #[cfg(feature = "debug")]
            println!("Init com apartment, and retry...");

            init_apartment(ApartmentType::Multithreaded).map_err(HRESULT::from_i32)?;

            // Try to reinit
            #[allow(unused_must_use)]
            {
                service.replace(Err(Error::ServiceNotCreated));
            }

            recreate(cb, retry)
        }
        Error::ComRpcUnavailable | Error::ComClassNotRegistered => {
            #[cfg(feature = "debug")]
            println!("RPC Went away, try to recreate...");

            // Try to reinit
            #[allow(unused_must_use)]
            {
                service.replace(Err(Error::ServiceNotCreated));
            }
            recreate(cb, retry)
        }
        e => Err(e),
    }
}

fn recreate<T, F>(cb: F, retry: u32) -> Result<T, Error>
where
    F: Fn(&VirtualDesktopService) -> Result<T, Error>,
{
    SERVICE.with(|service| {
        let bb = service.borrow();
        let b = bb.as_ref();

        match b {
            Err(er) => {
                let e = er.clone();
                // Drop is important! Otherwise this will give borrow panics on edge cases
                drop(bb);
                errorhandler(service, e, cb, retry - 1)
            }
            Ok(v) => match cb(v) {
                Ok(v) => Ok(v),
                Err(er) => {
                    // Drop is important! Otherwise this will give borrow panics on edge cases
                    drop(bb);
                    errorhandler(service, er, cb, retry - 1)
                }
            },
        }
    })
}

fn with_service<T, F>(cb: F) -> Result<T, Error>
where
    F: Fn(&VirtualDesktopService) -> Result<T, Error>,
{
    recreate(cb, 6)
}

pub fn recreate_listener() -> Result<(), Error> {
    recreate(|_| Ok(()), 3)
}

pub fn get_listener() -> Result<Receiver<VirtualDesktopEvent>, Error> {
    Err(Error::ComAllocatedNullPtr)
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
