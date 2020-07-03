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
use crossbeam_channel::{unbounded, Receiver, Sender};

use service::VirtualDesktopService;
use std::cell::RefCell;
use std::{
    sync::atomic::{AtomicPtr, Ordering},
    thread,
};

pub use changelistener::VirtualDesktopEvent;
pub use desktopid::DesktopID;
pub use error::Error;
pub use hresult::HRESULT;
pub use interfaces::HWND;
use once_cell::sync::Lazy;

// Notice that VirtualDesktopService, and all ComRc types are not thread-safe,
// so we must allocate VirtualDesktopService per thread.
thread_local! {
    static SERVICE: RefCell<Result<VirtualDesktopService, Error>> = RefCell::new(Err(Error::ServiceNotCreated));
}

static EVENTLISTENER: Lazy<EventListener> = Lazy::new(EventListener::new);

// static EVENTS: Lazy<(Sender<VirtualDesktopEvent>, Receiver<VirtualDesktopEvent>)> =
//     Lazy::new(unbounded);

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
    println!("{:?} thread: {:?}", error, std::thread::current().id());

    if retry == 0 {
        return Err(Error::ServiceNotCreated);
    }
    match error {
        Error::ServiceNotCreated => {
            // Try to reinit
            #[allow(unused_must_use)]
            {
                service.replace(VirtualDesktopService::create());
            }
            recreate(cb, retry)
        }
        Error::ComNotInitialized => {
            init_apartment(ApartmentType::Multithreaded).map_err(HRESULT::from_i32)?;
            // Try to reinit
            #[allow(unused_must_use)]
            {
                service.replace(Err(Error::ServiceNotCreated));
            }
            recreate(cb, retry)
        }
        Error::ComRpcUnavailable | Error::ComClassNotRegistered => {
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
        let last_service = service.borrow();
        match last_service.as_ref() {
            Err(er) => {
                let e = er.clone();
                // Drop is important! Otherwise this will give borrow panics
                drop(last_service);
                errorhandler(service, e, cb, retry - 1)
            }
            Ok(v) => match cb(v) {
                Ok(v) => Ok(v),
                Err(er) => {
                    // Drop is important! Otherwise this will give borrow panics
                    drop(last_service);
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

/// Should be called when explorer is restarted
pub fn notify_explorer_restarted() -> Result<(), Error> {
    SERVICE.with(|service| {
        // let f = service.borrow();
        // let ff = f.as_ref().unwrap();
        // ff.
        errorhandler(service, Error::ServiceNotCreated, |_| Ok(()), 6)
    })
}

pub fn get_listener() -> Result<Receiver<VirtualDesktopEvent>, Error> {
    Err(Error::ComAllocatedNullPtr)
    // Ok(EVENTS.1.clone())
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
