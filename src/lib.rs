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
    static SERVICE: RefCell<Result<VirtualDesktopService, Error>> = RefCell::new(Err(Error::NullPtr));
}

fn with_service<T, F>(cb: F) -> Result<T, Error>
where
    F: Fn(&VirtualDesktopService) -> Result<T, Error>,
{
    SERVICE.with(|f| {
        #[allow(unused_must_use)]
        let bb = f.borrow();

        // This first tries to allocate normal VirtualDesktopService
        match bb.as_ref() {
            Err(e) => {
                // Dropping bb allows to replace the value again
                drop(bb);

                #[cfg(feature = "debug")]
                println!("Service is not allocated, try to allocate without COM");

                #[allow(unused_must_use)]
                {
                    f.replace(VirtualDesktopService::create());
                }

                let bb = f.borrow();
                match bb.as_ref() {
                    Ok(v) => cb(v),
                    Err(Error::ComNotInitialized) => {
                        // Dropping bb allows to replace the value again
                        drop(bb);

                        #[cfg(feature = "debug")]
                        println!("Com was not initialized, try to initialize COM first");

                        init_apartment(ApartmentType::Multithreaded).map_err(HRESULT::from_i32)?;

                        #[allow(unused_must_use)]
                        {
                            f.replace(VirtualDesktopService::create());
                        }
                        let bb = f.borrow();
                        let b = bb.as_ref();
                        match b {
                            Err(v) => Err(v.clone()),
                            Ok(v) => cb(v),
                        }
                    }
                    Err(v) => Err(v.clone()),
                }
            }
            Ok(v) => cb(v),
        }
    })
}

pub fn recreate_listener() -> Result<(), Error> {
    // SOMETHING.store(None);
    // match EVENTS {
    //     Ok(v) => Ok(()),
    //     Err(v) => Err(v),
    // }
    Err(Error::NullPtr)
}

pub fn get_listener() -> Result<Receiver<VirtualDesktopEvent>, Error> {
    Err(Error::NullPtr)
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
