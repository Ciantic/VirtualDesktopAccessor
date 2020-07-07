mod changelistener;
mod comhelpers;
mod desktopid;
mod error;
mod hresult;
mod hstring;
mod immersive;
mod interfaces;
mod service;
use crate::comhelpers::ComError;
use crate::desktopid::DesktopID;
use crate::service::VirtualDesktopService;
use com::runtime::init_runtime;
use crossbeam_channel::{unbounded, Receiver, Sender};
use once_cell::sync::Lazy;
use std::cell::{Ref, RefCell};
use std::sync::{
    atomic::{AtomicBool, Ordering},
    Mutex,
};

pub use crate::changelistener::VirtualDesktopEvent;
pub use crate::error::Error;
pub use crate::hresult::HRESULT;
pub use crate::interfaces::HWND;

static SERVICE: Lazy<Mutex<RefCell<Result<Box<VirtualDesktopService>, Error>>>> =
    Lazy::new(|| Mutex::new(RefCell::new(Err(Error::ServiceNotCreated))));

static EVENTS: Lazy<(Sender<VirtualDesktopEvent>, Receiver<VirtualDesktopEvent>)> =
    Lazy::new(unbounded);

static HAS_LISTENERS: AtomicBool = AtomicBool::new(false);

fn error_side_effect(err: &Error) -> Result<bool, Error> {
    match err {
        Error::ComError(hresult) => {
            let comerror = ComError::from(*hresult);

            #[cfg(feature = "debug")]
            println!("ComError::{:?}", comerror);

            match comerror {
                ComError::NotInitialized => {
                    #[cfg(feature = "debug")]
                    println!("Com initialize");

                    init_runtime().map_err(HRESULT::from_i32)?;
                    // init_apartment(ApartmentType::Multithreaded).map_err(HRESULT::from_i32)?;

                    Ok(true)
                }
                ComError::ClassNotRegistered => Ok(true),
                ComError::RpcUnavailable => Ok(true),
                ComError::ObjectNotConnected => Ok(true),
                ComError::Unknown(_) => Ok(false),
            }
        }
        Error::ServiceNotCreated => Ok(true),
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
                let service_ref: Ref<Result<Box<VirtualDesktopService>, Error>> = cell.borrow();
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
                let _ = cell.replace(VirtualDesktopService::create());
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
pub(crate) fn get_desktops() -> Result<Vec<DesktopID>, Error> {
    with_service(|s| s.get_desktops())
}

/// Get index of a desktop
pub(crate) fn get_index_by_desktop(desktop: DesktopID) -> Result<usize, Error> {
    with_service(|s| s.get_index_by_desktop(desktop))
}

/// Get desktop names
pub fn get_desktop_names() -> Result<Vec<String>, Error> {
    with_service(|s| s.get_desktop_names())
}

/// Get number of desktops
pub fn get_desktop_count() -> Result<u32, Error> {
    with_service(|s| s.get_desktop_count())
}

/// Get current desktop number
pub fn get_current_desktop() -> Result<usize, Error> {
    with_service(|s| s.get_index_by_desktop(s.get_current_desktop()?))
}

/// Get desktop number by window
pub fn get_desktop_by_window(hwnd: HWND) -> Result<usize, Error> {
    with_service(|s| s.get_index_by_desktop(s.get_desktop_by_window(hwnd)?))
}

/// Is window on current virtual desktop
pub fn is_window_on_current_virtual_desktop(hwnd: HWND) -> Result<bool, Error> {
    with_service(|s| s.is_window_on_current_virtual_desktop(hwnd))
}

/// Is window on desktop number
pub fn is_window_on_desktop(hwnd: HWND, desktop: usize) -> Result<bool, Error> {
    with_service(|s| s.is_window_on_desktop(hwnd, &s.get_desktop_by_index(desktop)?))
}

/// Rename desktop
pub fn rename_desktop(desktop: usize, name: &str) -> Result<(), Error> {
    with_service(|s| s.rename_desktop(s.get_desktop_by_index(desktop)?, name))
}

/// Move window to desktop number
pub fn move_window_to_desktop(hwnd: HWND, desktop: usize) -> Result<(), Error> {
    with_service(|s| s.move_window_to_desktop(hwnd, &s.get_desktop_by_index(desktop)?))
}

/// Go to desktop number
pub fn go_to_desktop(desktop: usize) -> Result<(), Error> {
    with_service(|s| s.go_to_desktop(&s.get_desktop_by_index(desktop)?))
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

#[cfg(test)]
mod tests {
    use super::*;
    use std::time::Duration;

    #[test]
    fn test_threads() {
        std::thread::spawn(|| {
            let get_count = || {
                get_desktop_count().unwrap();
            };
            let mut threads = vec![];
            for _ in 0..16 {
                threads.push(std::thread::spawn(get_count));
            }
            for t in threads {
                t.join().unwrap();
            }
        })
        .join()
        .unwrap();
    }

    #[test]
    fn test_desktop_moves() {
        let current_desktop = get_current_desktop().unwrap();

        // Go to desktop 0, ensure it worked
        go_to_desktop(0).unwrap();
        assert_eq!(get_current_desktop().unwrap(), 0);
        std::thread::sleep(Duration::from_secs(1));

        // Go to desktop 1, ensure it worked
        go_to_desktop(1).unwrap();
        assert_eq!(get_current_desktop().unwrap(), 1);
        std::thread::sleep(Duration::from_secs(1));

        // Go to desktop where it was, ensure it worked
        go_to_desktop(current_desktop).unwrap();
        assert_eq!(get_current_desktop().unwrap(), current_desktop);
        std::thread::sleep(Duration::from_secs(1));
    }
}
