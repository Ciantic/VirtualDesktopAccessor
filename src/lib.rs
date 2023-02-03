mod changelistener;
mod comhelpers;
mod desktop;
mod desktopid;
mod error;
mod hresult;
mod hstring;
mod immersive;
mod interfaces;
mod service;

use crate::comhelpers::ComError;
use crate::service::VirtualDesktopService;
use changelistener::RegisteredListener;
use com::runtime::init_apartment;
use com::runtime::init_runtime;
use com::runtime::ApartmentType;
use once_cell::sync::Lazy;
use std::borrow::Borrow;
use std::sync::Arc;
use std::sync::Mutex;

pub mod helpers;
pub use crate::changelistener::{VirtualDesktopEvent, VirtualDesktopEventSender};
pub use crate::desktop::Desktop;
pub use crate::desktopid::DesktopID;
pub use crate::error::Error;
pub use crate::hresult::HRESULT;
pub use crate::interfaces::HWND;

static SERVICE: Lazy<Arc<Mutex<Result<Box<VirtualDesktopService>, Error>>>> =
    Lazy::new(|| Arc::new(Mutex::new(Err(Error::ServiceNotCreated))));

thread_local!(
    static LISTENER: Mutex<Option<RegisteredListener>> = Mutex::new(None);
);

static COM_RUNTIME_INITIALIZED: Lazy<Mutex<bool>> = Lazy::new(|| Mutex::new(false));

fn error_side_effect(err: &Error) -> Result<bool, Error> {
    match err {
        Error::ComError(hresult) => {
            let comerror = ComError::from(*hresult);

            #[cfg(feature = "debug")]
            log_output(&format!("ComError::{:?}", comerror));

            match comerror {
                ComError::NotInitialized => {
                    let mut v = COM_RUNTIME_INITIALIZED.lock().unwrap();
                    if *v {
                        return Ok(true);
                    }

                    // This is the right initialization, it uses
                    // CoIncrementMTAUsage inside, and no CoInitialize function
                    // at all
                    init_runtime().map_err(HRESULT::from_i32)?;
                    *v = true;

                    // Following gives STATUS_ACCESS_VIOLATION in the threading
                    // test, it uses CoInitializeEx with COINIT_MULTITHREADED
                    // inside
                    // init_apartment(ApartmentType::Multithreaded).map_err(HRESULT::from_i32)?;
                    // init_apartment(ApartmentType::SingleThreaded).map_err(HRESULT::from_i32)?;

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
        Ok(mut cell) => {
            // println!("Thread id is {:?}", std::thread::current().id());
            for _ in 0..6 {
                let service_ref = cell.borrow();
                let result = service_ref.as_ref();

                match result {
                    Ok(mut v) => match cb(v) {
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
                };
                (*cell) = Err(Error::ServiceNotCreated);
                let res = VirtualDesktopService::create();
                match res {
                    Ok(new_service) => {
                        #[cfg(feature = "debug")]
                        log_output(&format!("Set service"));

                        // Store service
                        (*cell) = Ok(new_service);
                    }
                    Err(err) => {
                        (*cell) = Err(err);
                    }
                }
            }
            Err(Error::ServiceNotCreated)
        }
        Err(_) => {
            #[cfg(feature = "debug")]
            log_output(&format!("Lock failed: SERVICE.lock()"));
            Err(Error::ServiceNotCreated)
        }
    }
}

/// Should be called when explorer is restarted
pub fn notify_explorer_restarted() -> Result<(), Error> {
    Ok(())
    // if let Ok(mut cell) = SERVICE.lock() {
    //     let old = cell.borrow().as_ref();
    //     match old {
    //         Ok(v) => {
    //             (*cell) = v.recreate();
    //         }
    //         Err(_) => {
    //             (*cell) = VirtualDesktopService::create(None);
    //         }
    //     }
    //     Ok(())
    // } else {
    //     Ok(())
    // }
}

pub fn set_event_sender(sender: VirtualDesktopEventSender) -> Result<(), Error> {
    std::thread::spawn(|| {
        init_runtime().map_err(HRESULT::from_i32).unwrap();
        LISTENER.with(|f| {
            let mut a = f.lock().unwrap();
            *a = Some(VirtualDesktopService::create_listener(Some(sender)).unwrap());
        });
        loop {
            std::thread::sleep(std::time::Duration::from_millis(1000));
        }
    });

    // println!("Create event listener");
    // let _ = with_service(move |s| Ok(s.set_event_sender(sender.clone())));
    Ok(())
}

/// Get desktop name
pub fn get_desktop_name(desktop: &Desktop) -> Result<String, Error> {
    with_service(|s| s.get_desktop_name(desktop))
}

/// Get desktop name
pub(crate) fn get_index_by_desktop(desktop: &Desktop) -> Result<usize, Error> {
    with_service(|s| s.get_index_by_desktop(desktop))
}

/// Set desktop name
pub fn set_desktop_name(desktop: &Desktop, name: &str) -> Result<(), Error> {
    with_service(|s| s.set_desktop_name(desktop, name))
}

/// Get desktop number
pub fn get_desktop_by_index(number: usize) -> Result<Desktop, Error> {
    with_service(|s| s.get_desktop_by_index(number))
}

/// Get desktop by GUID
pub fn get_desktop_by_guid(id: &DesktopID) -> Result<Desktop, Error> {
    with_service(|s| s.get_desktop_by_guid(&id))
}

/// Get desktops
pub fn get_desktops() -> Result<Vec<Desktop>, Error> {
    with_service(|s| s.get_desktops())
}

/// Get current desktop
pub fn get_current_desktop() -> Result<Desktop, Error> {
    with_service(|s| s.get_current_desktop())
}

/// Get desktop by window
pub fn get_desktop_by_window(hwnd: HWND) -> Result<Desktop, Error> {
    with_service(|s| s.get_desktop_by_window(hwnd))
}

/// Is window on desktop number
pub fn is_window_on_desktop(hwnd: HWND, desktop: &Desktop) -> Result<bool, Error> {
    with_service(|s| s.is_window_on_desktop(hwnd, &desktop))
}

/// Move window to desktop
pub fn move_window_to_desktop(hwnd: HWND, desktop: &Desktop) -> Result<(), Error> {
    with_service(|s| s.move_window_to_desktop(hwnd, desktop))
}

/// Go to desktop
pub fn go_to_desktop(desktop: &Desktop) -> Result<(), Error> {
    with_service(|s| s.go_to_desktop(desktop))
}

/// Create desktop
pub fn create_desktop() -> Result<Desktop, Error> {
    with_service(|s| s.create_desktop())
}

/// Remove desktop
pub fn remove_desktop(remove_desktop: &Desktop, fallback_desktop: &Desktop) -> Result<(), Error> {
    with_service(|s| s.remove_desktop(remove_desktop, fallback_desktop))
}

/// Is window on current  desktop
pub fn is_window_on_current_desktop(hwnd: HWND) -> Result<bool, Error> {
    with_service(|s| s.is_window_on_current_desktop(hwnd))
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

/// Is pinned app?
pub fn is_pinned_app(hwnd: HWND) -> Result<bool, Error> {
    with_service(|s| s.is_pinned_app(hwnd))
}

/// Pin entire app and all it's windows
pub fn pin_app(hwnd: HWND) -> Result<(), Error> {
    with_service(|s| s.pin_app(hwnd))
}

/// Unpin entire app and all it's windows
pub fn unpin_app(hwnd: HWND) -> Result<(), Error> {
    with_service(|s| s.unpin_app(hwnd))
}

// Import OutputDebugStringA
#[cfg(feature = "debug")]
extern "system" {
    fn OutputDebugStringA(lpOutputString: *const i8);
}

#[cfg(feature = "debug")]
pub(crate) fn log_output(s: &str) {
    unsafe {
        println!("{}", s);
        OutputDebugStringA(s.as_ptr() as *const i8);
    }
}

#[cfg(test)]
mod tests {
    use super::helpers::*;
    use super::*;
    use std::sync::Once;
    use std::thread;
    use std::time::Duration;
    use winapi::um::winuser::FindWindowW;

    static INIT: Once = Once::new();

    // Run the tests synchronously
    fn sync_test<T>(test: T)
    where
        T: FnOnce() -> (),
    {
        static SEMAPHORE: Lazy<Mutex<()>> = Lazy::new(|| Mutex::new(()));

        INIT.call_once(|| {
            let (a, b) = std::sync::mpsc::channel::<VirtualDesktopEvent>();

            set_event_sender(changelistener::VirtualDesktopEventSender::Std(a.clone())).unwrap();

            thread::spawn(move || {
                b.iter().for_each(|msg| match msg {
                    VirtualDesktopEvent::DesktopChanged(old, new) => {
                        println!(
                            "<- Desktop changed from {:?} to {:?}",
                            old.get_index().unwrap(),
                            new.get_index().unwrap()
                        );
                    }
                    VirtualDesktopEvent::DesktopCreated(desk) => {
                        println!("<- New desktop created {:?}", desk);
                    }
                    VirtualDesktopEvent::DesktopDestroyed(desk) => {
                        println!("<- Desktop destroyed {:?}", desk);
                    }
                    VirtualDesktopEvent::WindowChanged(hwnd) => {
                        println!("<- Window changed {:?}", hwnd);
                    }
                    VirtualDesktopEvent::DesktopNameChanged(desk, name) => {
                        println!("<- Name of {:?} changed to {}", desk, name);
                    }
                    VirtualDesktopEvent::DesktopWallpaperChanged(desk, name) => {
                        println!("<- Wallpaper of {:?} changed to {}", desk, name);
                    }
                    VirtualDesktopEvent::DesktopMoved(desk, old, new) => {
                        println!("<- Desktop {:?} moved from {} to {}", desk, old, new);
                    }
                });
            });
        });

        let _t = SEMAPHORE.lock().unwrap();
        test()
    }

    #[test]
    fn test_threads() {
        sync_test(|| {
            std::thread::spawn(|| {
                // let get_count = || {
                //     get_desktop_count().unwrap();
                // };
                let mut threads = vec![];
                for _ in 0..555 {
                    threads.push(std::thread::spawn(|| {
                        get_desktops().unwrap().iter().for_each(|d| {
                            let n = d.get_name().unwrap();
                            let i = d.get_index().unwrap();
                            let j = d.get_index().unwrap();
                            println!("Thread {n} {i} {j} {:?}", std::thread::current().id());
                        })
                    }));
                }
                thread::sleep(Duration::from_millis(2500));
                for t in threads {
                    t.join().unwrap();
                }
            })
            .join()
            .unwrap();
        })
    }

    #[test] // TODO: Commented out, use only on occasion when needed!
    fn test_threading_two() {
        sync_test(|| {
            let current_desktop = get_current_desktop_number().unwrap();

            for _ in 0..999 {
                go_to_desktop_number(0).unwrap();
                std::thread::sleep(Duration::from_millis(4));
                go_to_desktop_number(1).unwrap();
            }
            std::thread::sleep(Duration::from_millis(3));
            go_to_desktop_number(current_desktop).unwrap();
        })
    }

    #[test]
    fn test_desktop_get() {
        sync_test(|| {
            let desktop = get_desktop_by_index(0).unwrap();
            let id = desktop.get_id();
            let (data1, _, _, _) = id.get_data();
            assert_ne!(data1, 0);

            // No errors by getting desktop
            get_desktop_by_guid(&id).unwrap();
        })
    }

    #[test]
    fn test_desktop_moves() {
        sync_test(|| {
            let current_desktop = get_current_desktop_number().unwrap();

            // Go to desktop 0, ensure it worked
            go_to_desktop_number(0).unwrap();
            assert_eq!(get_current_desktop_number().unwrap(), 0);
            std::thread::sleep(Duration::from_millis(400));

            // Go to desktop 1, ensure it worked
            go_to_desktop_number(1).unwrap();
            assert_eq!(get_current_desktop_number().unwrap(), 1);
            std::thread::sleep(Duration::from_millis(400));

            // Go to desktop where it was, ensure it worked
            go_to_desktop_number(current_desktop).unwrap();
            assert_eq!(get_current_desktop_number().unwrap(), current_desktop);
            std::thread::sleep(Duration::from_millis(400));
        })
    }

    #[test]
    fn test_move_notepad_between_desktops() {
        sync_test(|| {
            // Get notepad
            let notepad_hwnd: HWND = unsafe {
                let notepad = "notepad\0".encode_utf16().collect::<Vec<_>>();
                FindWindowW(notepad.as_ptr(), std::ptr::null()) as HWND
            };
            assert!(
                notepad_hwnd != 0,
                "Notepad requires to be running for this test"
            );

            let current_desktop = get_current_desktop_number().unwrap();
            assert!(current_desktop != 0, "Current desktop must not be 0");

            let notepad_is_on_current_desktop = is_window_on_current_desktop(notepad_hwnd).unwrap();
            let notepad_is_on_specific_desktop =
                is_window_on_desktop_number(notepad_hwnd, current_desktop).unwrap();
            assert!(
                notepad_is_on_current_desktop,
                "Notepad must be on this desktop"
            );
            assert!(
                notepad_is_on_specific_desktop,
                "Notepad must be on this desktop"
            );

            // Move notepad current desktop -> 0 -> 1 -> current desktop
            move_window_to_desktop_number(notepad_hwnd, 0).unwrap();
            let notepad_desktop = get_desktop_number_by_window(notepad_hwnd).unwrap();
            assert_eq!(notepad_desktop, 0, "Notepad should have moved to desktop 0");
            std::thread::sleep(Duration::from_millis(300));

            move_window_to_desktop_number(notepad_hwnd, 1).unwrap();
            let notepad_desktop = get_desktop_number_by_window(notepad_hwnd).unwrap();
            assert_eq!(notepad_desktop, 1, "Notepad should have moved to desktop 1");
            std::thread::sleep(Duration::from_millis(300));

            move_window_to_desktop_number(notepad_hwnd, current_desktop).unwrap();
            let notepad_desktop = get_desktop_number_by_window(notepad_hwnd).unwrap();
            assert_eq!(
                notepad_desktop, current_desktop,
                "Notepad should have moved to desktop 0"
            );
        })
    }

    #[test]
    fn test_pin_notepad() {
        sync_test(|| {
            // Get notepad
            let notepad_hwnd: HWND = unsafe {
                let notepad = "notepad\0".encode_utf16().collect::<Vec<_>>();
                FindWindowW(notepad.as_ptr(), std::ptr::null()) as HWND
            };
            assert!(
                notepad_hwnd != 0,
                "Notepad requires to be running for this test"
            );
            assert_eq!(
                is_window_on_current_desktop(notepad_hwnd).unwrap(),
                true,
                "Notepad must be on current desktop to test this"
            );

            assert_eq!(
                is_pinned_window(notepad_hwnd).unwrap(),
                false,
                "Notepad must not be pinned at the start of the test"
            );

            let current_desktop = get_current_desktop_number().unwrap();

            // Pin notepad and go to desktop 0 and back
            pin_window(notepad_hwnd).unwrap();
            go_to_desktop_number(0).unwrap();

            assert_eq!(is_pinned_window(notepad_hwnd).unwrap(), true);
            std::thread::sleep(Duration::from_millis(1000));

            go_to_desktop_number(current_desktop).unwrap();
            unpin_window(notepad_hwnd).unwrap();
            assert_eq!(
                is_window_on_desktop_number(notepad_hwnd, current_desktop).unwrap(),
                true
            );
            std::thread::sleep(Duration::from_millis(1000));
        })
    }

    #[test]
    fn test_pin_notepad_app() {
        sync_test(|| {
            // Get notepad
            let notepad_hwnd: HWND = unsafe {
                let notepad = "notepad\0".encode_utf16().collect::<Vec<_>>();
                FindWindowW(notepad.as_ptr(), std::ptr::null()) as HWND
            };
            assert!(
                notepad_hwnd != 0,
                "Notepad requires to be running for this test"
            );
            assert_eq!(
                is_window_on_current_desktop(notepad_hwnd).unwrap(),
                true,
                "Notepad must be on current desktop to test this"
            );

            assert_eq!(
                is_pinned_app(notepad_hwnd).unwrap(),
                false,
                "Notepad must not be pinned at the start of the test"
            );

            let current_desktop = get_current_desktop_number().unwrap();

            // Pin notepad and go to desktop 0 and back
            pin_app(notepad_hwnd).unwrap();
            assert_eq!(is_pinned_app(notepad_hwnd).unwrap(), true);

            go_to_desktop_number(0).unwrap();
            std::thread::sleep(Duration::from_millis(1000));
            go_to_desktop_number(current_desktop).unwrap();

            unpin_app(notepad_hwnd).unwrap();
            assert_eq!(
                is_window_on_desktop_number(notepad_hwnd, current_desktop).unwrap(),
                true
            );
            std::thread::sleep(Duration::from_millis(1000));
        })
    }

    /// Rename first desktop to Foo, and then back to what it was
    #[test]
    fn test_rename_desktop() {
        let desktops = get_desktops().unwrap();
        let first_desktop = desktops.get(0).take().unwrap();
        let first_desktop_name_before = first_desktop.get_name().unwrap();

        // Pre-condition
        assert_ne!(
            first_desktop_name_before, "Example Desktop",
            "Your first desktop must be something else than \"Example Desktop\" to run this test."
        );

        // Rename
        first_desktop.set_name("Example Desktop").unwrap();

        // Ensure it worked
        assert_eq!(
            first_desktop.get_name().unwrap(),
            "Example Desktop",
            "Rename failed"
        );

        // Return to normal
        first_desktop.set_name(&first_desktop_name_before).unwrap();
    }

    /// Test some errors
    #[test]
    fn test_errors() {
        let err = set_name_by_desktop_number(99999, "").unwrap_err();
        assert_eq!(err, Error::DesktopNotFound);

        let err = go_to_desktop_number(99999).unwrap_err();
        assert_eq!(err, Error::DesktopNotFound);

        let err = get_desktop_number_by_window(9999999).unwrap_err();
        assert_eq!(err, Error::WindowNotFound);

        let err = move_window_to_desktop_number(0, 99999).unwrap_err();
        assert_eq!(err, Error::DesktopNotFound);

        let err = move_window_to_desktop_number(999999, 0).unwrap_err();
        assert_eq!(err, Error::WindowNotFound);
    }
}
