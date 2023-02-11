mod comapi;
mod error;
mod hresult;

pub use comapi::normal::*;
pub use comapi::numbered::*;
pub use comapi::windowing::*;
pub use error::Error;
pub(crate) use hresult::HRESULT;

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
    use super::*;
    use once_cell::sync::Lazy;
    use std::sync::{Mutex, Once};
    use std::thread;
    use std::time::Duration;
    use winapi::um::winuser::FindWindowW;

    static INIT: Once = Once::new();

    type HWND = u32;

    // Run the tests synchronously
    fn sync_test<T>(test: T)
    where
        T: FnOnce() -> (),
    {
        static SEMAPHORE: Lazy<Mutex<()>> = Lazy::new(|| Mutex::new(()));

        /*
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
         */

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
                // std::thread::sleep(Duration::from_millis(4));
                go_to_desktop_number(1).unwrap();
            }
            // std::thread::sleep(Duration::from_millis(3));
            go_to_desktop_number(current_desktop).unwrap();
        })
    }

    #[test]
    fn test_desktop_get() {
        sync_test(|| {
            let desktop = get_desktop_by_index(0).unwrap();
            let id = desktop.get_id();

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
