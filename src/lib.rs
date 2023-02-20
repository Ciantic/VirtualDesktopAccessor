//! winvd - crate for accessing the Windows Virtual Desktop API
//!
//! All functions taking `Into<Desktop>` can take either a index or a GUID.
//!
//! # Examples
//! * Get first desktop name by index `get_desktop(0).get_name().unwrap()`
//! * Get second desktop name by index `get_desktop(1).get_name().unwrap()`
//! * Get desktop name by GUID `get_desktop(GUID(123...)).get_name().unwrap()`
//! * Switch to fifth desktop by index `switch_desktop(4).unwrap()`
//! * Get third desktop name `get_desktop(2).get_name().unwrap()`
mod comobjects;
mod desktop;
mod events;
mod hresult;
mod interfaces;
mod listener;

pub use comobjects::Error;
pub use desktop::*;
pub use events::*;
pub type Result<T> = std::result::Result<T, Error>;

// Import OutputDebugStringA
#[cfg(debug_assertions)]
extern "system" {
    fn OutputDebugStringA(lpOutputString: *const i8);
}

#[cfg(debug_assertions)]
pub(crate) fn log_output(s: &str) {
    unsafe {
        println!("{}", s);
        OutputDebugStringA(s.as_ptr() as *const i8);
    }
}

#[cfg(not(debug_assertions))]
#[inline]
pub(crate) fn log_output(_s: &str) {}

// Log format macro
#[macro_export]
macro_rules! log_format {
    ($($arg:tt)*) => {
        #[cfg(debug_assertions)]
        $crate::log_output(&format!($($arg)*));
    };
}

#[cfg(test)]
mod tests {
    use super::*;
    use once_cell::sync::Lazy;
    use std::sync::{Mutex, Once};
    use std::thread;
    use std::time::Duration;
    use windows::core::PCWSTR;
    use windows::Win32::Foundation::HWND;
    use windows::Win32::UI::WindowsAndMessaging::FindWindowW;

    static INIT: Once = Once::new();

    // Run the tests synchronously
    fn sync_test<T>(test: T)
    where
        T: FnOnce() -> (),
    {
        static SEMAPHORE: Lazy<Mutex<()>> = Lazy::new(|| Mutex::new(()));

        // !! TODO: Start a listener

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
                            println!("Thread {n} {i} {:?}", std::thread::current().id());
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
            let current_desktop = get_current_desktop().unwrap();

            for _ in 0..999 {
                switch_desktop(0).unwrap();
                // std::thread::sleep(Duration::from_millis(4));
                switch_desktop(1).unwrap();
            }
            std::thread::sleep(Duration::from_millis(15));
            switch_desktop(current_desktop).unwrap();
        })
    }

    #[test]
    fn test_desktop_get() {
        sync_test(|| {
            let desktop = get_desktop(0).get_id().unwrap();
            get_desktop(&desktop).get_index().unwrap();
        })
    }

    #[test]
    fn test_desktop_moves() {
        sync_test(|| {
            let current_desktop = get_current_desktop().unwrap().get_index().unwrap();

            // Go to desktop 0, ensure it worked
            switch_desktop(0).unwrap();
            assert_eq!(get_current_desktop().unwrap().get_index().unwrap(), 0);
            std::thread::sleep(Duration::from_millis(400));

            // Go to desktop 1, ensure it worked
            switch_desktop(1).unwrap();
            assert_eq!(get_current_desktop().unwrap().get_index().unwrap(), 1);
            std::thread::sleep(Duration::from_millis(400));

            // Go to desktop where it was, ensure it worked
            switch_desktop(current_desktop).unwrap();
            assert_eq!(
                get_current_desktop().unwrap().get_index().unwrap(),
                current_desktop
            );
            std::thread::sleep(Duration::from_millis(400));
        })
    }

    #[test]
    fn test_move_notepad_between_desktops() {
        sync_test(|| {
            // Get notepad
            let notepad_hwnd = unsafe {
                let notepad = "notepad\0".encode_utf16().collect::<Vec<_>>();
                let pw = PCWSTR::from_raw(notepad.as_ptr());
                FindWindowW(pw, PCWSTR::null())
            };
            assert!(
                notepad_hwnd != HWND::default(),
                "Notepad requires to be running for this test"
            );

            let current_desktop = get_current_desktop().unwrap();
            assert!(
                0 != current_desktop.get_index().unwrap(),
                "Current desktop must not be 0"
            );

            let notepad_is_on_current_desktop = is_window_on_current_desktop(notepad_hwnd).unwrap();
            let notepad_is_on_specific_desktop =
                is_window_on_desktop(current_desktop, notepad_hwnd).unwrap();
            assert!(
                notepad_is_on_current_desktop,
                "Notepad must be on this desktop"
            );
            assert!(
                notepad_is_on_specific_desktop,
                "Notepad must be on this desktop"
            );

            // Move notepad current desktop -> 0 -> 1 -> current desktop
            move_window_to_desktop(0, &notepad_hwnd).unwrap();
            let notepad_desktop = get_desktop_by_window(notepad_hwnd)
                .unwrap()
                .get_index()
                .unwrap();
            assert_eq!(notepad_desktop, 0, "Notepad should have moved to desktop 0");
            std::thread::sleep(Duration::from_millis(300));

            move_window_to_desktop(1, &notepad_hwnd).unwrap();
            let notepad_desktop = get_desktop_by_window(notepad_hwnd)
                .unwrap()
                .get_index()
                .unwrap();
            assert_eq!(notepad_desktop, 1, "Notepad should have moved to desktop 1");
            std::thread::sleep(Duration::from_millis(300));

            move_window_to_desktop(current_desktop, &notepad_hwnd).unwrap();
            let notepad_desktop = get_desktop_by_window(notepad_hwnd).unwrap();
            assert!(
                notepad_desktop.try_eq(&current_desktop).unwrap(),
                "Notepad should have moved to desktop 0"
            );
        })
    }

    #[test]
    fn test_pin_notepad() {
        sync_test(|| {
            // Get notepad
            let notepad_hwnd = unsafe {
                let notepad = "notepad\0".encode_utf16().collect::<Vec<_>>();
                let pw = PCWSTR::from_raw(notepad.as_ptr());
                FindWindowW(pw, PCWSTR::null())
            };
            assert!(
                notepad_hwnd != HWND::default(),
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

            let current_desktop = get_current_desktop().unwrap();

            // Pin notepad and go to desktop 0 and back
            pin_window(notepad_hwnd).unwrap();
            switch_desktop(0).unwrap();

            assert_eq!(is_pinned_window(notepad_hwnd).unwrap(), true);
            std::thread::sleep(Duration::from_millis(1000));

            switch_desktop(current_desktop).unwrap();
            unpin_window(notepad_hwnd).unwrap();
            assert_eq!(
                is_window_on_desktop(current_desktop, notepad_hwnd).unwrap(),
                true
            );
            std::thread::sleep(Duration::from_millis(1000));
        })
    }

    #[test]
    fn test_pin_notepad_app() {
        sync_test(|| {
            // Get notepad
            let notepad_hwnd = unsafe {
                let notepad = "notepad\0".encode_utf16().collect::<Vec<_>>();
                let pw = PCWSTR::from_raw(notepad.as_ptr());
                FindWindowW(pw, PCWSTR::null())
            };
            assert!(
                notepad_hwnd != HWND::default(),
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

            let current_desktop = get_current_desktop().unwrap();

            // Pin notepad and go to desktop 0 and back
            pin_app(notepad_hwnd).unwrap();
            assert_eq!(is_pinned_app(notepad_hwnd).unwrap(), true);

            switch_desktop(0).unwrap();
            std::thread::sleep(Duration::from_millis(1000));
            switch_desktop(current_desktop).unwrap();

            unpin_app(notepad_hwnd).unwrap();
            assert_eq!(
                is_window_on_desktop(current_desktop, notepad_hwnd).unwrap(),
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
        let err = get_desktop(99999).set_name("").unwrap_err();
        assert_eq!(err, Error::DesktopNotFound);

        let err = switch_desktop(99999).unwrap_err();
        assert_eq!(err, Error::DesktopNotFound);

        let err = get_desktop_by_window(HWND(9999999)).unwrap_err();
        assert_eq!(err, Error::WindowNotFound);

        let err = move_window_to_desktop(99999, &HWND::default()).unwrap_err();
        assert_eq!(err, Error::WindowNotFound);

        let err = move_window_to_desktop(0, &HWND(999999)).unwrap_err();
        assert_eq!(err, Error::WindowNotFound);
    }
}
