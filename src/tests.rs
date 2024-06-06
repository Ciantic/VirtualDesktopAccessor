use super::*;
use once_cell::sync::Lazy;
use std::sync::{Arc, Mutex, PoisonError};
use std::thread;
use std::time::Duration;
use windows::core::PCWSTR;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::FindWindowW;

static SEMAPHORE: Lazy<Arc<Mutex<u32>>> = Lazy::new(|| Arc::new(Mutex::new(0)));

// Run the tests synchronously
pub fn sync_test<T>(test: T)
where
    T: FnOnce(),
{
    let mut tests_ran = SEMAPHORE.lock().unwrap_or_else(PoisonError::into_inner);
    test();
    *tests_ran += 1;
    drop(tests_ran);
}

#[test]
fn test_desktop_get() {
    sync_test(|| {
        let desktop = get_desktop(0).get_id().unwrap();
        get_desktop(desktop).get_index().unwrap();
    })
}

#[test]
fn test_desktop_moves() {
    sync_test(|| {
        let current_desktop = get_current_desktop().unwrap().get_index().unwrap();

        // Listen for desktop changes
        let (tx, rx) = std::sync::mpsc::channel::<DesktopEvent>();
        let mut _notifications_thread = listen_desktop_events(tx).unwrap();
        let receiver = std::thread::spawn(move || {
            let mut count = 0;
            for item in rx {
                if let DesktopEvent::DesktopChanged { new, old } = item {
                    count += 1;
                    println!(
                        "Desktop changed from {:?} to {:?} count {}",
                        old, new, count
                    );
                }
                if count == 3 {
                    // Stopping the notification thread drops the sender, and iteration ends
                    _notifications_thread.stop().unwrap();
                }
            }
            count
        });

        // Wait for listener to have started
        std::thread::sleep(Duration::from_millis(400));

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

        // Ensure desktop changed three times
        assert_eq!(3, receiver.join().unwrap());
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
            notepad_desktop == current_desktop,
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
    sync_test(|| {
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
    })
}

/// Test some errors
#[test]
fn test_errors() {
    sync_test(|| {
        // Get notepad
        let notepad_hwnd = unsafe {
            let notepad = "notepad\0".encode_utf16().collect::<Vec<_>>();
            let pw = PCWSTR::from_raw(notepad.as_ptr());
            FindWindowW(pw, PCWSTR::null())
        };

        assert_ne!(notepad_hwnd.0, 0, "Notepad must be running for this test");

        let err = get_desktop(99999).set_name("").unwrap_err();
        assert_eq!(err, Error::DesktopNotFound);

        let err = switch_desktop(99999).unwrap_err();
        assert_eq!(err, Error::DesktopNotFound);

        let err = get_desktop_by_window(HWND(9999999)).unwrap_err();
        assert_eq!(err, Error::WindowNotFound);

        let err = move_window_to_desktop(99999, &notepad_hwnd).unwrap_err();
        assert_eq!(err, Error::DesktopNotFound);

        let err = move_window_to_desktop(0, &HWND(999999)).unwrap_err();
        assert_eq!(err, Error::WindowNotFound);
    });
}

#[test]
fn test_threads() {
    sync_test(|| {
        // let get_count = || {
        //     get_desktop_count().unwrap();
        // };
        let mut threads = vec![];
        for _ in 0..555 {
            threads.push(std::thread::spawn(|| -> bool {
                for d in get_desktops().unwrap() {
                    let _n = match d.get_name() {
                        Ok(n) => n,
                        Err(Error::ComNotImplemented) => return true,
                        Err(e) => panic!("Failed to get name of desktop: {e:?}"),
                    };
                    let _i = d.get_index().unwrap();
                    // println!("Thread {n} {i} {:?}", std::thread::current().id());
                }
                false
            }));
        }
        thread::sleep(Duration::from_millis(150));
        for t in threads {
            if t.join().unwrap() {
                panic!("Failed to get name of desktop because that feature wasn't implemented on current Windows version");
            }
        }
    })
}

#[test]
fn test_listener_manual() {
    // This test can be run only individually
    let args = std::env::args().collect::<Vec<_>>();
    if !args.contains(&"tests::test_listener_manual".to_owned()) {
        return;
    }
    sync_test(|| {
        let (tx, rx) = std::sync::mpsc::channel::<DesktopEvent>();
        let _notifications_thread = listen_desktop_events(tx);

        std::thread::spawn(|| {
            for item in rx {
                println!("{:?}", item);
            }
        });

        // Wait for keypress
        println!("⛔ Press enter to stop");
        let mut input = String::new();
        std::io::stdin().read_line(&mut input).unwrap();
    })
}

#[test]
fn test_kill_explorer_exe_manually() {
    // This test can be run only individually
    let args = std::env::args().collect::<Vec<_>>();
    if !args.contains(&"tests::test_kill_explorer_exe_manually".to_owned()) {
        return;
    }
    sync_test(|| {
        // Kill explorer.exe
        let mut cmd = std::process::Command::new("taskkill");
        cmd.arg("/F").arg("/IM").arg("explorer.exe");
        cmd.output().unwrap();
        std::thread::sleep(Duration::from_secs(2));

        // Create listener
        let (tx, rx) = std::sync::mpsc::channel::<DesktopEvent>();
        let _notifications_thread = listen_desktop_events(tx);
        let _receiver = std::thread::spawn(|| {
            let mut count = 0;
            for item in rx {
                println!("{:?}", item);
                if let DesktopEvent::DesktopChanged { new: _, old: _ } = item {
                    count += 1;
                }
            }
            count
        });

        // Try switching desktops, can't work
        let error = switch_desktop(0);
        assert_eq!(error, Err(Error::ClassNotRegistered));

        // Start explorer exe
        let mut cmd = std::process::Command::new("explorer.exe");
        cmd.spawn().unwrap();
        std::thread::sleep(Duration::from_secs(4));

        // Try switching desktops, should work now
        switch_desktop(0).unwrap();
        std::thread::sleep(Duration::from_secs(1));
        switch_desktop(1).unwrap();
        std::thread::sleep(Duration::from_secs(1));
        switch_desktop(2).unwrap();
        std::thread::sleep(Duration::from_secs(1));
        drop(_notifications_thread);
        let count = _receiver.join().unwrap();
        assert_eq!(count, 3);
    })
}

#[test]
fn test_switch_desktops_rapidly_manual() {
    // This test can be run only individually
    let args = std::env::args().collect::<Vec<_>>();
    if !args.contains(&"tests::test_switch_desktops_rapidly_manual".to_owned()) {
        return;
    }
    sync_test(|| {
        let (tx, rx) = std::sync::mpsc::channel::<DesktopEvent>();
        // let (tx, rx) = crossbeam_channel::unbounded::<DesktopEvent>();

        let mut _notifications_thread = listen_desktop_events(tx).unwrap();
        let receiver = std::thread::spawn(move || {
            let mut count = 0;
            for item in rx {
                if let DesktopEvent::DesktopChanged {
                    new: _new,
                    old: _old,
                } = item
                {
                    count += 1;
                }
                if count == 7998 {
                    // This stops the receiver loop as the sender thread ends
                    _notifications_thread.stop().unwrap();
                }
            }

            count
        });

        let current_desktop = get_current_desktop().unwrap();

        for _ in 0..3999 {
            switch_desktop(0).unwrap();
            // std::thread::sleep(Duration::from_millis(4));
            switch_desktop(1).unwrap();
        }

        // Finally return to same desktop we were
        std::thread::sleep(Duration::from_millis(130));
        switch_desktop(current_desktop).unwrap();

        let count = receiver.join().unwrap();
        assert!(count >= 1999);
        println!("End of program, starting to drop stuff...");
    })
}

#[test]
fn test_desktop_count() {
    sync_test(|| {
        let count = get_desktop_count().unwrap();
        assert!(count > 1);
    })
}
