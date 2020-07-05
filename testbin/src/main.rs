use winapi::um::winuser::FindWindowW;

use std::{ptr::null, thread, time::Duration};
use winvd::{
    get_current_desktop, get_desktop_by_window, get_desktop_count, get_desktops,
    get_event_receiver, go_to_desktop, is_window_on_current_virtual_desktop, is_window_on_desktop,
    move_window_to_desktop, pin_window, unpin_window, Error, VirtualDesktopEvent, HWND,
};

fn main() {
    thread::spawn(|| {
        get_event_receiver().iter().for_each(|msg| match msg {
            VirtualDesktopEvent::DesktopChanged(old, new) => {
                println!(
                    "<- Desktop changed from {:?} to {:?} {:?}",
                    old,
                    new,
                    thread::current().id()
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
        })
    });

    thread::spawn(|| {
        thread::sleep(Duration::from_secs(2));
        let desktop_count = get_desktop_count();
        println!("----------------------------------------------------");
        println!(
            "In separate thread {:?}, call desktop count {:?}",
            thread::current().id(),
            desktop_count
        );
    });

    // Test desktop retrieval methods ----------------------------------------
    let desktops = get_desktops().unwrap();
    println!("All desktops {:?}", desktops);

    let desktop_count = get_desktop_count();
    println!("Desktop count {:?}", desktop_count);

    let current_desktop_id = get_current_desktop().unwrap();
    println!("Current desktop ID {:?}", current_desktop_id);

    // Test window manipulation methods ----------------------------------------
    println!("Start notepad, and press enter key to continue...");
    std::io::stdin().read_line(&mut String::new()).unwrap();

    // Get notepad
    let notepad_hwnd: HWND = unsafe {
        FindWindowW(
            "notepad\0".encode_utf16().collect::<Vec<_>>().as_ptr(),
            null(),
        ) as HWND
    };
    if (notepad_hwnd) == 0 {
        println!("You must start notepad to continue tests.");
        return;
    }

    let notepad_desktop = get_desktop_by_window(notepad_hwnd);
    println!(
        "Desktop of notepad: {:?}, hwnd: {:?}",
        notepad_desktop, notepad_hwnd
    );

    // Is on current desktop
    let notepad_is_on_current_desktop = is_window_on_current_virtual_desktop(notepad_hwnd);
    println!(
        "Notepad is on current desktop: {:?}",
        notepad_is_on_current_desktop
    );

    // Is on specific desktop
    let notepad_is_on_specific_desktop = is_window_on_desktop(notepad_hwnd, &current_desktop_id);
    println!(
        "Is notepad on desktop: {:?}, true or false: {:?}",
        current_desktop_id.clone(),
        notepad_is_on_specific_desktop
    );

    // Move window between desktops

    // Not a real window, testing error
    println!("Try to move non existant window...",);
    debug_assert!(
        move_window_to_desktop(999999999 as HWND, desktops.get(0).unwrap())
            == Err(Error::WindowNotFound)
    );

    // Move notepad
    println!("Move notepad to first desktop for three seconds, and then return it...");
    println!(
        "Move to first... {:?}",
        move_window_to_desktop(notepad_hwnd, desktops.get(0).unwrap())
    );
    println!("Wait three seconds...");
    std::thread::sleep(Duration::from_secs(3));
    println!(
        "Move back to this desktop {:?}",
        move_window_to_desktop(notepad_hwnd, &current_desktop_id)
    );

    println!("Pin the notepad window {:?}", pin_window(notepad_hwnd));

    // Test desktop manipulation methods ----------------------------------------

    // Switch to desktop and back
    println!("Switch between desktops 1 and this one...");

    // Wait a bit
    std::thread::sleep(Duration::from_secs(1));

    // Do it!
    println!(
        "Move to first... {:?}",
        go_to_desktop(desktops.get(0).unwrap())
    );
    println!("Wait three seconds...");
    std::thread::sleep(Duration::from_secs(3));
    println!(
        "Move back to this desktop {:?}",
        go_to_desktop(&current_desktop_id)
    );

    println!("Unpin the notepad window {:?}", unpin_window(notepad_hwnd));

    println!("Press enter key to close...");
    std::io::stdin().read_line(&mut String::new()).unwrap();
}
