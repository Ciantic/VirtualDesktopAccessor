use winapi::um::winuser::FindWindowW;

use std::{ptr::null, time::Duration};
use winvd::{VirtualDesktopService, HWND};

fn main() {
    let service = VirtualDesktopService::create_with_com().unwrap();

    service.on_desktop_change(Box::new(|old, new| {
        println!("Desktop changed from {:?} to {:?}", old, new);
    }));

    service.on_window_change(Box::new(|hwnd| {
        println!("Window changed {:?} ", hwnd);
    }));

    service.on_desktop_created(Box::new(|desktop| {
        println!("Created desktop {:?} ", desktop);
    }));

    service.on_desktop_destroyed(Box::new(|desktop| {
        println!("Desktop destroyed {:?} ", desktop);
    }));

    // Test desktop retrieval methods ----------------------------------------
    let desktops = service.get_desktops().unwrap();
    println!("All desktops {:?}", desktops);

    let desktop_count = service.get_desktop_count();
    println!("Desktop count {:?}", desktop_count);

    let current_desktop_id = service.get_current_desktop().unwrap();
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

    let notepad_desktop = service.get_desktop_by_window(notepad_hwnd);
    println!(
        "Desktop of notepad: {:?}, hwnd: {:?}",
        notepad_desktop, notepad_hwnd
    );

    // Is on current desktop
    let notepad_is_on_current_desktop = service.is_window_on_current_virtual_desktop(notepad_hwnd);
    println!(
        "Notepad is on current desktop: {:?}",
        notepad_is_on_current_desktop
    );

    // Is on specific desktop
    let notepad_is_on_specific_desktop =
        service.is_window_on_desktop(notepad_hwnd, &current_desktop_id);
    println!(
        "Is notepad on desktop: {:?}, true or false: {:?}",
        current_desktop_id.clone(),
        notepad_is_on_specific_desktop
    );

    // Move window between desktops

    // Not a real window, testing error
    println!(
        "Try to move non existant window... {:?}",
        service.move_window_to_desktop(999999999 as HWND, desktops.get(0).unwrap())
    );

    // Move notepad
    println!("Move notepad to first desktop for three seconds, and then return it...");
    println!(
        "Move to first... {:?}",
        service.move_window_to_desktop(notepad_hwnd, desktops.get(0).unwrap())
    );
    println!("Wait three seconds...");
    std::thread::sleep(Duration::from_secs(3));
    println!(
        "Move back to this desktop {:?}",
        service.move_window_to_desktop(notepad_hwnd, &current_desktop_id)
    );

    println!(
        "Pin the notepad window {:?}",
        service.pin_window(notepad_hwnd)
    );

    // Switch to desktop and back
    println!("Switch between desktops 1 and this one...");

    // Wait a bit
    std::thread::sleep(Duration::from_secs(1));

    // Do it!
    println!(
        "Move to first... {:?}",
        service.go_to_desktop(desktops.get(0).unwrap())
    );
    println!("Wait three seconds...");
    std::thread::sleep(Duration::from_secs(3));
    println!(
        "Move back to this desktop {:?}",
        service.go_to_desktop(&current_desktop_id)
    );

    println!(
        "Unpin the notepad window {:?}",
        service.unpin_window(notepad_hwnd)
    );

    println!("Press enter key to close...");
    std::io::stdin().read_line(&mut String::new()).unwrap();
}
