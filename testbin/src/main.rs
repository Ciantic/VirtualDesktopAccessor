use winapi::um::winuser::FindWindowW;

use std::{ptr::null, time::Duration};
use winapi::shared::windef::HWND;
use winvirtualdesktops::VirtualDesktopService;

fn main() {
    let service = VirtualDesktopService::initialize().unwrap();

    // Test desktop retrieval methods ----------------------------------------
    let desktops = service.get_desktops().unwrap();
    println!("All desktops {:?}", desktops);

    let desktop_count = service.get_desktop_count();
    println!("Desktop count {:?}", desktop_count);

    let current_desktop_id = service.get_current_desktop().unwrap();
    println!("Current desktop ID {:?}", current_desktop_id);

    // Test window manipulation methods ----------------------------------------

    // Get notepad
    let notepad_hwnd: HWND = unsafe {
        FindWindowW(
            "notepad\0".encode_utf16().collect::<Vec<_>>().as_ptr(),
            null(),
        )
    };
    if (notepad_hwnd as u32) == 0 {
        println!("You must start notepad to continue tests.");
        return;
    }

    let notepad_desktop = service.get_desktop_by_window(notepad_hwnd);
    println!("Desktop of notepad: {:?}", notepad_desktop);

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

    // Switch to desktop and back
    println!("Move notepad to first desktop for three seconds, and then return it...");
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

    println!("Press enter key to close...");
    std::io::stdin().read_line(&mut String::new()).unwrap();

    /*
    // init_apartment(ApartmentType::Multithreaded).unwrap();
    let notepad_hwnd: HWND = unsafe {
        FindWindowW(
            "notepad\0".encode_utf16().collect::<Vec<_>>().as_ptr(),
            null(),
        )
    };

    let mut service_provider =
        create_instance::<dyn IServiceProvider>(&CLSID_ImmersiveShell).unwrap();
    let virtual_desktop_manager =
        get_immersive_service::<dyn IVirtualDesktopManager>(&service_provider).unwrap();
    let virtualdesktop_notification_service =
        get_immersive_service_for_class::<dyn IVirtualDesktopNotificationService>(
            &service_provider,
            CLSID_IVirtualNotificationService,
        )
        .unwrap();

    println!("IServiceProvider: {:?}", &service_provider as *const _);
    println!(
        "IVirtualDesktopManager: {:?}",
        &virtual_desktop_manager as *const _
    );

    if (notepad_hwnd as u32) == 0 {
        println!("You must start notepad to run this.");
        return;
    }

    println!("notepad {:?}", notepad_hwnd);
    let desktop_id: GUID = unsafe {
        let mut desktop_id_mut = empty_guid();

        virtual_desktop_manager.get_window_desktop_id(notepad_hwnd, &mut desktop_id_mut as *mut _);
        desktop_id_mut
    };
    println!("Desktop ID for Notepad {:?}", desktop_id);

    let ptr = create_change_listener().unwrap();

    let cookie = {
        let mut cookiee: DWORD = 0;
        let res: i32 = unsafe { virtualdesktop_notification_service.register(ptr, &mut cookiee) };
        if FAILED(res) {
            println!("Failure to register {:?} {:?}", res as u32, cookiee);
        } else {
            println!("Registered listener {:?}", cookiee);
        }
        cookiee
    };

    let mut stdin = io::stdin();
    // let mut stdout = io::stdout();
    println!("Press enter key to continue...");
    // write!(stdout, "Press any key to continue...").unwrap();
    // stdout.flush().unwrap();
    // Read a single byte and discard
    stdin.read_line(&mut String::new()).unwrap();
    // let _ = stdin.read(&mut [0u8]).unwrap();
    unsafe {
        virtualdesktop_notification_service.unregister(cookie);
    }
    // deinit_apartment();
    */
}
