use winapi::um::winuser::FindWindowW;
use winvirtualdesktops::initialize;

// fn empty_guid() -> GUID {
//     GUID {
//         data1: 0,
//         data2: 0,
//         data3: 0,
//         data4: [0, 0, 0, 0, 0, 0, 0, 0],
//     }
// }

fn main() {
    let service = initialize().unwrap();
    println!("Press enter key to continue...");
    service.get_current_desktop();
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
