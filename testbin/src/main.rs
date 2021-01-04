use std::{thread, time::Duration};
use winvd::{
    create_desktop, get_current_desktop, get_desktops, get_event_receiver,
    helpers::get_desktop_count, remove_desktop, VirtualDesktopEvent,
};

fn main() {
    // Desktop count
    let desktops = get_desktop_count().unwrap();
    println!("Desktops {:?}", desktops);

    // Desktops are:
    println!("Desktops are: {:?}", get_desktops().unwrap());

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

        // Create and remove a desktop
        let desk = create_desktop().unwrap();
        println!("Create desktop {:?}", desk);

        remove_desktop(&desk, &get_current_desktop().unwrap()).unwrap();
        println!("Deleted desktop {:?}", desk);
    });

    println!("Press enter key to close...");
    std::io::stdin().read_line(&mut String::new()).unwrap();
}
