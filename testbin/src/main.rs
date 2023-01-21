use std::{thread, time::Duration};
use winvd::{
    create_desktop, get_current_desktop, get_desktops, helpers::get_desktop_count, remove_desktop,
    set_event_sender, VirtualDesktopEvent, VirtualDesktopEventSender,
};

fn main() {
    // Desktop count
    let desktops = get_desktop_count().unwrap();
    println!("Desktops {:?}", desktops);

    // Desktops are:
    println!("Desktops are: {:?}", get_desktops().unwrap());

    thread::spawn(|| {
        let (sender, receiver) = std::sync::mpsc::channel();
        set_event_sender(VirtualDesktopEventSender::Std(sender)).unwrap();

        receiver.iter().for_each(|msg| match msg {
            VirtualDesktopEvent::DesktopChanged(old, new) => {
                println!(
                    "<- Desktop changed from {:?} to {:?} {:?}",
                    old.get_index().unwrap_or(999),
                    new.get_index().unwrap_or(999),
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
            VirtualDesktopEvent::DesktopNameChanged(desk, name) => {
                println!("<- Name of {:?} changed to {}", desk, name);
            }
            VirtualDesktopEvent::DesktopWallpaperChanged(desk, name) => {
                println!("<- Wallpaper of {:?} changed to {}", desk, name);
            }
            VirtualDesktopEvent::DesktopMoved(desk, old, new) => {
                println!("<- Desktop {:?} moved from {} to {}", desk, old, new);
            }
        })
    });

    thread::spawn(|| {
        thread::sleep(Duration::from_secs(2));

        // Create and remove a desktop
        let desk = create_desktop().unwrap();
        println!("Create desktop {:?}", desk);

        // Set and get the name of the new desktop
        desk.set_name("This is a new desktop!").unwrap();
        println!("New desktop with name: {}", desk.get_name().unwrap());

        remove_desktop(&desk, &get_current_desktop().unwrap()).unwrap();
        println!("Deleted desktop {:?}", desk);
    });

    println!("Press enter key to close...");
    std::io::stdin().read_line(&mut String::new()).unwrap();
}
