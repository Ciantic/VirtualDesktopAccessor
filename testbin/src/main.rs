use std::{thread, time::Duration};
use winvd::*;

fn main() {
    // Desktop count
    let desktops = get_desktop_count().unwrap();
    println!("Desktops {:?}", desktops);

    // Desktops are:
    println!("Desktops are: {:?}", get_desktops().unwrap());

    // !! TODO: START A LISTENER

    thread::spawn(|| {
        thread::sleep(Duration::from_secs(1));

        // Create and remove a desktop
        let desk = create_desktop().unwrap();
        println!("Create desktop {:?}", desk);

        // Set and get the name of the new desktop
        get_desktop(desk)
            .set_name("This is a new desktop!")
            .unwrap();
        println!(
            "New desktop with name: {}",
            get_desktop(desk).get_name().unwrap()
        );

        remove_desktop(desk, get_current_desktop().unwrap()).unwrap();
        println!("Deleted desktop {:?}", desk);
    });

    println!("Press enter key to close...");
    std::io::stdin().read_line(&mut String::new()).unwrap();
}
