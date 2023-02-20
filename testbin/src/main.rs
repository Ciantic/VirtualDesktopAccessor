use winit::{
    event::{Event, WindowEvent},
    event_loop::{ControlFlow, EventLoopBuilder},
    window::WindowBuilder,
};
use winvd::*;

#[derive(Clone, Debug)]
enum MyCustomEvents {
    #[allow(dead_code)]
    MyEvent1,

    DesktopEvent(DesktopEvent),
}

// From DesktopEvent
impl From<DesktopEvent> for MyCustomEvents {
    fn from(e: DesktopEvent) -> Self {
        MyCustomEvents::DesktopEvent(e)
    }
}

fn main() {
    let event_loop = EventLoopBuilder::<MyCustomEvents>::with_user_event().build();
    let your_app_window = WindowBuilder::new().build(&event_loop).unwrap();

    let proxy = event_loop.create_proxy();
    let _thread = create_desktop_event_thread(proxy);

    event_loop.run(move |event, _, control_flow| {
        *control_flow = ControlFlow::Wait;

        match event {
            // Main window events
            Event::WindowEvent {
                event: WindowEvent::CloseRequested,
                window_id,
            } if window_id == your_app_window.id() => *control_flow = ControlFlow::Exit,

            // User events
            Event::UserEvent(e) => match e {
                MyCustomEvents::MyEvent1 => {
                    println!("MyEvent1");
                }
                MyCustomEvents::DesktopEvent(e) => {
                    println!("DesktopEvent: {:?}", e);
                }
            },
            _ => (),
        }
    });

    /*
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
    */
}
