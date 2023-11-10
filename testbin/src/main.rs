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
    let event_loop = EventLoopBuilder::<MyCustomEvents>::with_user_event()
        .build()
        .unwrap();
    let your_app_window = WindowBuilder::new().build(&event_loop).unwrap();

    let proxy = event_loop.create_proxy();
    let mut _thread = listen_desktop_events(proxy).unwrap();

    event_loop.set_control_flow(ControlFlow::Wait);

    event_loop
        .run(move |event, elewt| {
            match event {
                // Main window events
                Event::WindowEvent {
                    event: WindowEvent::CloseRequested,
                    window_id,
                } if window_id == your_app_window.id() => {
                    let _ = _thread.stop();

                    elewt.exit();
                }

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
        })
        .unwrap();
}
