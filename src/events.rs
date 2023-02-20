use crate::Desktop;
use crate::DesktopEventThread;
use windows::Win32::Foundation::HWND;

#[derive(Clone)]
pub enum DesktopEventSender<T>
where
    T: 'static,
{
    Std(std::sync::mpsc::Sender<T>),

    #[cfg(feature = "crossbeam-channel")]
    Crossbeam(crossbeam_channel::Sender<T>),

    #[cfg(feature = "winit")]
    Winit(winit::event_loop::EventLoopProxy<T>),
}

// From STD Sender
impl<T> From<std::sync::mpsc::Sender<T>> for DesktopEventSender<T>
where
    T: From<DesktopEvent> + Clone + Send + 'static,
{
    fn from(sender: std::sync::mpsc::Sender<T>) -> Self {
        DesktopEventSender::Std(sender)
    }
}

// From Crossbeam Sender
#[cfg(feature = "crossbeam-channel")]
impl<T> From<crossbeam_channel::Sender<T>> for DesktopEventSender<T>
where
    T: From<DesktopEvent> + Clone + Send + 'static,
{
    fn from(sender: crossbeam_channel::Sender<T>) -> Self {
        DesktopEventSender::Crossbeam(sender)
    }
}

// From Winit Sender
#[cfg(feature = "winit")]
impl<T> From<winit::event_loop::EventLoopProxy<T>> for DesktopEventSender<T>
where
    T: From<DesktopEvent> + Clone + Send + 'static,
{
    fn from(sender: winit::event_loop::EventLoopProxy<T>) -> Self {
        DesktopEventSender::Winit(sender)
    }
}

impl<T> DesktopEventSender<T> {
    pub fn try_send(&self, event: T) {
        match self {
            DesktopEventSender::Std(sender) => {
                let _ = sender.send(event);
            }

            #[cfg(feature = "crossbeam-channel")]
            DesktopEventSender::Crossbeam(sender) => {
                let _ = sender.try_send(event);
            }

            #[cfg(feature = "winit")]
            DesktopEventSender::Winit(sender) => {
                let _ = sender.send_event(event);
            }
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum DesktopEvent {
    DesktopCreated(Desktop),
    DesktopDestroyed {
        destroyed: Desktop,
        fallback: Desktop,
    },
    DesktopChanged {
        new: Desktop,
        old: Desktop,
    },
    DesktopNameChanged(Desktop, String),
    DesktopWallpaperChanged(Desktop, String),
    DesktopMoved {
        desktop: Desktop,
        old_index: i64,
        new_index: i64,
    },
    WindowChanged(HWND),
}

/// Create event sending thread, give this `crossbeam_channel::Sender<T>`, `winit::event_loop::EventLoopProxy<T>`, or `std::sync::mpsc::Sender<T>`.
///
/// Your message type `T` needs to be convertible to `DesktopEvent`.
///
/// This function returns `DesktopEventThread`, you must keep the value alive,
/// when the value is dropped the listener is closed and thread joined.
///
/// # Example
///
/// ```rust
/// let (tx, rx) = std::sync::mpsc::channel::<DesktopEvent>();
/// let _notifications_thread = create_desktop_event_thread(tx);
/// // Do with receiver something
/// for item in rx {
///    println!("{:?}", item);
/// }
/// // When `_notifications_thread` is dropped the thread is joined and listener closed.
/// ```
///
/// Additionally you can pass crossbeam-channel sender, or winit eventloop proxy
/// to the function.
///
pub fn create_desktop_event_thread<T, S>(sender: S) -> DesktopEventThread
where
    T: From<DesktopEvent> + Clone + Send + 'static,
    S: Into<DesktopEventSender<T>> + Clone,
{
    DesktopEventThread::new(sender.into())
}
