use crate::listener::DesktopEventThread;
use crate::Desktop;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::convert::{TryFrom, TryInto};
use std::pin::Pin;
use std::sync::{Arc, Condvar, Mutex};
use std::{cell::RefCell, rc::Rc};
use windows::Win32::Foundation::HWND;

#[derive(Debug, Clone)]
pub enum DesktopEventSender {
    Std(std::sync::mpsc::Sender<DesktopEvent>),

    // #[cfg(feature = "crossbeam-channel")]
    Crossbeam(crossbeam_channel::Sender<DesktopEvent>),
}

impl DesktopEventSender {
    pub fn try_send(&self, event: DesktopEvent) {
        match self {
            DesktopEventSender::Std(sender) => {
                let _ = sender.send(event);
            }

            // #[cfg(feature = "crossbeam-channel")]
            DesktopEventSender::Crossbeam(sender) => {
                let _ = sender.try_send(event);
            }
        }
    }
}

#[derive(Debug, Clone)]
pub enum DesktopEvent {
    DesktopCreated(Desktop),
    DesktopDestroyed(Desktop),
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

static LISTENER: Lazy<Arc<Mutex<Option<DesktopEventThread>>>> =
    Lazy::new(|| Arc::new(Mutex::new(None)));

static MAIN_SENDER: Lazy<Arc<Mutex<Option<DesktopEventSender>>>> =
    Lazy::new(|| Arc::new(Mutex::new(None)));

pub fn add_event_sender(sender: &DesktopEventSender) {
    let mut main_sender = MAIN_SENDER.lock().unwrap();
    if main_sender.is_none() {
        *main_sender = Some(sender.clone());

        let mut listener = LISTENER.lock().unwrap();
        *listener = Some(DesktopEventThread::new(sender.clone()));
    }
}

pub fn remove_event_sender() {
    let mut listener = LISTENER.lock().unwrap();
    listener.take();
}
