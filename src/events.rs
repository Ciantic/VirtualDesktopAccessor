use crate::Desktop;
use crate::DesktopEventThread;
use once_cell::sync::Lazy;
use std::collections::HashMap;
use std::convert::{TryFrom, TryInto};
use std::pin::Pin;
use std::sync::{Arc, Condvar, Mutex};
use std::{cell::RefCell, rc::Rc};
use windows::Win32::Foundation::HWND;

#[derive(Debug, Clone)]
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

/// Create event sending thread, give this crossbeam, winit eventloop proxy or std mpsc sender
pub fn create_event_thread<T, S>(sender: S) -> DesktopEventThread
where
    T: From<DesktopEvent> + Clone + Send + 'static,
    S: Into<DesktopEventSender<T>> + Clone,
{
    DesktopEventThread::new(sender.into())
}
