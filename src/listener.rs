use std::convert::TryInto;
use std::pin::Pin;
use std::time::Duration;

use windows::core::{Vtable, HRESULT, HSTRING};
use windows::Win32::Foundation::HWND;
use windows::Win32::System::Threading::{
    GetCurrentThread, SetThreadPriority, THREAD_PRIORITY_TIME_CRITICAL,
};
use windows::Win32::UI::Shell::Common::IObjectArray;

use crate::comobjects::ComObjects;
use crate::interfaces::{
    ComIn, IApplicationView, IVirtualDesktop, IVirtualDesktopNotification,
    IVirtualDesktopNotification_Impl,
};
use crate::log::log_output;
use crate::DesktopEventSender;
use crate::{DesktopEvent, Result};

// Log format macro
macro_rules! log_format {
    ($($arg:tt)*) => {
        #[cfg(debug_assertions)]
        $crate::log::log_output(&format!($($arg)*));
    };
}

enum DekstopEventThreadMsg {
    Quit,
}

/// Event listener thread, create with `listen_desktop_events(sender)`,
/// value must be held in the state of the program, the thread is joined when
/// the value is dropped.
#[derive(Debug)]
pub struct DesktopEventThread {
    thread_control_sender: Option<std::sync::mpsc::Sender<DekstopEventThreadMsg>>,
    thread: Option<std::thread::JoinHandle<()>>,
}

impl DesktopEventThread {
    pub(crate) fn new<T>(sender: DesktopEventSender<T>) -> Result<Self>
    where
        T: From<DesktopEvent> + Clone + Send + 'static,
    {
        // Channel for quitting
        let (tx, rx) = std::sync::mpsc::channel::<DekstopEventThreadMsg>();

        // Main notification thread, with STA message loop
        let notification_thread = std::thread::spawn(move || {
            let com_objects = ComObjects::new();
            log_format!("Listener thread started {:?}", std::thread::current().id());

            // Set thread priority to time critical, explorer.exe really hates if your listener thread is slow
            unsafe { SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL) };

            // Create listener
            let sender_new = sender.clone();
            let mut listener = VirtualDesktopNotificationWrapper::new(
                &com_objects,
                Box::new(move |event| {
                    sender_new.try_send(event.into());
                }),
            );

            loop {
                let item = rx.recv_timeout(Duration::from_secs(3));
                match item {
                    Ok(DekstopEventThreadMsg::Quit) => {
                        log_output("Listener thread received quit message");
                        break;
                    }
                    Err(_) => {
                        if !com_objects.is_connected() || listener.is_err() {
                            log_output(
                                "Listener is not connected, or failed to register, trying again",
                            );

                            // Drop will unregister the old listener before the
                            // new one is created, this is required, read more
                            // from note-IVirtualDesktopNotification.md
                            drop(listener);
                            let sender_new = sender.clone();
                            listener = VirtualDesktopNotificationWrapper::new(
                                &com_objects,
                                Box::new(move |event| {
                                    sender_new.try_send(event.into());
                                }),
                            );
                        }
                    }
                }
            }

            log_format!("Listener thread finished {:?}", std::thread::current().id());
        });

        // Store the new thread
        Ok(DesktopEventThread {
            thread_control_sender: Some(tx),
            thread: Some(notification_thread),
        })
    }

    /// Stops the listener, and join the thread if it is still running, normally
    /// you don't need to call this as drop calls this automatically
    pub fn stop(&mut self) -> std::thread::Result<()> {
        if let Some(thread_control_sender) = self.thread_control_sender.take() {
            let _ = thread_control_sender.send(DekstopEventThreadMsg::Quit);
        }

        if let Some(thread) = self.thread.take() {
            thread.join()?;
        }
        Ok(())
    }
}

impl Drop for DesktopEventThread {
    fn drop(&mut self) {
        let res = self.stop();

        #[cfg(debug_assertions)]
        if let Err(err) = res {
            log_format!("Could not stop listener thread {:?}", err);
        }
    }
}

/// Wrapper registers the actual IVirtualDesktopNotification and on drop unregisters the notification
struct VirtualDesktopNotificationWrapper<'a> {
    #[allow(dead_code)]
    ptr: Pin<Box<IVirtualDesktopNotification>>,
    cookie: u32,
    com_objects: &'a ComObjects,
}

impl<'a> VirtualDesktopNotificationWrapper<'a> {
    pub fn new(
        com_objects: &'a ComObjects,
        sender: Box<dyn Fn(DesktopEvent)>,
    ) -> Result<Pin<Box<VirtualDesktopNotificationWrapper>>> {
        let ptr: Pin<Box<IVirtualDesktopNotification>> =
            Pin::new(Box::new(VirtualDesktopNotification { sender }.into()));
        let raw_ptr = ptr.as_raw();
        let cookie = com_objects.register_for_notifications(raw_ptr)?;
        let notification = Pin::new(Box::new(VirtualDesktopNotificationWrapper {
            com_objects,
            cookie,
            ptr,
        }));
        log_format!(
            "Registered notification {:?} {} {:?}",
            raw_ptr,
            notification.cookie,
            std::thread::current().id()
        );

        Ok(notification)
    }
}

impl<'a> Drop for VirtualDesktopNotificationWrapper<'a> {
    fn drop(&mut self) {
        log_format!(
            "Unregistering notification {} {:?}",
            self.cookie,
            std::thread::current().id()
        );

        let cookie = self.cookie;
        let _ = self.com_objects.unregister_for_notifications(cookie);
    }
}

#[windows::core::implement(IVirtualDesktopNotification)]
struct VirtualDesktopNotification {
    sender: Box<dyn Fn(DesktopEvent)>,
}

fn eat_error<T>(func: impl FnOnce() -> Result<T>) -> Option<T> {
    let res = func();
    match res {
        Ok(v) => Some(v),
        Err(er) => {
            log_format!("Error in listener: {:?}", er);
            None
        }
    }
}

// Allow unused variable warnings
#[allow(unused_variables)]
impl IVirtualDesktopNotification_Impl for VirtualDesktopNotification {
    unsafe fn current_virtual_desktop_changed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_old: ComIn<IVirtualDesktop>,
        desktop_new: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        eat_error(|| {
            Ok((self.sender)(DesktopEvent::DesktopChanged {
                old: desktop_old.try_into()?,
                new: desktop_new.try_into()?,
            }))
        });
        HRESULT(0)
    }

    unsafe fn virtual_desktop_wallpaper_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        eat_error(|| {
            Ok((self.sender)(DesktopEvent::DesktopWallpaperChanged(
                desktop.try_into()?,
                name.to_string(),
            )))
        });
        HRESULT(0)
    }

    unsafe fn virtual_desktop_created(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        eat_error(|| {
            Ok((self.sender)(DesktopEvent::DesktopCreated(
                desktop.try_into()?,
            )))
        });
        HRESULT(0)
    }

    unsafe fn virtual_desktop_destroy_begin(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        HRESULT(0)
    }

    unsafe fn virtual_desktop_destroy_failed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        HRESULT(0)
    }

    unsafe fn virtual_desktop_destroyed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        // Desktop destroyed is not anymore in the stack
        eat_error(|| {
            Ok((self.sender)(DesktopEvent::DesktopDestroyed {
                destroyed: desktop_destroyed.try_into()?,
                fallback: desktop_fallback.try_into()?,
            }))
        });
        HRESULT(0)
    }

    unsafe fn virtual_desktop_is_per_monitor_changed(&self, is_per_monitor: i32) -> HRESULT {
        log_format!("Desktop is per monitor changed: {}", is_per_monitor != 0);

        HRESULT(0)
    }

    unsafe fn virtual_desktop_moved(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
        old_index: i64,
        new_index: i64,
    ) -> HRESULT {
        eat_error(|| {
            Ok((self.sender)(DesktopEvent::DesktopMoved {
                desktop: desktop.try_into()?,
                old_index,
                new_index,
            }))
        });
        HRESULT(0)
    }

    unsafe fn virtual_desktop_name_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        eat_error(|| {
            Ok((self.sender)(DesktopEvent::DesktopNameChanged(
                desktop.try_into()?,
                name.to_string(),
            )))
        });
        HRESULT(0)
    }

    unsafe fn view_virtual_desktop_changed(&self, view: ComIn<IApplicationView>) -> HRESULT {
        let mut hwnd = HWND::default();
        let _ = view.get_thumbnail_window(&mut hwnd);
        (self.sender)(DesktopEvent::WindowChanged(hwnd));
        HRESULT(0)
    }
}
