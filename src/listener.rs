use std::convert::TryInto;
use std::pin::Pin;
use std::time::Duration;

use windows::core::HSTRING;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::Threading::{
    GetCurrentThread, GetCurrentThreadId, SetThreadPriority, THREAD_PRIORITY_TIME_CRITICAL,
};
use windows::Win32::UI::Shell::Common::IObjectArray;
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, GetMessageW, PostQuitMessage, PostThreadMessageW, SetTimer, TranslateMessage,
    MSG, WM_TIMER, WM_USER,
};

use crate::comobjects::ComObjects;
use crate::hresult::HRESULT;
use crate::interfaces::{
    ComIn, IApplicationView, IVirtualDesktop, IVirtualDesktopNotification,
    IVirtualDesktopNotification_Impl,
};
use crate::log::log_output;
use crate::{DesktopEvent, Result};
use crate::{DesktopEventSender, Error};

// Log format macro
macro_rules! log_format {
    ($($arg:tt)*) => {
        #[cfg(debug_assertions)]
        $crate::log::log_output(&format!($($arg)*));
    };
}

const WM_USER_QUIT: u32 = WM_USER + 0x10;

/// Event listener thread, create with `create_desktop_event_thread(sender)`, value must be held in the state of the program, the thread is joined when the value is dropped.
#[derive(Debug)]
pub struct DesktopEventThread {
    windows_thread_id: Option<u32>,
    thread: Option<std::thread::JoinHandle<()>>,
}

impl DesktopEventThread {
    pub(crate) fn new<T>(sender: DesktopEventSender<T>) -> Result<Self>
    where
        T: From<DesktopEvent> + Clone + Send + 'static,
    {
        // Channel for thread id
        let (tx, rx) = std::sync::mpsc::channel();

        // Main notification thread, with STA message loop
        let notification_thread = std::thread::spawn(move || {
            log_format!("Listener thread started {:?}", std::thread::current().id());

            let win_thread_id = unsafe { GetCurrentThreadId() };

            // Send the Windows specific thread id to the main thread
            let res = tx.send(win_thread_id);
            drop(tx);
            if let Err(er) = res {
                log_format!("Could not send thread id to main thread {:?}", er);
                return;
            }

            // Set thread priority to time critical, explorer.exe really hates if your listener thread is slow
            unsafe { SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL) };

            // Set a timer to check if the listener is still alive
            unsafe {
                SetTimer(HWND::default(), 0, 3000, None);
            }

            let com_objects = ComObjects::new();
            loop {
                log_output("Try to create listener service...");
                let mut quit = false;
                let sender_new = sender.clone();
                let listener = VirtualDesktopNotificationWrapper::new(
                    &com_objects,
                    Box::new(move |event| {
                        sender_new.try_send(event.into());
                    }),
                );

                // Retry if the listener could not be created after every three seconds
                if let Err(er) = listener {
                    log_format!(
                        "Listener service could not be created, retrying in three seconds {:?}",
                        er
                    );
                    std::thread::sleep(Duration::from_secs(3));
                    continue;
                }

                // STA message loop
                let mut msg = MSG::default();
                unsafe {
                    loop {
                        let continuation = GetMessageW(&mut msg, HWND::default(), 0, 0);
                        if (continuation.0 == 0) || (continuation.0 == -1) {
                            break;
                        }

                        if msg.message == WM_USER_QUIT {
                            quit = true;
                            PostQuitMessage(0);
                        } else if msg.message == WM_TIMER {
                            // If com objects aren't connected anymore, drop them and recreate
                            if !com_objects.is_connected() {
                                log_output("Listener is not connected, restarting...");
                                com_objects.drop_services();
                                TranslateMessage(&msg);
                                DispatchMessageW(&msg);

                                // Break out of the while message loop, and restart the listener
                                break;
                            }
                        }
                        TranslateMessage(&msg);
                        DispatchMessageW(&msg);
                    }

                    if quit {
                        // Break out of the loop, and drop the listener
                        break;
                    }
                }
            }

            log_format!("Listener thread finished {:?}", std::thread::current().id());
        });

        // Wait until the thread has started, and sent its Windows specific thread id
        let win_thread_id = rx
            .recv_timeout(Duration::from_secs(1))
            .map_err(|_| Error::ListenerThreadIdNotCreated)?;
        drop(rx);

        // Store the new thread
        Ok(DesktopEventThread {
            windows_thread_id: Some(win_thread_id),
            thread: Some(notification_thread),
        })
    }

    /// Stops the listener, and join the thread if it is still running, normally
    /// you don't need to call this as drop calls this automatically
    pub fn stop(&mut self) -> std::thread::Result<()> {
        if let Some(thread_id) = self.windows_thread_id.take() {
            log_output("Stopping listener thread");
            unsafe {
                PostThreadMessageW(
                    thread_id,
                    WM_USER_QUIT,
                    WPARAM::default(),
                    LPARAM::default(),
                );
            }
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

    com_objects: &'a ComObjects,
    cookie: u32,
}

impl<'a> VirtualDesktopNotificationWrapper<'a> {
    pub fn new(
        com_objects: &'a ComObjects,
        sender: Box<dyn Fn(DesktopEvent)>,
    ) -> Result<Pin<Box<VirtualDesktopNotificationWrapper>>> {
        let ptr = Pin::new(Box::new(VirtualDesktopNotification { sender }.into()));

        let notification = Pin::new(Box::new(VirtualDesktopNotificationWrapper {
            cookie: com_objects.register_for_notifications(&ptr)?,
            ptr,
            com_objects,
        }));

        log_format!(
            "Registered notification {} {:?}",
            notification.cookie,
            std::thread::current().id()
        );

        Ok(notification)
    }
}

impl Drop for VirtualDesktopNotificationWrapper<'_> {
    fn drop(&mut self) {
        log_format!(
            "Unregistering notification {} {:?}",
            self.cookie,
            std::thread::current().id()
        );

        let _ = self.com_objects.unregister_for_notifications(self.cookie);
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

    unsafe fn view_virtual_desktop_changed(&self, view: IApplicationView) -> HRESULT {
        let mut hwnd = HWND::default();
        view.get_thumbnail_window(&mut hwnd);
        (self.sender)(DesktopEvent::WindowChanged(hwnd));
        HRESULT(0)
    }
}
