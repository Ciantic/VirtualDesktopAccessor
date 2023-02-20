use std::collections::HashMap;
use std::convert::{TryFrom, TryInto};
use std::pin::Pin;
use std::sync::{Arc, Condvar, Mutex};
use std::{cell::RefCell, rc::Rc};

use once_cell::sync::Lazy;
use windows::core::{IUnknown, IUnknownImpl, GUID, HSTRING};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::Shell::Common::IObjectArray;
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, GetMessageW, PostQuitMessage, PostThreadMessageW, SetTimer, TranslateMessage,
    MSG, TIMERPROC, WM_TIMER, WM_USER,
};

use crate::hresult::HRESULT;
use crate::interfaces::{
    ComIn, IApplicationView, IVirtualDesktop, IVirtualDesktopNotification,
    IVirtualDesktopNotification_Impl,
};
use crate::{comobjects::*, log_format, log_output};
use crate::{Desktop, DesktopEventSender};
use crate::{DesktopEvent, Result};

const WM_USER_QUIT: u32 = WM_USER + 0x10;

/// Event listener thread, create with `create_event_thread(sender)`, value must be held in the state of the program, the thread is joined when the value is dropped.
pub struct DesktopEventThread {
    windows_thread_id: Option<u32>,
    thread: Option<std::thread::JoinHandle<()>>,
}

impl DesktopEventThread {
    pub(crate) fn new<T>(sender: DesktopEventSender<T>) -> Self
    where
        T: From<DesktopEvent> + Clone + Send + 'static,
    {
        // Channel for thread id
        let (tx, rx) = std::sync::mpsc::channel();

        // Main notification thread, with STA message loop
        let notification_thread = std::thread::spawn(move || {
            log_format!("Listener thread started {:?}", std::thread::current().id());

            // Send the Windows specific thread id to the main thread
            tx.send(unsafe { GetCurrentThreadId() }).unwrap();
            drop(tx);

            // Set a timer to check if the listener is still alive
            unsafe {
                SetTimer(HWND::default(), 0, 3000, None);
            }

            let com_objects = ComObjects::new();
            loop {
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
                }

                // STA message loop
                let mut msg = MSG::default();
                unsafe {
                    while GetMessageW(&mut msg, HWND::default(), 0, 0).as_bool() {
                        if msg.message == WM_USER_QUIT {
                            quit = true;
                            PostQuitMessage(0);
                        }

                        if msg.message == WM_TIMER {
                            // If com objects aren't connected anymore, drop them
                            if !com_objects.is_connected() {
                                log_output("Not alive, restarting");
                                com_objects.drop_services();

                                // TODO: THIS IS DEEPLY WRONG
                                PostQuitMessage(0);
                            } else {
                                log_output("Is alive");
                            }
                        }
                        TranslateMessage(&msg);
                        DispatchMessageW(&msg);
                    }
                    if quit {
                        break;
                    }
                }
            }

            log_format!("Listener thread finished {:?}", std::thread::current().id());
        });

        // Wait until the thread has started, and sent its Windows specific thread id
        let win_thread_id = rx.recv().unwrap();
        drop(rx);

        // Store the new thread
        DesktopEventThread {
            windows_thread_id: Some(win_thread_id),
            thread: Some(notification_thread),
        }
    }

    fn drop_thread(&mut self) -> std::thread::Result<()> {
        if let Some(thread_id) = self.windows_thread_id.take() {
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
        log_output("Stopping listener thread");
        self.drop_thread().unwrap();
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

        #[cfg(debug_assertions)]
        log_output(&format!(
            "Registered notification {} {:?}",
            notification.cookie,
            std::thread::current().id()
        ));

        Ok(notification)
    }
}

impl Drop for VirtualDesktopNotificationWrapper<'_> {
    fn drop(&mut self) {
        #[cfg(debug_assertions)]
        log_output(&format!(
            "Unregistering notification {} {:?}",
            self.cookie,
            std::thread::current().id()
        ));
        let _ = self.com_objects.unregister_for_notifications(self.cookie);
    }
}

#[windows::core::implement(IVirtualDesktopNotification)]
struct VirtualDesktopNotification {
    sender: Box<dyn Fn(DesktopEvent)>,
}

#[cfg(debug_assertions)]
fn debug_desktop(desktop_new: &IVirtualDesktop, prefix: &str) {
    let mut gid = GUID::default();
    unsafe { desktop_new.get_id(&mut gid).panic_if_failed() };

    let name = "";

    // let mut name = HSTRING::new();
    // unsafe { desktop_new.get_name(&mut name).panic_if_failed() };

    #[cfg(debug_assertions)]
    log_output(&format!(
        "{}: {:?} {:?} {:?}",
        prefix,
        gid,
        name.to_string(),
        std::thread::current().id()
    ));
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
        (self.sender)(DesktopEvent::DesktopChanged {
            old: desktop_old.try_into().unwrap(),
            new: desktop_new.try_into().unwrap(),
        });
        HRESULT(0)
    }

    unsafe fn virtual_desktop_wallpaper_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        (self.sender)(DesktopEvent::DesktopWallpaperChanged(
            desktop.try_into().unwrap(),
            name.to_string(),
        ));
        HRESULT(0)
    }

    unsafe fn virtual_desktop_created(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        (self.sender)(DesktopEvent::DesktopCreated(desktop.try_into().unwrap()));
        HRESULT(0)
    }

    unsafe fn virtual_desktop_destroy_begin(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        // Desktop destroyed is not anymore in the stack
        #[cfg(debug_assertions)]
        debug_desktop(&desktop_destroyed, "Desktop destroy begin");
        #[cfg(debug_assertions)]
        debug_desktop(&desktop_fallback, "Desktop destroy fallback");
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
        (self.sender)(DesktopEvent::DesktopDestroyed {
            destroyed: desktop_destroyed.try_into().unwrap(),
            fallback: desktop_fallback.try_into().unwrap(),
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
        (self.sender)(DesktopEvent::DesktopMoved {
            desktop: desktop.try_into().unwrap(),
            old_index,
            new_index,
        });
        HRESULT(0)
    }

    unsafe fn virtual_desktop_name_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        (self.sender)(DesktopEvent::DesktopNameChanged(
            desktop.try_into().unwrap(),
            name.to_string(),
        ));
        HRESULT(0)
    }

    unsafe fn view_virtual_desktop_changed(&self, view: IApplicationView) -> HRESULT {
        let mut hwnd = HWND::default();
        view.get_thumbnail_window(&mut hwnd);
        (self.sender)(DesktopEvent::WindowChanged(hwnd));
        HRESULT(0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{create_event_thread, get_current_desktop, switch_desktop};
    use std::time::Duration;

    #[derive(Debug, Clone)]
    enum MainLoopEvents {
        SomeOtherEvent(String),
        DesktopEvent(DesktopEvent),
    }

    // Imp from DesktopEvent
    impl From<DesktopEvent> for MainLoopEvents {
        fn from(event: DesktopEvent) -> Self {
            MainLoopEvents::DesktopEvent(event)
        }
    }

    #[test]
    fn test_drop() {
        let (tx, rx) = crossbeam_channel::unbounded::<MainLoopEvents>();

        for _ in 0..100 {
            create_event_thread(tx.clone());
        }
    }

    #[test]
    fn test_listener_manual() {
        println!("Test thread is {:?}", std::thread::current().id());
        let (tx, rx) = crossbeam_channel::unbounded::<MainLoopEvents>();
        let notifications_thread = DesktopEventThread::new(DesktopEventSender::Crossbeam(tx));

        std::thread::spawn(|| {
            for item in rx {
                println!("Received {:?}", item);
            }
        });

        // Wait for keypress
        println!("⛔ Press enter to stop");
        let mut input = String::new();
        std::io::stdin().read_line(&mut input).unwrap();
    }

    #[test]
    fn test_switch_desktops_rapidly() {
        println!("Test thread is {:?}", std::thread::current().id());
        let (tx, rx) = crossbeam_channel::unbounded::<DesktopEvent>();
        std::thread::spawn(|| {
            for item in rx {
                println!("Received {:?}", item);
            }
        });

        let _notifications_thread = DesktopEventThread::new(DesktopEventSender::Crossbeam(tx));
        let current_desktop = get_current_desktop().unwrap();

        for _ in 0..5 {
            switch_desktop(0).unwrap();
            // std::thread::sleep(Duration::from_millis(4));
            switch_desktop(1).unwrap();
        }

        // Finally return to same desktop we were
        std::thread::sleep(Duration::from_millis(13));
        switch_desktop(current_desktop).unwrap();
        std::thread::sleep(Duration::from_millis(1000));

        println!("End of program");
    }
}
