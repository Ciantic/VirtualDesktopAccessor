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
    DispatchMessageW, GetMessageW, PostQuitMessage, PostThreadMessageW, TranslateMessage, MSG,
    WM_USER,
};

use crate::hresult::HRESULT;
use crate::interfaces::{
    ComIn, IApplicationView, IVirtualDesktop, IVirtualDesktopNotification,
    IVirtualDesktopNotification_Impl,
};
use crate::Desktop;
use crate::Result;
use crate::{comobjects::*, log_output};

#[derive(Debug, Clone)]
pub enum DesktopEventSender {
    Std(std::sync::mpsc::Sender<DesktopEvent>),

    // #[cfg(feature = "crossbeam-channel")]
    Crossbeam(crossbeam_channel::Sender<DesktopEvent>),
}

impl DesktopEventSender {
    fn try_send(&self, event: DesktopEvent) {
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
    DesktopChanged { new: Desktop, old: Desktop },
    DesktopNameChanged(Desktop, String),
    DesktopWallpaperChanged(Desktop, String),
    DesktopMoved(Desktop, i64, i64),
    WindowChanged(HWND),
}

struct DesktopEventListener<'a> {
    cookie: u32,
    ptr: Pin<Box<IVirtualDesktopNotification>>,
    com_objects: &'a ComObjects,
}

/// Starts a listener thread, returns a stopping function
pub(crate) fn start_listener_thread(sender: DesktopEventSender) -> Box<dyn FnOnce()> {
    let winapi_thread_id_pair = Arc::new((Mutex::new(0 as u32), Condvar::new()));
    let winapi_thread_id_pair_2 = Arc::clone(&winapi_thread_id_pair);
    let notification_thread = std::thread::spawn(move || {
        {
            // Send the current thread id to parent thread
            let (lock, cvar) = &*winapi_thread_id_pair_2;
            let mut started = lock.lock().unwrap();
            *started = unsafe { GetCurrentThreadId() };
            cvar.notify_one();
        }
        let com_objects = ComObjects::new();
        let _listener = DesktopEventListener::new(&com_objects, sender).unwrap();
        let mut msg = MSG::default();
        unsafe {
            while GetMessageW(&mut msg, HWND::default(), 0, 0).as_bool() {
                if msg.message == WM_USER + 0x10 {
                    PostQuitMessage(0);
                }
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
    });

    // Wait until the thread has started, and sent it's Windows specific thread id
    let win_thread_id = {
        let (lock, cvar) = &*winapi_thread_id_pair;
        let mut started = lock.lock().unwrap();
        while *started == 0 {
            started = cvar.wait(started).unwrap();
        }
        *started
    };

    Box::new(move || {
        log_output("Stopping listener thread");

        unsafe {
            PostThreadMessageW(
                win_thread_id,
                WM_USER + 0x10,
                WPARAM::default(),
                LPARAM::default(),
            );
        }
        notification_thread.join().unwrap();
    })
}

impl<'a> DesktopEventListener<'a> {
    pub fn new(
        com_objects: &'a ComObjects,
        sender: DesktopEventSender,
    ) -> Result<Pin<Box<DesktopEventListener>>> {
        let ptr = Pin::new(Box::new(VirtualDesktopNotification { sender }.into()));
        let notification = Pin::new(Box::new(DesktopEventListener {
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

impl Drop for DesktopEventListener<'_> {
    fn drop(&mut self) {
        #[cfg(debug_assertions)]
        log_output(&format!(
            "Unregistering notification {} {:?}",
            self.cookie,
            std::thread::current().id()
        ));
        self.com_objects
            .unregister_for_notifications(self.cookie)
            .unwrap();
    }
}

#[windows::core::implement(IVirtualDesktopNotification)]
struct VirtualDesktopNotification {
    sender: DesktopEventSender,
}

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
        debug_desktop(&desktop_new, "Desktop changed");
        self.sender.try_send(DesktopEvent::DesktopChanged {
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
        debug_desktop(&desktop, "Desktop wallpaper changed");
        self.sender.try_send(DesktopEvent::DesktopWallpaperChanged(
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
        debug_desktop(&desktop, "Desktop created");
        self.sender
            .try_send(DesktopEvent::DesktopCreated(desktop.try_into().unwrap()));
        HRESULT(0)
    }

    unsafe fn virtual_desktop_destroy_begin(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        // Desktop destroyed is not anymore in the stack
        debug_desktop(&desktop_destroyed, "Desktop destroy begin");
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
        debug_desktop(&desktop_destroyed, "Desktop destroyed");
        debug_desktop(&desktop_fallback, "Desktop destroyed fallback");
        self.sender.try_send(DesktopEvent::DesktopDestroyed(
            desktop_destroyed.try_into().unwrap(),
        ));
        HRESULT(0)
    }

    unsafe fn virtual_desktop_is_per_monitor_changed(&self, is_per_monitor: i32) -> HRESULT {
        #[cfg(debug_assertions)]
        log_output(&format!(
            "Desktop is per monitor changed: {}",
            is_per_monitor != 0
        ));

        HRESULT(0)
    }

    unsafe fn virtual_desktop_moved(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
        old_index: i64,
        new_index: i64,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop moved");
        self.sender.try_send(DesktopEvent::DesktopMoved(
            desktop.try_into().unwrap(),
            old_index,
            new_index,
        ));
        HRESULT(0)
    }

    unsafe fn virtual_desktop_name_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop renamed");
        self.sender.try_send(DesktopEvent::DesktopNameChanged(
            desktop.try_into().unwrap(),
            name.to_string(),
        ));
        HRESULT(0)
    }

    unsafe fn view_virtual_desktop_changed(&self, view: IApplicationView) -> HRESULT {
        let mut hwnd = HWND::default();
        view.get_thumbnail_window(&mut hwnd);

        #[cfg(debug_assertions)]
        log_output(&format!("View in desktop changed, HWND {:?}", hwnd));

        self.sender.try_send(DesktopEvent::WindowChanged(hwnd));
        HRESULT(0)
    }
}

static SENDERS: Lazy<Arc<Mutex<HashMap<u32, DesktopEventSender>>>> =
    Lazy::new(|| Arc::new(Mutex::new(HashMap::new())));

static SENDER_INDEX: Lazy<Arc<Mutex<u32>>> = Lazy::new(|| Arc::new(Mutex::new(0)));

pub fn add_event_sender(sender: &DesktopEventSender) -> u32 {
    let mut next_index = SENDER_INDEX.lock().unwrap();
    *next_index += 1;
    let mut senders = SENDERS.lock().unwrap();
    senders.insert(*next_index, sender.clone());
    *next_index
}

pub fn remove_event_sender(index: u32) {
    let mut senders = SENDERS.lock().unwrap();
    senders.remove(&index);
}

#[cfg(test)]
mod tests {
    use windows::Win32::{
        Foundation::{LPARAM, WPARAM},
        UI::WindowsAndMessaging::PostThreadMessageW,
    };

    use super::*;
    use crate::{get_current_desktop, switch_desktop};
    use std::time::Duration;

    #[test]
    fn test_listener_manual() {
        println!("This thread is {:?}", std::thread::current().id());
        let (tx, rx) = crossbeam_channel::unbounded();
        let notifications_stopper = start_listener_thread(DesktopEventSender::Crossbeam(tx));
        std::thread::spawn(|| {
            for item in rx {
                println!("Received {:?}", item);
            }
        });

        // Wait for keypress
        println!("⛔ Press enter to stop");
        let mut input = String::new();
        std::io::stdin().read_line(&mut input).unwrap();

        notifications_stopper();
    }

    #[test]
    fn test_switch_desktops_rapidly() {
        println!("This thread is {:?}", std::thread::current().id());
        let (tx, rx) = crossbeam_channel::unbounded();
        let notifications_thread_stopper = start_listener_thread(DesktopEventSender::Crossbeam(tx));

        let current_desktop = get_current_desktop().unwrap();

        for _ in 0..5 {
            switch_desktop(0).unwrap();
            // std::thread::sleep(Duration::from_millis(4));
            switch_desktop(1).unwrap();
        }

        // Finally return to same desktop we were
        std::thread::sleep(Duration::from_millis(13));
        switch_desktop(current_desktop).unwrap();
        std::thread::sleep(Duration::from_millis(13));

        for item in rx {
            println!("Received {:?}", item);
        }

        notifications_thread_stopper();
        println!("End of program");
    }
}
