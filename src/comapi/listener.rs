use std::collections::HashMap;
use std::convert::{TryFrom, TryInto};
use std::pin::Pin;
use std::sync::{Arc, Mutex};
use std::{cell::RefCell, rc::Rc};

use once_cell::sync::Lazy;
use windows::core::{GUID, HSTRING};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::Shell::Common::IObjectArray;
use windows::Win32::UI::WindowsAndMessaging::{
    DispatchMessageW, GetMessageW, PostQuitMessage, TranslateMessage, MSG, WM_USER,
};

use crate::hresult::HRESULT;
use crate::Desktop;

use super::interfaces::{
    ComIn, IApplicationView, IVirtualDesktop, IVirtualDesktopNotification,
    IVirtualDesktopNotification_Impl,
};

use super::raw2::*;
use super::Result;

#[derive(Debug, Clone)]
pub enum VirtualDesktopEventSender {
    Std(std::sync::mpsc::Sender<VirtualDesktopEvent>),

    #[cfg(feature = "crossbeam-channel")]
    Crossbeam(crossbeam_channel::Sender<VirtualDesktopEvent>),
}

impl VirtualDesktopEventSender {
    fn try_send(&self, event: VirtualDesktopEvent) -> Result<()> {
        match self {
            VirtualDesktopEventSender::Std(sender) => {
                sender.send(event).map_err(|_| Error::SenderError)
            }

            #[cfg(feature = "crossbeam-channel")]
            VirtualDesktopEventSender::Crossbeam(sender) => {
                sender.try_send(event).map_err(|_| Error::SenderError)
            }
        }
    }
}

#[derive(Debug, Clone)]
pub enum VirtualDesktopEvent {
    DesktopCreated(Desktop),
    DesktopDestroyed(Desktop),
    DesktopChanged { new: Desktop, old: Desktop },
    DesktopNameChanged(Desktop, String),
    DesktopWallpaperChanged(Desktop, String),
    DesktopMoved(Desktop, i64, i64),
    WindowChanged(HWND),
}

struct SimpleVirtualDesktopNotificationWrapper {
    cookie: u32,
    ptr: IVirtualDesktopNotification,
    number_times_desktop_changed: Rc<RefCell<u32>>,
}

impl SimpleVirtualDesktopNotificationWrapper {
    pub fn new(
        sender: VirtualDesktopEventSender,
    ) -> Result<Pin<Box<SimpleVirtualDesktopNotificationWrapper>>> {
        println!(
            "Notification service created in thread {:?}",
            std::thread::current().id()
        );
        let number_times_desktop_changed = Rc::new(RefCell::new(0));

        let ptr = SimpleVirtualDesktopNotification { sender };
        let mut notification = Pin::new(Box::new(SimpleVirtualDesktopNotificationWrapper {
            cookie: 0,
            ptr: ptr.into(),
            number_times_desktop_changed,
        }));

        notification.cookie = com_objects().register_for_notifications(&notification.ptr)?;
        println!(
            "Registered notification {} {:?}",
            notification.cookie,
            std::thread::current().id()
        );

        Ok(notification)
    }

    pub fn msg_loop(&self) {
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
    }
}

#[windows::core::implement(IVirtualDesktopNotification)]
struct SimpleVirtualDesktopNotification {
    sender: VirtualDesktopEventSender,
}

fn debug_desktop(desktop_new: &IVirtualDesktop, prefix: &str) {
    let mut gid = GUID::default();
    unsafe { desktop_new.get_id(&mut gid).panic_if_failed() };

    let name = "";

    // let mut name = HSTRING::new();
    // unsafe { desktop_new.get_name(&mut name).panic_if_failed() };

    println!(
        "{}: {:?} {:?} {:?}",
        prefix,
        gid,
        name.to_string(),
        std::thread::current().id()
    );
}

// Allow unused variable warnings
#[allow(unused_variables)]
impl IVirtualDesktopNotification_Impl for SimpleVirtualDesktopNotification {
    unsafe fn current_virtual_desktop_changed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_old: ComIn<IVirtualDesktop>,
        desktop_new: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        debug_desktop(&desktop_new, "Desktop changed");
        self.sender
            .try_send(VirtualDesktopEvent::DesktopChanged {
                old: desktop_old.try_into().unwrap(),
                new: desktop_new.try_into().unwrap(),
            })
            .unwrap();
        HRESULT(0)
    }

    unsafe fn virtual_desktop_wallpaper_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop wallpaper changed");
        self.sender
            .try_send(VirtualDesktopEvent::DesktopWallpaperChanged(
                desktop.try_into().unwrap(),
                name.to_string(),
            ))
            .unwrap();
        HRESULT(0)
    }

    unsafe fn virtual_desktop_created(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop created");
        self.sender
            .try_send(VirtualDesktopEvent::DesktopCreated(
                desktop.try_into().unwrap(),
            ))
            .unwrap();
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
        self.sender
            .try_send(VirtualDesktopEvent::DesktopDestroyed(
                desktop_destroyed.try_into().unwrap(),
            ))
            .unwrap();
        HRESULT(0)
    }

    unsafe fn virtual_desktop_is_per_monitor_changed(&self, is_per_monitor: i32) -> HRESULT {
        println!("Desktop is per monitor changed: {}", is_per_monitor != 0);
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
        self.sender
            .try_send(VirtualDesktopEvent::DesktopMoved(
                desktop.try_into().unwrap(),
                old_index,
                new_index,
            ))
            .unwrap();
        HRESULT(0)
    }

    unsafe fn virtual_desktop_name_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT {
        debug_desktop(&desktop, "Desktop renamed");
        self.sender
            .try_send(VirtualDesktopEvent::DesktopNameChanged(
                desktop.try_into().unwrap(),
                name.to_string(),
            ))
            .unwrap();
        HRESULT(0)
    }

    unsafe fn view_virtual_desktop_changed(&self, view: IApplicationView) -> HRESULT {
        let mut hwnd = HWND::default();
        view.get_thumbnail_window(&mut hwnd);
        println!("View in desktop changed, HWND {:?}", hwnd);
        self.sender
            .try_send(VirtualDesktopEvent::WindowChanged(hwnd))
            .unwrap();
        HRESULT(0)
    }
}

static SENDERS: Lazy<Arc<Mutex<HashMap<u32, VirtualDesktopEventSender>>>> =
    Lazy::new(|| Arc::new(Mutex::new(HashMap::new())));

static SENDER_INDEX: Lazy<Arc<Mutex<u32>>> = Lazy::new(|| Arc::new(Mutex::new(0)));

pub fn add_event_sender(sender: &VirtualDesktopEventSender) -> u32 {
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
    use windows::Win32::UI::WindowsAndMessaging::{
        DispatchMessageW, GetMessageW, PostQuitMessage, TranslateMessage, MSG, WM_USER,
    };

    use super::*;

    #[test]
    fn test_listener_manual() {
        println!("This thread is {:?}", std::thread::current().id());
        let (tx, rx) = crossbeam_channel::unbounded();
        let notifications_thread = std::thread::spawn(|| {
            let notifications = SimpleVirtualDesktopNotificationWrapper::new(
                VirtualDesktopEventSender::Crossbeam(tx),
            )
            .unwrap();
            notifications.msg_loop();
        });

        for item in rx {
            println!("Received {:?}", item);
        }

        notifications_thread.join().unwrap();
    }
}
