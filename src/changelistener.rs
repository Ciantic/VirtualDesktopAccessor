// Some reason the co_class macro uses null comparison
#![allow(clippy::cmp_null)]

use com::{co_class, interfaces::IUnknown, ComRc};

use crate::{
    hresult::HRESULT,
    interfaces::{
        IApplicationView, IObjectArray, IVirtualDesktop, IVirtualDesktopNotification,
        IVirtualDesktopNotificationService,
    },
    Desktop, HWND,
};
use crossbeam_channel::{Receiver, Sender};

unsafe impl Send for VirtualDesktopEventSender {}
unsafe impl Sync for VirtualDesktopEventSender {}

#[derive(Debug, Clone)]
pub enum VirtualDesktopEventSender {
    Std(std::sync::mpsc::Sender<VirtualDesktopEvent>),
    // #[cfg(feature = "crossbeam-channel")]
    // Crossbeam(crossbeam_channel::Sender<VirtualDesktopEvent>),
}

impl VirtualDesktopEventSender {
    pub fn try_send(&self, event: VirtualDesktopEvent) -> Result<(), String> {
        match self {
            VirtualDesktopEventSender::Std(sender) => sender
                .send(event)
                .map_err(|_| "Failed to send event".to_owned()),
            //
            // #[cfg(feature = "crossbeam-channel")]
            // VirtualDesktopEventSender::Crossbeam(sender) => {
            //     sender.send(event).expect("Failed to send event");
            // }
        }
    }
}

pub enum VirtualDesktopEvent {
    DesktopCreated(Desktop),
    DesktopDestroyed(Desktop),
    DesktopChanged(Desktop, Desktop),
    DesktopNameChanged(Desktop, String),
    DesktopWallpaperChanged(Desktop, String),
    DesktopMoved(Desktop, i64, i64),
    WindowChanged(HWND),
}

pub struct RegisteredListener {
    // This is the value for registrations and unregistrations
    cookie: u32,

    // Listener holds the value on which the IVirtualDesktopNotificationService points
    #[allow(dead_code)]
    listener: Box<VirtualDesktopChangeListener>,

    // Receiver
    // receiver: Receiver<VirtualDesktopEvent>,
    // sender: VirtualDesktopEventSender,

    // Unregistration on drop requires a notification service
    service: ComRc<dyn IVirtualDesktopNotificationService>,
}
unsafe impl Send for RegisteredListener {}
unsafe impl Sync for RegisteredListener {}

impl RegisteredListener {
    pub fn register(
        sender: VirtualDesktopEventSender,
        service: ComRc<dyn IVirtualDesktopNotificationService>,
    ) -> Result<RegisteredListener, HRESULT> {
        let listener = VirtualDesktopChangeListener::create(sender);
        let ptr: ComRc<dyn IVirtualDesktopNotification> = unsafe {
            ComRc::from_raw(&listener.__ivirtualdesktopnotificationvptr as *const _ as *mut _)
        };

        // Register the IVirtualDesktopNotification to the service
        let mut cookie = 0;
        let res = unsafe { service.register(ptr.clone(), &mut cookie) };
        if res.failed() {
            #[cfg(feature = "debug")]
            println!("Registration failed {:?}", res);

            Err(res)
        } else {
            #[cfg(feature = "debug")]
            println!(
                "Register a listener {:?} {:?} {:?}",
                listener.__refcnt,
                cookie,
                std::thread::current().id()
            );

            Ok(RegisteredListener {
                cookie,
                listener,
                //sender,
                service: service.clone(),
            })
        }
    }
    pub(crate) fn get_sender(&self) -> &VirtualDesktopEventSender {
        &self.listener.get_sender()
    }
}

impl Drop for RegisteredListener {
    fn drop(&mut self) {
        #[cfg(feature = "debug")]
        println!("Unregister a listener {:?}", self.cookie);
        unsafe {
            self.service.unregister(self.cookie);
        }
    }
}

#[co_class(implements(IVirtualDesktopNotification))]
struct VirtualDesktopChangeListener {
    sender: VirtualDesktopEventSender,
}

impl VirtualDesktopChangeListener {
    // Notice that com-rs package requires empty new, even though it's not used
    // for anything in this case, because we are not creating a COM server
    fn new() -> Box<VirtualDesktopChangeListener> {
        panic!()
        // VirtualDesktopChangeListener::allocate()
    }

    fn create(sender: VirtualDesktopEventSender) -> Box<VirtualDesktopChangeListener> {
        let v = VirtualDesktopChangeListener::allocate(sender);
        unsafe {
            v.add_ref();
        }
        v
    }

    pub(crate) fn get_sender(&self) -> &VirtualDesktopEventSender {
        &self.sender
    }
}

impl Drop for VirtualDesktopChangeListener {
    fn drop(&mut self) {
        #[cfg(feature = "debug")]
        println!("Drop VirtualDesktopChangeListener");
        unsafe {
            self.release();
        }
    }
}

impl IVirtualDesktopNotification for VirtualDesktopChangeListener {
    /// On desktop creation
    unsafe fn virtual_desktop_created(
        &self,
        _monitors: ComRc<dyn IObjectArray>,
        idesktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        let mut desktop: Desktop = Desktop::empty();
        idesktop.get_id(&mut desktop.id);
        let _ = self
            .sender
            .try_send(VirtualDesktopEvent::DesktopCreated(desktop));
        HRESULT::ok()
    }

    /// On desktop destroy begin
    unsafe fn virtual_desktop_destroy_begin(
        &self,
        _monitors: ComRc<dyn IObjectArray>,
        _destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        HRESULT::ok()
    }

    /// On desktop destroy failed
    unsafe fn virtual_desktop_destroy_failed(
        &self,
        _monitors: ComRc<dyn IObjectArray>,
        _destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        HRESULT::ok()
    }

    /// On desktop destory
    unsafe fn virtual_desktop_destroyed(
        &self,
        _monitors: ComRc<dyn IObjectArray>,
        destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        let mut desktop = Desktop::empty();
        destroyed_desktop.get_id(&mut desktop.id);
        let _ = self
            .sender
            .try_send(VirtualDesktopEvent::DesktopDestroyed(desktop));

        HRESULT::ok()
    }

    unsafe fn virtual_desktop_is_per_monitor_changed(&self, isPerMonitor: bool) -> HRESULT {
        // TODO: !?
        HRESULT::ok()
    }

    unsafe fn virtual_desktop_moved(
        &self,
        _monitors: ComRc<dyn IObjectArray>,
        desktop: ComRc<dyn IVirtualDesktop>,
        old_index: i64,
        new_index: i64,
    ) -> HRESULT {
        let mut new = Desktop::empty();
        desktop.get_id(&mut new.id);
        let _ = self
            .sender
            .try_send(VirtualDesktopEvent::DesktopMoved(new, old_index, new_index));
        HRESULT::ok()
    }

    unsafe fn virtual_desktop_name_changed(
        &self,
        desktop: ComRc<dyn IVirtualDesktop>,
        name: crate::hstring::HSTRING,
    ) -> HRESULT {
        // TODO: ! THIS IS NOT CALLED ! ??
        let namestr = name.get().unwrap();
        let mut new = Desktop::empty();
        desktop.get_id(&mut new.id);

        let _ = self
            .sender
            .try_send(VirtualDesktopEvent::DesktopNameChanged(new, namestr));
        HRESULT::ok()
    }

    /// On view/window change
    unsafe fn view_virtual_desktop_changed(&self, view: ComRc<dyn IApplicationView>) -> HRESULT {
        let mut hwnd = 0 as _;
        view.get_thumbnail_window(&mut hwnd);

        let _ = self
            .sender
            .try_send(VirtualDesktopEvent::WindowChanged(hwnd));

        HRESULT::ok()
    }

    /// On desktop change
    unsafe fn current_virtual_desktop_changed(
        &self,
        _monitors: ComRc<dyn IObjectArray>,
        old_desktop: ComRc<dyn IVirtualDesktop>,
        new_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        let mut old = Desktop::empty();
        let mut new = Desktop::empty();
        old_desktop.get_id(&mut old.id);
        new_desktop.get_id(&mut new.id);

        let _ = self
            .sender
            .try_send(VirtualDesktopEvent::DesktopChanged(old, new));

        HRESULT::ok()
    }

    unsafe fn virtual_desktop_wallpaper_changed(
        &self,
        desktop: ComRc<dyn IVirtualDesktop>,
        name: crate::hstring::HSTRING,
    ) -> HRESULT {
        let namestr = name.get().unwrap();
        let mut new = Desktop::empty();
        desktop.get_id(&mut new.id);

        let _ = self
            .sender
            .try_send(VirtualDesktopEvent::DesktopWallpaperChanged(new, namestr));
        HRESULT::ok()
    }

    // // Virtual desktop was renamed
    // unsafe fn virtual_desktop_renamed(
    //     &self,
    //     desktop: ComRc<dyn IVirtualDesktop>,
    //     newName: HSTRING,
    // ) -> HRESULT {
    //     HRESULT::ok()
    // }
}
