// Some reason the co_class macro uses null comparison
#![allow(clippy::cmp_null)]

use com::{co_class, interfaces::IUnknown, ComRc};

use crate::{
    comhelpers::create_instance,
    hresult::HRESULT,
    immersive::get_immersive_service_for_class,
    interfaces::{
        CLSID_IVirtualNotificationService, CLSID_ImmersiveShell, IApplicationView,
        IID_IVirtualDesktopNotification, IServiceProvider, IVirtualDesktop,
        IVirtualDesktopNotification, IVirtualDesktopNotificationService,
    },
    DesktopID, Error, HWND,
};
use crossbeam_channel::{unbounded, Receiver, Sender};
use std::{cell::Cell, ptr, sync::Mutex};

pub enum VirtualDesktopEvent {
    DesktopCreated(DesktopID),
    DesktopDestroyed(DesktopID),
    DesktopChanged(DesktopID, DesktopID),
    WindowChanged(HWND),
}

fn recreate_listener(sender: Sender<VirtualDesktopEvent>) -> Result<RegisteredListener, Error> {
    let service_provider = create_instance::<dyn IServiceProvider>(&CLSID_ImmersiveShell)?;
    let virtualdesktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService> =
        get_immersive_service_for_class(&service_provider, CLSID_IVirtualNotificationService)?;

    RegisteredListener::new(virtualdesktop_notification_service, sender)
}

pub struct EventListener {
    listener: Mutex<Cell<Result<RegisteredListener, Error>>>,
    sender: Sender<VirtualDesktopEvent>,
    receiver: Receiver<VirtualDesktopEvent>,
}

impl EventListener {
    pub fn new() -> EventListener {
        let (sender, receiver) = unbounded();
        EventListener {
            listener: Mutex::new(Cell::new(recreate_listener(sender.clone()))),
            sender,
            receiver,
        }
    }

    pub fn recreate(&self) -> Result<(), Error> {
        // return Ok(());
        match self.listener.lock() {
            Ok(cell) => {
                cell.replace(recreate_listener(self.sender.clone()))?;
                Ok(())
            }
            Err(_) => Err(Error::ComAllocatedNullPtr),
        }
    }
}

unsafe impl Send for EventListener {}
unsafe impl Sync for EventListener {}

fn register(
    service: &ComRc<dyn IVirtualDesktopNotificationService>,
) -> Result<(u32, VirtualDesktopChangeListener), HRESULT> {
    let listener: Box<VirtualDesktopChangeListener> = VirtualDesktopChangeListener::new();

    // Retrieve interface pointer to IVirtualDesktopNotification
    let mut ipv = ptr::null_mut();
    let res = HRESULT::from_i32(unsafe {
        listener.query_interface(&IID_IVirtualDesktopNotification, &mut ipv)
    });
    if !res.failed() && !ipv.is_null() {
        let ptr: ComRc<dyn IVirtualDesktopNotification> =
            unsafe { ComRc::from_raw(ipv as *mut *mut _) };

        // Register the IVirtualDesktopNotification to the service
        let mut cookie = 0;
        let res2 = unsafe { service.register(ptr, &mut cookie) };
        if res2.failed() {
            Err(res)
        } else {
            #[cfg(feature = "debug")]
            println!("Register a listener {:?}", cookie);
            Ok((cookie, *listener))
        }
    } else {
        Err(res)
    }
}

struct RegisteredListener {
    cookie: u32,
    listener: VirtualDesktopChangeListener,
    sender: Sender<VirtualDesktopEvent>,
    service: ComRc<dyn IVirtualDesktopNotificationService>,
}

impl RegisteredListener {
    pub fn new(
        service: ComRc<dyn IVirtualDesktopNotificationService>,
        sender: Sender<VirtualDesktopEvent>,
    ) -> Result<RegisteredListener, Error> {
        let (cookie, listener) = register(&service)?;
        Ok(RegisteredListener {
            cookie,
            listener,
            service,
            sender,
        })
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
    sender: Sender<VirtualDesktopEvent>,
    // _on_desktop_change: RefCell<Option<Box<OnDesktopChange>>>,
    // _on_desktop_created: RefCell<Option<Box<OnDesktopCreated>>>,
    // _on_desktop_destroyed: RefCell<Option<Box<OnDesktopDestroyed>>>,
    // _on_window_change: RefCell<Option<Box<OnDesktopWindowChange>>>,
}

impl VirtualDesktopChangeListener {
    fn new() -> Box<VirtualDesktopChangeListener> {
        panic!()
    }

    fn create(sender: Sender<VirtualDesktopEvent>) -> Box<VirtualDesktopChangeListener> {
        VirtualDesktopChangeListener::allocate(sender)
    }

    // Try to create and register a change listener
}

impl IVirtualDesktopNotification for VirtualDesktopChangeListener {
    unsafe fn virtual_desktop_created(&self, desktop: ComRc<dyn IVirtualDesktop>) -> HRESULT {
        let mut id: DesktopID = Default::default();
        desktop.get_id(&mut id);
        // TODO: ...
        HRESULT::ok()
    }
    unsafe fn virtual_desktop_destroy_begin(
        &self,
        _destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        HRESULT::ok()
    }
    unsafe fn virtual_desktop_destroy_failed(
        &self,
        _destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        HRESULT::ok()
    }
    unsafe fn virtual_desktop_destroyed(
        &self,
        destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        let mut id: DesktopID = Default::default();
        destroyed_desktop.get_id(&mut id);
        // TODO: ...
        HRESULT::ok()
    }
    unsafe fn view_virtual_desktop_changed(&self, view: ComRc<dyn IApplicationView>) -> HRESULT {
        let mut hwnd = 0 as _;
        view.get_thumbnail_window(&mut hwnd);
        // TODO: ...
        HRESULT::ok()
    }
    unsafe fn current_virtual_desktop_changed(
        &self,
        old_desktop: ComRc<dyn IVirtualDesktop>,
        new_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT {
        let mut old_id: DesktopID = Default::default();
        let mut new_id: DesktopID = Default::default();
        old_desktop.get_id(&mut old_id);
        new_desktop.get_id(&mut new_id);
        // TODO: ...
        HRESULT::ok()
    }
}
