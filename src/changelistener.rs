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
use std::{
    cell::{Cell, RefCell},
    ptr,
    sync::{Arc, Mutex, RwLock},
};

pub enum VirtualDesktopEvent {
    DesktopChanged(DesktopID),
    WindowChanged(HWND),
}

pub struct EventListener {
    pub sender: Sender<VirtualDesktopEvent>,
    pub receiver: Receiver<VirtualDesktopEvent>,
    listener: Mutex<Cell<VirtualDesktopChangeListener>>,
}

unsafe impl Send for EventListener {}
unsafe impl Sync for EventListener {}

fn recreate_listener() -> Result<VirtualDesktopChangeListener, Error> {
    let service_provider = create_instance::<dyn IServiceProvider>(&CLSID_ImmersiveShell)?;
    let virtualdesktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService> =
        get_immersive_service_for_class(&service_provider, CLSID_IVirtualNotificationService)?;
    Ok(*VirtualDesktopChangeListener::register(
        virtualdesktop_notification_service,
    )?)
}

impl EventListener {
    pub fn new() -> Result<EventListener, Error> {
        let (sender, receiver) = unbounded::<VirtualDesktopEvent>();
        Ok(EventListener {
            sender,
            receiver,
            listener: Mutex::new(Cell::new(recreate_listener()?)),
        })
    }

    pub fn recreate(&self) -> Result<(), Error> {
        // return Ok(());
        match self.listener.lock() {
            Ok(v) => {
                v.replace(recreate_listener()?);
                Ok(())
            }
            Err(_) => Err(Error::ComAllocatedNullPtr),
        }
    }
}

#[co_class(implements(IVirtualDesktopNotification))]
struct VirtualDesktopChangeListener {
    service: Cell<Option<ComRc<dyn IVirtualDesktopNotificationService>>>,
    cookie: Cell<u32>,
    sender: Cell<Option<Sender<VirtualDesktopEvent>>>,
    // _on_desktop_change: RefCell<Option<Box<OnDesktopChange>>>,
    // _on_desktop_created: RefCell<Option<Box<OnDesktopCreated>>>,
    // _on_desktop_destroyed: RefCell<Option<Box<OnDesktopDestroyed>>>,
    // _on_window_change: RefCell<Option<Box<OnDesktopWindowChange>>>,
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

impl Drop for VirtualDesktopChangeListener {
    fn drop(&mut self) {
        if self.cookie.get() == 0 {
            return;
        }
        if let Some(s) = self.service.take() {
            #[cfg(feature = "debug")]
            println!("Unregister a listener {:?}", self.cookie.get());
            unsafe {
                s.unregister(self.cookie.get());
            }
        }
    }
}

impl VirtualDesktopChangeListener {
    // Try to create and register a change listener
    pub(crate) fn register(
        service: ComRc<dyn IVirtualDesktopNotificationService>,
    ) -> Result<Box<VirtualDesktopChangeListener>, HRESULT> {
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

                listener.service.set(Some(service));
                listener.cookie.set(cookie);
                Ok(listener)
            }
        } else {
            Err(res)
        }
    }

    fn new() -> Box<VirtualDesktopChangeListener> {
        VirtualDesktopChangeListener::allocate(Cell::new(None), Cell::new(0), Cell::new(None))
    }
}
