// Some reason the co_class macro uses null comparison
#![allow(clippy::cmp_null)]

use com::{co_class, interfaces::IUnknown, ComPtr, ComRc};

use crate::{
    hresult::HRESULT,
    interfaces::{
        IApplicationView, IID_IVirtualDesktopNotification, IVirtualDesktop,
        IVirtualDesktopNotification, IVirtualDesktopNotificationService,
    },
    DesktopID, Error, HWND,
};
use crossbeam_channel::{Receiver, Sender};
use std::ptr;

pub enum VirtualDesktopEvent {
    DesktopCreated(DesktopID),
    DesktopDestroyed(DesktopID),
    DesktopChanged(DesktopID, DesktopID),
    WindowChanged(HWND),
}

pub struct RegisteredListener {
    // This is the value for registrations and unregistrations
    cookie: u32,

    // Listener holds the value on which the IVirtualDesktopNotificationService points
    #[allow(dead_code)]
    listener: VirtualDesktopChangeListener,

    // Receiver
    receiver: Receiver<VirtualDesktopEvent>,

    // Unregistration on drop requires a notification service
    service: ComRc<dyn IVirtualDesktopNotificationService>,
}
unsafe impl Send for RegisteredListener {}
unsafe impl Sync for RegisteredListener {}

impl RegisteredListener {
    pub fn register(
        sender: Sender<VirtualDesktopEvent>,
        receiver: Receiver<VirtualDesktopEvent>,
        service: ComRc<dyn IVirtualDesktopNotificationService>,
    ) -> Result<RegisteredListener, HRESULT> {
        let listener = *VirtualDesktopChangeListener::create(sender);
        // unsafe {
        //     listener.add_ref();
        //     // listener.add_ref();
        //     // listener.add_ref();
        //     // listener.add_ref();
        // }
        #[cfg(feature = "debug")]
        println!("Fresh listener {:?}", listener.__refcnt);
        unsafe {
            listener.add_ref();
            listener.add_ref();
            listener.add_ref();
            listener.add_ref();
            listener.add_ref();
            listener.add_ref();
        }

        // Retrieve interface pointer to IVirtualDesktopNotification
        let mut ipv = ptr::null_mut();
        // unsafe {
        //     listener.add_ref();
        // }
        #[cfg(feature = "debug")]
        println!("1 listener {:?}", listener.__refcnt);
        let res = HRESULT::from_i32(unsafe {
            listener.query_interface(&IID_IVirtualDesktopNotification, &mut ipv)
        });
        #[cfg(feature = "debug")]
        println!("2 listener {:?}", listener.__refcnt);
        // unsafe {
        //     listener.release();
        // }
        if !res.failed() && !ipv.is_null() {
            let ptr: ComRc<dyn IVirtualDesktopNotification> =
                unsafe { ComRc::from_raw(ipv as *mut *mut _) };
            // let ptr = unsafe { ComPtr::new(ipv as *mut *mut _) };
            #[cfg(feature = "debug")]
            println!("RC'dd listener {:?}", listener.__refcnt);

            // #[cfg(feature = "debug")]
            // println!(
            //     "Register a listener for IVirtualDesktopNotification {:?}",
            //     ptr.__refcnt
            // );

            // Register the IVirtualDesktopNotification to the service
            let mut cookie = 0;
            let res2 = unsafe { service.register(ptr, &mut cookie) };
            if res2.failed() {
                #[cfg(feature = "debug")]
                println!("Registration failed {:?}", res2);
                Err(res)
            } else {
                #[cfg(feature = "debug")]
                println!("Register a listener {:?} {:?}", listener.__refcnt, cookie);
                Ok(RegisteredListener {
                    cookie,
                    listener,
                    receiver,
                    service: service.clone(),
                })
            }
        } else {
            Err(res)
        }
    }

    pub fn get_receiver(&self) -> Receiver<VirtualDesktopEvent> {
        self.receiver.clone()
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
}

impl VirtualDesktopChangeListener {
    // Notice that com-rs package requires empty new, even though it's not used
    // for anything in this case, because we are not creating a COM server
    fn new() -> Box<VirtualDesktopChangeListener> {
        panic!()
        // VirtualDesktopChangeListener::allocate()
    }

    fn create(sender: Sender<VirtualDesktopEvent>) -> Box<VirtualDesktopChangeListener> {
        let v = VirtualDesktopChangeListener::allocate(sender);
        // unsafe {
        //     v.add_ref();
        // }
        v
    }
}

impl Drop for VirtualDesktopChangeListener {
    fn drop(&mut self) {
        #[cfg(feature = "debug")]
        println!("Free listener");

        // unsafe {
        //     self.release();
        // }
    }
}

impl IVirtualDesktopNotification for VirtualDesktopChangeListener {
    /// On desktop creation
    unsafe fn virtual_desktop_created(
        &self,
        desktop: ComRc<dyn IVirtualDesktop>,
    ) -> com::sys::HRESULT {
        let mut id: DesktopID = Default::default();
        desktop.get_id(&mut id);
        let _ = self.sender.send(VirtualDesktopEvent::DesktopCreated(id));
        0 //HRESULT::ok()
    }

    /// On desktop destroy begin
    unsafe fn virtual_desktop_destroy_begin(
        &self,
        _destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> com::sys::HRESULT {
        0 //HRESULT::ok()
    }

    /// On desktop destroy failed
    unsafe fn virtual_desktop_destroy_failed(
        &self,
        _destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> com::sys::HRESULT {
        0 //HRESULT::ok()
    }

    /// On desktop destory
    unsafe fn virtual_desktop_destroyed(
        &self,
        destroyed_desktop: ComRc<dyn IVirtualDesktop>,
        _fallback_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> com::sys::HRESULT {
        let mut id: DesktopID = Default::default();
        destroyed_desktop.get_id(&mut id);
        let _ = self.sender.send(VirtualDesktopEvent::DesktopDestroyed(id));
        0 //HRESULT::ok()
    }

    /// On view/window change
    unsafe fn view_virtual_desktop_changed(
        &self,
        view: ComRc<dyn IApplicationView>,
    ) -> com::sys::HRESULT {
        let mut hwnd = 0 as _;
        view.get_thumbnail_window(&mut hwnd);

        #[cfg(feature = "debug")]
        println!("-> Window changed {:?}", std::thread::current().id());

        #[cfg(feature = "debug")]
        println!("-> self ptr {:?}", self.__refcnt);

        let _ = self.sender.send(VirtualDesktopEvent::WindowChanged(hwnd));

        0 //HRESULT::ok()
    }

    /// On desktop change
    unsafe fn current_virtual_desktop_changed(
        &self,
        old_desktop: ComRc<dyn IVirtualDesktop>,
        new_desktop: ComRc<dyn IVirtualDesktop>,
    ) -> com::sys::HRESULT {
        let mut old_id: DesktopID = Default::default();
        let mut new_id: DesktopID = Default::default();
        old_desktop.get_id(&mut old_id);
        new_desktop.get_id(&mut new_id);

        #[cfg(feature = "debug")]
        println!("-> Desktop change {:?}", std::thread::current().id());

        let _ = self
            .sender
            .send(VirtualDesktopEvent::DesktopChanged(old_id, new_id));
        0 //HRESULT::ok()
    }
}
