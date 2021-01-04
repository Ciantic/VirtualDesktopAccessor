use crate::comhelpers::create_instance;
use crate::immersive::{get_immersive_service, get_immersive_service_for_class};
use crate::{
    changelistener::RegisteredListener,
    hstring::HSTRING,
    interfaces::{
        CLSID_IVirtualNotificationService, CLSID_ImmersiveShell,
        CLSID_VirtualDesktopManagerInternal, CLSID_VirtualDesktopPinnedApps, IApplicationView,
        IApplicationViewCollection, IObjectArray, IServiceProvider, IVirtualDesktop,
        IVirtualDesktop2, IVirtualDesktopManager, IVirtualDesktopManagerInternal,
        IVirtualDesktopManagerInternal2, IVirtualDesktopNotificationService,
        IVirtualDesktopPinnedApps,
    },
    Desktop, DesktopID, Error, VirtualDesktopEvent, EVENTS, HAS_LISTENERS, HRESULT, HWND,
};
use com::{ComInterface, ComRc};
use crossbeam_channel::Receiver;
use std::{cell::RefCell, sync::atomic::Ordering};

/// Provides the stateful helper to accessing the Windows 10 Virtual Desktop
/// functions.
pub struct VirtualDesktopService {
    virtual_desktop_manager: ComRc<dyn IVirtualDesktopManager>,
    virtual_desktop_manager_internal: ComRc<dyn IVirtualDesktopManagerInternal>,
    virtual_desktop_manager_internal2: ComRc<dyn IVirtualDesktopManagerInternal2>,
    virtual_desktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService>,
    app_view_collection: ComRc<dyn IApplicationViewCollection>,
    pinned_apps: ComRc<dyn IVirtualDesktopPinnedApps>,
    registered_listener: RefCell<Option<RegisteredListener>>,
}

// Let's throw the last of the remaining safety away and implement the send and
// sync 🤞. COM pointers are usually thread safe, but accessing them needs to be
// synced. Access is synced by having Mutex in main Lazy initialization.
unsafe impl Send for VirtualDesktopService {}
unsafe impl Sync for VirtualDesktopService {}

impl VirtualDesktopService {
    /// Initialize only the service, must be-created on TaskbarCreated message
    pub fn create() -> Result<Box<VirtualDesktopService>, Error> {
        let service_provider = create_instance::<dyn IServiceProvider>(&CLSID_ImmersiveShell)?;

        let virtual_desktop_manager =
            get_immersive_service::<dyn IVirtualDesktopManager>(&service_provider)?;

        let virtual_desktop_manager_internal = get_immersive_service_for_class(
            &service_provider,
            CLSID_VirtualDesktopManagerInternal,
        )?;

        let virtual_desktop_manager_internal2: ComRc<dyn IVirtualDesktopManagerInternal2> =
            get_immersive_service_for_class(
                &service_provider,
                CLSID_VirtualDesktopManagerInternal,
            )?;

        let app_view_collection = get_immersive_service(&service_provider)?;

        let pinned_apps =
            get_immersive_service_for_class(&service_provider, CLSID_VirtualDesktopPinnedApps)?;

        let virtual_desktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService> =
            get_immersive_service_for_class(&service_provider, CLSID_IVirtualNotificationService)?;

        #[cfg(feature = "debug")]
        println!("VirtualDesktopService created.");

        Ok(Box::new(VirtualDesktopService {
            registered_listener: if HAS_LISTENERS.load(Ordering::SeqCst) {
                #[cfg(feature = "debug")]
                println!("Has listeners, so try to recreate...");
                RefCell::new(Some(RegisteredListener::register(
                    EVENTS.0.clone(),
                    EVENTS.1.clone(),
                    virtual_desktop_notification_service.clone(),
                )?))
            } else {
                RefCell::new(None)
            },
            virtual_desktop_manager,
            virtual_desktop_manager_internal,
            virtual_desktop_manager_internal2,
            app_view_collection,
            virtual_desktop_notification_service,
            pinned_apps,
        }))
    }

    /// Get raw desktop list
    fn _get_idesktops(&self) -> Result<Vec<ComRc<dyn IVirtualDesktop>>, Error> {
        let mut ptr = None;
        Result::from(unsafe { self.virtual_desktop_manager_internal.get_desktops(&mut ptr) })?;
        match ptr {
            Some(objectarray) => {
                let mut count = 0;
                Result::from(unsafe { objectarray.get_count(&mut count) })?;
                let mut desktops: Vec<ComRc<dyn IVirtualDesktop>> = vec![];

                for i in 0..count {
                    let mut ptr = std::ptr::null_mut();
                    Result::from(unsafe {
                        objectarray.get_at(i, &IVirtualDesktop::IID, &mut ptr)
                    })?;
                    let desktop = unsafe { ComRc::from_raw(ptr as *mut _) };
                    desktops.push(desktop);
                }
                Ok(desktops)
            }
            None => Err(Error::ComAllocatedNullPtr),
        }
    }

    /// Get raw desktop by ID
    fn _get_idesktop_by_id(
        &self,
        desktop: &DesktopID,
    ) -> Result<ComRc<dyn IVirtualDesktop>, Error> {
        let mut o = None;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .find_desktop(desktop, &mut o)
        })
        .map_err(|hr| match hr {
            // Does not exist
            Error::ComError(HRESULT(0x8002802B)) => Error::DesktopNotFound,
            e => e,
        })?;

        if let Some(d) = o {
            Ok(d)
        } else {
            Err(Error::DesktopNotFound)
        }
    }

    /// Get application view for raw window
    fn _get_iapplication_view_for_hwnd(
        &self,
        hwnd: HWND,
    ) -> Result<ComRc<dyn IApplicationView>, Error> {
        let mut ptr = None;
        Result::from(unsafe {
            self.app_view_collection
                .get_view_for_hwnd(hwnd as _, &mut ptr)
        })
        .map_err(|hr| match hr {
            // View does not exist
            Error::ComError(HRESULT(0x8002802B)) => Error::WindowNotFound,
            e => e,
        })?;
        match ptr {
            Some(ptr) => Ok(ptr),
            None => Err(Error::ComAllocatedNullPtr),
        }
    }

    /// Get event receiver
    pub fn get_event_receiver(&self) -> Result<Receiver<VirtualDesktopEvent>, Error> {
        #[cfg(feature = "debug")]
        println!("Get event receiver...");

        let v = self.registered_listener.borrow();
        match v.as_ref() {
            Some(listener) => Ok(listener.get_receiver()),
            None => {
                drop(v);
                let _ = self
                    .registered_listener
                    .replace(Some(RegisteredListener::register(
                        EVENTS.0.clone(),
                        EVENTS.1.clone(),
                        self.virtual_desktop_notification_service.clone(),
                    )?));
                Ok(EVENTS.1.clone())
            }
        }
    }

    /// Get desktop index
    pub fn get_desktop_by_index(&self, index: usize) -> Result<Desktop, Error> {
        self.get_desktops()?
            .get(index)
            .cloned()
            .ok_or(Error::DesktopNotFound)
    }

    /// Get desktop index
    pub fn get_index_by_desktop(&self, desktop: &Desktop) -> Result<usize, Error> {
        self.get_desktops()?
            .iter()
            .position(|x| x == desktop)
            .ok_or(Error::DesktopNotFound)
    }

    /// Rename desktop
    pub fn rename_desktop(&self, desktop: &Desktop, name: &str) -> Result<(), Error> {
        let idesktop = self._get_idesktop_by_id(&desktop.id)?;
        let hstring = HSTRING::create(name).map_err(HRESULT::from_i32)?;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal2
                .rename_desktop(idesktop, hstring)
        })?;
        Ok(())
    }

    /// Get desktop name
    pub fn get_desktop_name(&self, desktop: &Desktop) -> Result<String, Error> {
        let mut ptr = None;
        Result::from(unsafe { self.virtual_desktop_manager_internal.get_desktops(&mut ptr) })?;
        match ptr {
            Some(objectarray) => {
                let mut count = 0;
                Result::from(unsafe { objectarray.get_count(&mut count) })?;
                for i in 0..count {
                    let mut ptr = std::ptr::null_mut();
                    Result::from(unsafe {
                        objectarray.get_at(i, &IVirtualDesktop2::IID, &mut ptr)
                    })?;
                    let idesktop: ComRc<dyn IVirtualDesktop2> =
                        unsafe { ComRc::from_raw(ptr as *mut _) };
                    let mut cdesktop = Desktop::empty();
                    Result::from(unsafe { idesktop.get_id(&mut cdesktop.id) })?;
                    if &cdesktop == desktop {
                        let mut hstr = HSTRING::create("").map_err(HRESULT::from_i32)?;
                        Result::from(unsafe { idesktop.get_name(&mut hstr) })?;
                        if let Some(s) = hstr.get() {
                            return Ok(s);
                        } else {
                            return Ok("".to_string());
                        }
                    }
                }
                Err(Error::DesktopNotFound)
            }
            None => Err(Error::ComAllocatedNullPtr),
        }
    }

    /// Get desktop IDs
    pub fn get_desktops(&self) -> Result<Vec<Desktop>, Error> {
        self._get_idesktops()?
            .iter()
            .map(|f| {
                let mut desktop = Desktop::empty();
                Result::from(unsafe { f.get_id(&mut desktop.id) }).map(|_| desktop)
            })
            .collect()
    }

    /// Get number of desktops
    /*
    pub fn get_desktop_count(&self) -> Result<u32, Error> {
        let mut ptr = None;
        Result::from(unsafe { self.virtual_desktop_manager_internal.get_desktops(&mut ptr) })?;
        match ptr {
            Some(objectarray) => {
                let mut count = 0;
                Result::from(unsafe { objectarray.get_count(&mut count) }).map(|_| count)
            }
            None => Err(Error::ComAllocatedNullPtr),
        }
    }
    */

    /// Get current desktop ID
    pub fn get_current_desktop(&self) -> Result<Desktop, Error> {
        let mut ptr: Option<ComRc<dyn IVirtualDesktop>> = None;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .get_current_desktop(&mut ptr)
        })?;
        match ptr {
            Some(idesktop) => {
                let mut desktop = Desktop::empty();
                Result::from(unsafe { idesktop.get_id(&mut desktop.id) }).map(|_| desktop)
            }
            None => Err(Error::ComAllocatedNullPtr),
        }
    }

    /// Get window desktop ID
    pub fn get_desktop_by_window(&self, hwnd: HWND) -> Result<Desktop, Error> {
        let mut desktop = Desktop::empty();
        Result::from(unsafe {
            self.virtual_desktop_manager
                .get_desktop_by_window(hwnd as _, &mut desktop.id)
        })
        .map_err(|er| match er {
            Error::ComError(HRESULT(0x8002802B)) => Error::WindowNotFound,
            e => e,
        })
        .map(|_| desktop)
    }

    /// Is window on current virtual desktop
    pub fn is_window_on_current_desktop(&self, hwnd: HWND) -> Result<bool, Error> {
        let mut isit = false;
        Result::from(unsafe {
            self.virtual_desktop_manager
                .is_window_on_current_desktop(hwnd as _, &mut isit)
        })
        .map(|_| isit)
    }

    /// Is window on desktop
    pub fn is_window_on_desktop(&self, hwnd: HWND, desktop: &Desktop) -> Result<bool, Error> {
        let window_desktop = self.get_desktop_by_window(hwnd)?;
        Ok(&window_desktop == desktop)
    }

    /// Move window to desktop
    pub fn move_window_to_desktop(&self, hwnd: HWND, desktop: &Desktop) -> Result<(), Error> {
        let idesktop = self._get_idesktop_by_id(&desktop.id)?;
        let ptr = idesktop
            .get_interface::<dyn IVirtualDesktop>()
            .ok_or(Error::DesktopNotFound)?;
        let view = self._get_iapplication_view_for_hwnd(hwnd)?;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .move_view_to_desktop(view, ptr)
        })
    }

    /// Go to desktop
    pub fn go_to_desktop(&self, desktop: &Desktop) -> Result<(), Error> {
        let idesktop = self._get_idesktop_by_id(&desktop.id)?;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .switch_desktop(idesktop)
        })
    }

    pub fn create_desktop(&self) -> Result<Desktop, Error> {
        let mut idesk_opt: Option<ComRc<dyn IVirtualDesktop>> = None;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .create_desktop(&mut idesk_opt)
        })?;
        let idesk = idesk_opt.ok_or(Error::CreateDesktopFailed)?;
        let mut new_desk = Desktop::empty();
        Result::from(unsafe { idesk.get_id(&mut new_desk.id) })?;
        if new_desk.id == DesktopID::default() {
            Err(Error::CreateDesktopFailed)
        } else {
            Ok(new_desk)
        }
    }

    pub fn remove_desktop(&self) -> Result<(), Error> {}

    /// Is window pinned?
    pub fn is_pinned_window(&self, hwnd: HWND) -> Result<bool, Error> {
        let view = self._get_iapplication_view_for_hwnd(hwnd)?;
        let mut test: bool = false;
        Result::from(unsafe { self.pinned_apps.is_view_pinned(view, &mut test) }).map(|_| test)
    }

    /// Pin window
    pub fn pin_window(&self, hwnd: HWND) -> Result<(), Error> {
        let view = self._get_iapplication_view_for_hwnd(hwnd)?;
        Result::from(unsafe { self.pinned_apps.pin_view(view) })
    }

    /// Unpin window
    pub fn unpin_window(&self, hwnd: HWND) -> Result<(), Error> {
        let view = self._get_iapplication_view_for_hwnd(hwnd)?;
        Result::from(unsafe { self.pinned_apps.unpin_view(view) })
    }
}

// #[cfg(debug_assertions)]
// #[cfg(feature = "debug")]
// impl Drop for VirtualDesktopService {
//     fn drop(&mut self) {
//         // This panics on debug mode
//         println!("Deallocate VirtualDesktopService in thread.");
//     }
// }
