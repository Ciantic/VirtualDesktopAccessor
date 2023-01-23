use crate::changelistener::VirtualDesktopEventSender;
use crate::comhelpers::create_instance;
use crate::immersive::{get_immersive_service, get_immersive_service_for_class};
use crate::{
    changelistener::RegisteredListener,
    hstring::HSTRING,
    interfaces::{
        CLSID_IVirtualNotificationService, CLSID_ImmersiveShell,
        CLSID_VirtualDesktopManagerInternal, CLSID_VirtualDesktopPinnedApps, IApplicationView,
        IApplicationViewCollection, IObjectArray, IServiceProvider, IVirtualDesktop,
        IVirtualDesktopManager, IVirtualDesktopManagerInternal, IVirtualDesktopNotificationService,
        IVirtualDesktopPinnedApps,
    },
    Desktop, DesktopID, Error, HRESULT, HWND,
};
use com::{ComInterface, ComRc};
use std::sync::Mutex;

// This is is not thread safe, but it's ok for now
static mut DESKTOPS: Mutex<Vec<Desktop>> = Mutex::new(vec![]);

pub(crate) fn clear_desktops() {
    unsafe {
        DESKTOPS = Mutex::new(vec![]);
    }
}

/// Provides the stateful helper to accessing the Windows 10 Virtual Desktop
/// functions.
pub struct VirtualDesktopService {
    // service_provider: ComRc<dyn IServiceProvider>,
    // sender: Option<VirtualDesktopEventSender>,
    virtual_desktop_manager: ComRc<dyn IVirtualDesktopManager>,
    virtual_desktop_manager_internal: ComRc<dyn IVirtualDesktopManagerInternal>,
    // virtual_desktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService>,
    app_view_collection: ComRc<dyn IApplicationViewCollection>,
    pinned_apps: ComRc<dyn IVirtualDesktopPinnedApps>,
    registered_listener: RegisteredListener,
}

// Let's throw the last of the remaining safety away and implement the send and
// sync 🤞. COM pointers are usually thread safe, but accessing them needs to be
// synced. Access is synced by having Mutex in main Lazy initialization.
unsafe impl Send for VirtualDesktopService {}
unsafe impl Sync for VirtualDesktopService {}

impl VirtualDesktopService {
    /// Initialize only the service, must be-created on TaskbarCreated message
    pub fn create(
        sender: Option<VirtualDesktopEventSender>,
    ) -> Result<Box<VirtualDesktopService>, Error> {
        clear_desktops();

        let service_provider = create_instance::<dyn IServiceProvider>(&CLSID_ImmersiveShell)?;

        let virtual_desktop_manager =
            get_immersive_service::<dyn IVirtualDesktopManager>(&service_provider)?;

        let virtual_desktop_manager_internal = get_immersive_service_for_class(
            &service_provider,
            CLSID_VirtualDesktopManagerInternal,
        )?;

        let app_view_collection = get_immersive_service(&service_provider)?;

        let pinned_apps =
            get_immersive_service_for_class(&service_provider, CLSID_VirtualDesktopPinnedApps)?;

        let virtual_desktop_notification_service: ComRc<dyn IVirtualDesktopNotificationService> =
            get_immersive_service_for_class(&service_provider, CLSID_IVirtualNotificationService)?;

        let registered_listener =
            RegisteredListener::register(sender.clone(), virtual_desktop_notification_service)
                .map_err(Error::ComError)?;

        #[cfg(feature = "debug")]
        println!(
            "VirtualDesktopService created, thread id is {:?}",
            std::thread::current().id()
        );

        Ok(Box::new(VirtualDesktopService {
            // service_provider,
            // virtual_desktop_notification_service,
            // sender,
            virtual_desktop_manager,
            virtual_desktop_manager_internal,
            app_view_collection,
            pinned_apps,
            registered_listener,
        }))
    }

    /// Get raw desktop list
    fn _get_idesktops(&self) -> Result<Vec<ComRc<dyn IVirtualDesktop>>, Error> {
        let mut ptr = None;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .get_desktops(0, &mut ptr)
        })?;
        match ptr {
            Some(objectarray) => {
                // println!("objectarray {:?}", &objectarray as *const _);

                let mut count = 0;
                Result::from(unsafe { objectarray.get_count(&mut count) })?;
                let mut desktops: Vec<ComRc<dyn IVirtualDesktop>> = vec![];

                for i in 0..count {
                    let mut ptr = std::ptr::null_mut();

                    Result::from(unsafe {
                        objectarray.get_at(i, &IVirtualDesktop::IID, &mut ptr)
                    })?;
                    let desktop = unsafe { ComRc::from_raw(ptr as *mut _) };
                    desktops.push(desktop.clone());
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

    fn _get_iapplication_id_for_hwnd(
        &self,
        hwnd: HWND,
    ) -> Result<*mut *mut std::ffi::c_void, Error> {
        let view = self._get_iapplication_view_for_hwnd(hwnd)?;

        // TODO: We probably should convert this to string or slice, so that
        // it's released normally
        let mut app_id: *mut *mut std::ffi::c_void = std::ptr::null_mut();
        Result::from(unsafe { view.get_app_user_model_id(&mut app_id as *mut _ as *mut _) })?;
        Ok(app_id)
    }

    pub fn recreate(&self) -> Result<Box<VirtualDesktopService>, Error> {
        #[cfg(feature = "debug")]
        crate::log_output(&format!("Recreate service"));

        let sender = self.registered_listener.get_sender().clone();
        VirtualDesktopService::create(sender)
    }

    pub fn set_event_sender(&self, sender: VirtualDesktopEventSender) {
        self.registered_listener.set_sender(Some(sender));
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
    pub fn set_desktop_name(&self, desktop: &Desktop, name: &str) -> Result<(), Error> {
        let idesktop = self._get_idesktop_by_id(&desktop.id)?;
        let hstring = HSTRING::create(name).map_err(HRESULT::from_i32)?;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .set_name(idesktop, hstring)
        })?;
        Ok(())
    }

    /// Get desktop name
    pub fn get_desktop_name(&self, desktop: &Desktop) -> Result<String, Error> {
        let idesktop = self._get_idesktop_by_id(&desktop.id)?;
        let mut name = HSTRING::create("                         ").unwrap();
        Result::from(unsafe { idesktop.get_name(&mut name) })?;
        Ok(name.get().unwrap_or_default())
    }

    /// Get desktop IDs
    pub fn get_desktops(&self) -> Result<Vec<Desktop>, Error> {
        // This is cached, as rapid access in `test_threading_two` test causes occasional crashes
        let mut _desktops = unsafe { DESKTOPS.lock().unwrap() };
        if _desktops.is_empty() {
            let desks: Result<Vec<Desktop>, Error> = self
                ._get_idesktops()?
                .iter()
                .map(|f| {
                    let mut desktop = Desktop::empty();
                    Result::from(unsafe { f.get_id(&mut desktop.id) }).map(|_| desktop)
                })
                .collect();
            *_desktops = desks?;
        }
        Ok(_desktops.clone())
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

    pub fn get_desktop_by_guid(&self, desktop_id: &DesktopID) -> Result<Desktop, Error> {
        let mut ptr: Option<ComRc<dyn IVirtualDesktop>> = None;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .find_desktop(desktop_id, &mut ptr)
        })?;
        match ptr {
            Some(idesktop) => {
                let mut desktop = Desktop::empty();
                Result::from(unsafe { idesktop.get_id(&mut desktop.id) }).map(|_| desktop)
            }
            None => Err(Error::ComAllocatedNullPtr),
        }
    }

    /// Get current desktop ID
    pub fn get_current_desktop(&self) -> Result<Desktop, Error> {
        let mut ptr: Option<ComRc<dyn IVirtualDesktop>> = None;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .get_current_desktop(0, &mut ptr)
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
                .switch_desktop(0, idesktop)
        })
    }

    pub fn create_desktop(&self) -> Result<Desktop, Error> {
        let mut idesk_opt: Option<ComRc<dyn IVirtualDesktop>> = None;
        Result::from(unsafe {
            self.virtual_desktop_manager_internal
                .create_desktop(0, &mut idesk_opt)
        })?;
        let idesk = idesk_opt.ok_or(Error::CreateDesktopFailed)?;
        let mut new_desk = Desktop::empty();
        Result::from(unsafe { idesk.get_id(&mut new_desk.id) })?;
        clear_desktops();
        if new_desk.id == DesktopID::default() {
            Err(Error::CreateDesktopFailed)
        } else {
            Ok(new_desk)
        }
    }

    pub fn remove_desktop(
        &self,
        remove_desktop: &Desktop,
        fallback_desktop: &Desktop,
    ) -> Result<(), Error> {
        let a = self._get_idesktop_by_id(&remove_desktop.id)?;
        let b = self._get_idesktop_by_id(&fallback_desktop.id)?;
        clear_desktops();
        Result::from(unsafe { self.virtual_desktop_manager_internal.remove_desktop(a, b) })
    }

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

    /// Is pinned app
    pub fn is_pinned_app(&self, hwnd: HWND) -> Result<bool, Error> {
        let app_id = self._get_iapplication_id_for_hwnd(hwnd)?;
        let mut is_it = false;
        Result::from(unsafe { self.pinned_apps.is_app_pinned(app_id as *mut _, &mut is_it) })?;
        Ok(is_it)
    }

    /// Pin app
    pub fn pin_app(&self, hwnd: HWND) -> Result<(), Error> {
        let app_id = self._get_iapplication_id_for_hwnd(hwnd)?;
        Result::from(unsafe { self.pinned_apps.pin_app(app_id as *mut _) })
    }

    /// Unpin app
    pub fn unpin_app(&self, hwnd: HWND) -> Result<(), Error> {
        let app_id = self._get_iapplication_id_for_hwnd(hwnd)?;
        Result::from(unsafe { self.pinned_apps.unpin_app(app_id as *mut _) })
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

/*
#[cfg(test)]
mod tests {

    use super::*;
    use com::runtime::init_runtime;

    #[test]
    fn test_init() {
        init_runtime().unwrap();
        VirtualDesktopService::create().unwrap();
    }
}
*/
