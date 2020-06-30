use crate::interfaces::IServiceProvider;
use com::{
    sys::{CoCreateInstance, FAILED, GUID, HRESULT},
    ComInterface, ComPtr, ComRc, CLSID, IID,
};
use std::ffi::c_void;
use winapi::shared::wtypesbase::CLSCTX_LOCAL_SERVER;

pub fn get_immersive_service<T: ComInterface + ?Sized>(
    service_provider: &ComRc<dyn IServiceProvider>,
) -> Result<ComRc<T>, HRESULT> {
    get_immersive_service_for_class::<T>(service_provider, T::IID)
}

pub fn get_immersive_service_for_class<T: ComInterface + ?Sized>(
    service_provider: &ComRc<dyn IServiceProvider>,
    class_id: GUID,
) -> Result<ComRc<T>, HRESULT> {
    let mut obj = std::ptr::null_mut::<c_void>();
    let res = unsafe { (*service_provider).query_service(&class_id, &T::IID, &mut obj) };

    if FAILED(res) {
        return Err(res);
    }

    unsafe { Ok(ComRc::new(ComPtr::new(obj as *mut _))) }
}

struct ImmersiveService {
    provider: ComRc<dyn IServiceProvider>,
}

impl ImmersiveService {
    pub fn get_virtual_desktop_manager() -> bool {
        true
    }

    pub fn get_virtual_desktop_notification_service() -> bool {
        true
    }
}
