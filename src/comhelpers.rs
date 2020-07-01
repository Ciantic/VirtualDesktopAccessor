use crate::HRESULT;
use com::{sys::CoCreateInstance, ComInterface, ComPtr, ComRc, CLSID, IID};
use std::ffi::c_void;
use winapi::shared::wtypesbase::CLSCTX_LOCAL_SERVER;

// https://github.com/microsoft/com-rs/issues/150

/// Create an instance of a CoClass with the associated class id
///
/// Calls `CoCreateInstance` internally
pub fn create_instance<T: ComInterface + ?Sized>(class_id: &CLSID) -> Result<ComRc<T>, HRESULT> {
    unsafe {
        Ok(ComRc::new(create_raw_instance::<T>(
            class_id,
            std::ptr::null_mut(),
        )?))
    }
}

/// A helper  for creating both regular and aggregated instances
pub unsafe fn create_raw_instance<T: ComInterface + ?Sized>(
    class_id: &CLSID,
    outer: *mut c_void,
) -> Result<ComPtr<T>, HRESULT> {
    let mut instance = std::ptr::null_mut::<c_void>();
    let hr = HRESULT::from_i32(CoCreateInstance(
        class_id as *const CLSID,
        outer,
        CLSCTX_LOCAL_SERVER,
        &T::IID as *const IID,
        &mut instance as *mut *mut c_void,
    ));
    if hr.failed() {
        return Err(hr);
    }

    Ok(ComPtr::new(instance as *mut _))
}
