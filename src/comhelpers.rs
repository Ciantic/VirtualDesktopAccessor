use crate::HRESULT;
use com::{sys::CoCreateInstance, ComInterface, ComPtr, ComRc, CLSID, IID};
use std::ffi::c_void;

const CLSCTX_LOCAL_SERVER: u32 = 0x4;

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

#[derive(Debug, PartialEq, Clone)]
pub enum ComError {
    ClassNotRegistered,
    NotInitialized,
    RpcUnavailable,
    ObjectNotConnected,
    Unknown(HRESULT),
}

impl From<HRESULT> for ComError {
    fn from(hr: HRESULT) -> Self {
        match hr {
            HRESULT(0x80040154) => ComError::ClassNotRegistered,
            HRESULT(0x800401F0) => ComError::NotInitialized,
            HRESULT(0x800706BA) => ComError::RpcUnavailable,
            HRESULT(0x800401FD) => ComError::ObjectNotConnected,
            _ => ComError::Unknown(hr),
        }
    }
}
