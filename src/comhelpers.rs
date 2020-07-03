use crate::HRESULT;
use com::{
    runtime::{init_apartment, ApartmentType},
    sys::CoCreateInstance,
    ComInterface, ComPtr, ComRc, CLSID, IID,
};
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

fn errorhandler<T, F>(f: F, error: HRESULT, retry: u32) -> Result<T, HRESULT>
where
    F: Fn() -> Result<T, HRESULT>,
{
    if retry == 0 {
        Err(error)
    } else {
        match error {
            // Com class is not registered
            HRESULT(0x80040154) => comrun(f, retry),

            // Com not initialized
            HRESULT(0x800401F0) => {
                init_apartment(ApartmentType::Multithreaded).map_err(HRESULT::from_i32)?;
                comrun(f, retry)
            }

            // RPC went away
            HRESULT(0x800706BA) => comrun(f, retry),

            // Others as is
            HRESULT(v) => Err(HRESULT(v)),
        }
    }
}

fn comrun<T, F>(f: F, retry: u32) -> Result<T, HRESULT>
where
    F: Fn() -> Result<T, HRESULT>,
{
    match f() {
        Ok(v) => Ok(v),
        Err(err) => errorhandler(f, err, retry - 1),
    }
}

pub fn run<F, T>(f: F) -> Result<T, HRESULT>
where
    F: Fn() -> Result<T, HRESULT>,
{
    comrun(f, 6)
}
