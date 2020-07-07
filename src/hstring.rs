// Dependency free HSTRING implementation

use std::ffi::OsStr;
use std::ffi::OsString;
use std::os::windows::ffi::OsStrExt;
use std::{os::windows::ffi::OsStringExt, ptr::NonNull};

type LPCWSTR = *const u16;
type HRESULT = i32;

#[derive(PartialEq, PartialOrd, Clone, Debug)]
#[repr(transparent)]
pub struct HSTRING(NonNull<i32>);

impl HSTRING {
    pub fn create(s: &str) -> Result<HSTRING, HRESULT> {
        let utf16bytes: Vec<u16> = OsStr::new(s)
            .encode_wide()
            // Null termination
            .chain(Some(0).into_iter())
            .collect();
        let lpwstr = utf16bytes.as_ptr();
        let mut hstring: Option<HSTRING> = None;

        // Length minus the zero terminator
        let length = utf16bytes.len() - 1;

        let res = unsafe { WindowsCreateString(lpwstr, length, &mut hstring) };
        if res < 0 {
            Err(res)
        } else if let Some(hstr) = hstring {
            Ok(hstr)
        } else {
            Err(1)
        }
    }

    #[allow(dead_code)]
    pub fn get(self) -> Option<String> {
        let mut len: usize = 0;

        let str = unsafe { WindowsGetStringRawBuffer(self, &mut len) };
        let strr = unsafe { std::slice::from_raw_parts(str, len) };
        let f = OsString::from_wide(strr);
        if let Ok(s) = f.into_string() {
            Some(s)
        } else {
            None
        }
    }
}

impl Drop for HSTRING {
    fn drop(&mut self) {
        let _res = unsafe { WindowsDeleteString(self.clone()) };

        #[cfg(feature = "debug")]
        if _res < 0 {
            panic!()
        }
    }
}

#[link(name = "MinCore")]
extern "system" {
    pub fn WindowsCreateString(
        sourceString: LPCWSTR,
        length: usize,
        string: &mut Option<HSTRING>,
    ) -> HRESULT;
    pub fn WindowsDeleteString(hstring: HSTRING) -> HRESULT;
    pub fn WindowsGetStringRawBuffer(hstring: HSTRING, len: *mut usize) -> LPCWSTR;
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hstring_creation() {
        HSTRING::create("Hello World!").unwrap();
    }

    #[test]
    fn test_hstring_get() {
        let hstr = HSTRING::create("Hello World!").unwrap();
        assert_eq!(hstr.get(), Some("Hello World!".to_string()));
    }

    #[test]
    fn test_hstring_drop() {
        let hstr = HSTRING::create("Hello World!").unwrap();
        drop(hstr);
    }
}
