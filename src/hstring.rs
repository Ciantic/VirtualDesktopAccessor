use crate::HRESULT;
use std::ffi::{c_void, OsStr};

type WCHAR = u16;
type PCWSTR = *const WCHAR;
type LPCWSTR = *const WCHAR;
type PCNZWCH = LPCWSTR;

#[derive(PartialEq, PartialOrd, Clone, Debug)]
pub struct HSTRING(*const u32);

impl Drop for HSTRING {
    fn drop(&mut self) {
        let _res = unsafe { WindowsDeleteString(self) };

        #[cfg(feature = "debug")]
        if _res.failed() {
            panic!()
        }
    }
}

impl HSTRING {
    pub fn create(s: &str) -> Result<HSTRING, HRESULT> {
        /*
        let text: Vec<u16> = OsStr::new(text).encode_wide(). chain(Some(0).into_iter()).collect();
        let lp_wstr = text.as_ptr(); //The LPCWSTR
        */
        let strr: String = s.to_string() + "\0";
        let mut hstr: HSTRING = HSTRING(&0);
        // OsStr::new(s).encode_wide();
        let u16str = strr.encode_utf16().collect::<Vec<_>>();
        let length = u16str.len() as u32 - 1;
        println!("Send hstring {:?}", hstr);
        let res = unsafe { WindowsCreateString(u16str.as_ptr(), length, &mut hstr) };
        if res.failed() {
            Err(res)
        } else {
            println!("Got hstring {:?}", hstr);
            Ok(hstr)
        }
    }

    pub fn get(&self) -> String {
        let mut len = 0;
        let str = unsafe { WindowsGetStringRawBuffer(self, &mut len) };
        // OsStr::from()
        // String::from_utf16(str).unwrap()
        "".to_string()
    }
}

#[link(name = "MinCore")]
extern "system" {
    pub fn WindowsCreateString(sourceString: PCNZWCH, length: u32, string: *mut HSTRING)
        -> HRESULT;
    pub fn WindowsDeleteString(hstring: *const HSTRING) -> HRESULT;
    pub fn WindowsGetStringRawBuffer(hstring: *const HSTRING, len: *const u32) -> PCWSTR;
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hstring_creation() {
        let hstr = HSTRING::create("Foolio test").unwrap();
        assert_eq!(hstr.get(), "Foolio test".to_string());
    }

    #[test]
    fn test_failure() {
        assert_eq!(HRESULT(0x800706BA).failed(), true);
    }
}
