#![allow(non_snake_case, clippy::not_unsafe_ptr_arg_deref)]

use once_cell::sync::Lazy;
use std::{
    collections::HashSet,
    ffi::{CStr, CString},
    sync::{Arc, Mutex},
};
use windows::{
    core::GUID,
    Win32::{
        Foundation::{HWND, LPARAM, WPARAM},
        UI::WindowsAndMessaging::PostMessageW,
    },
};
use winvd::*;

#[no_mangle]
pub extern "C" fn GetCurrentDesktopNumber() -> i32 {
    get_current_desktop().map_or(-1, |x| x.get_index().map_or(-1, |x| x as i32))
}

// #[no_mangle]
// pub extern "C" fn GetDesktopNumber() -> i32 {
//     get_current_desktop_index_OLDD().map_or(-1, |x| x as i32)
// }

#[no_mangle]
pub extern "C" fn GetDesktopCount() -> i32 {
    get_desktop_count().map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn GetDesktopIdByNumber(number: i32) -> GUID {
    if number < 0 {
        return GUID::default();
    }
    get_desktop(number).get_id().map_or(GUID::default(), |x| x)
}

#[no_mangle]
pub extern "C" fn GetDesktopNumberById(desktop_id: GUID) -> i32 {
    get_desktop(&desktop_id)
        .get_index()
        .map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn GetWindowDesktopId(hwnd: HWND) -> GUID {
    get_desktop_by_window(hwnd).map_or(GUID::default(), |x| {
        x.get_id().map_or(GUID::default(), |y| y)
    })
}

#[no_mangle]
pub extern "C" fn GetWindowDesktopNumber(hwnd: HWND) -> i32 {
    get_desktop_by_window(hwnd).map_or(-1, |x| x.get_index().map_or(-1, |y| y as i32))
}

#[no_mangle]
pub extern "C" fn IsWindowOnCurrentVirtualDesktop(hwnd: HWND) -> i32 {
    is_window_on_current_desktop(hwnd).map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn MoveWindowToDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32 {
    move_window_to_desktop(desktop_number as u32, &hwnd).map_or(-1, |_| 1)
}

#[no_mangle]
pub extern "C" fn GoToDesktopNumber(desktop_number: i32) -> i32 {
    switch_desktop(desktop_number as u32).map_or(-1, |_| 1)
}

#[no_mangle]
pub extern "C" fn SetDesktopName(desktop_number: i32, in_name_ptr: *const i8) -> i32 {
    let name_str = unsafe { CStr::from_ptr(in_name_ptr).to_string_lossy() };
    get_desktop(desktop_number)
        .set_name(&name_str)
        .map_or(-1, |_| 1)
}

#[no_mangle]
pub extern "C" fn GetDesktopName(
    desktop_number: i32,
    out_utf8_ptr: *mut u8,
    out_utf8_len: usize,
) -> i32 {
    if let Ok(name) = get_desktop(desktop_number).get_name() {
        let name_str = CString::new(name).unwrap();
        let name_bytes = name_str.as_bytes_with_nul();
        if name_bytes.len() > out_utf8_len {
            return -1;
        }
        unsafe {
            out_utf8_ptr.copy_from(name_bytes.as_ptr(), name_bytes.len());
        }
        1
    } else {
        0
    }
}

static LISTENER_HWNDS: Lazy<Arc<Mutex<HashSet<isize>>>> =
    Lazy::new(|| Arc::new(Mutex::new(HashSet::new())));

static SENDER_THREAD: Lazy<Arc<Mutex<Option<(DesktopEventThread, std::thread::JoinHandle<()>)>>>> =
    Lazy::new(|| Arc::new(Mutex::new(None)));

#[no_mangle]
pub extern "C" fn RegisterPostMessageHook(listener_hwnd: HWND, message_offset: u32) -> i32 {
    {
        let mut a = LISTENER_HWNDS.lock().unwrap();
        a.insert(listener_hwnd.0);
    }
    {
        let mut a = SENDER_THREAD.lock().unwrap();
        let (tx, rx) = crossbeam_channel::unbounded::<DesktopEvent>();
        if a.is_none() {
            log::log_output("RegisterPostMessageHook: create new threads");
            let listener_thread = std::thread::spawn(move || {
                for item in rx {
                    match item {
                        DesktopEvent::DesktopChanged { new, old } => {
                            let new_index = new.get_index().unwrap_or(0);
                            let old_index = old.get_index().unwrap_or(0);
                            let a = LISTENER_HWNDS.lock().unwrap();
                            for hwnd in a.iter() {
                                unsafe {
                                    let _ = PostMessageW(
                                        HWND(*hwnd as isize),
                                        message_offset,
                                        WPARAM(old_index as usize),
                                        LPARAM(new_index as isize),
                                    );
                                }
                            }
                        }
                        _ => (),
                    }
                }
            });
            let create_sender_result = listen_desktop_events(tx);
            match create_sender_result {
                Ok(sender_thread) => {
                    *a = Some((sender_thread, listener_thread));
                    return 1;
                }
                Err(_er) => {
                    #[cfg(debug_assertions)]
                    log::log_output(&format!("RegisterPostMessageHook failed: {:?}", _er));
                    return -1;
                }
            }
        }
        return 1;
    }
}

#[no_mangle]
pub extern "C" fn UnregisterPostMessageHook(listener_hwnd: HWND) {
    let mut a = LISTENER_HWNDS.lock().unwrap();
    a.remove(&listener_hwnd.0);
    if a.len() == 0 {
        let mut a = SENDER_THREAD.lock().unwrap();
        if let Some((mut sender_thread, listener_thread)) = a.take() {
            // By joining sender thread first it ensures the listener thread finishes when joined
            sender_thread.stop().unwrap();
            listener_thread.join().unwrap();
        }
    }
}
#[no_mangle]
pub extern "C" fn IsPinnedWindow(hwnd: HWND) -> i32 {
    is_pinned_window(hwnd).map_or(-1, |x| x as i32)
}
#[no_mangle]
pub extern "C" fn PinWindow(hwnd: HWND) -> i32 {
    pin_window(hwnd).map_or(-1, |_| 1)
}
#[no_mangle]
pub extern "C" fn UnPinWindow(hwnd: HWND) -> i32 {
    unpin_window(hwnd).map_or(-1, |_| 1)
}
#[no_mangle]
pub extern "C" fn IsPinnedApp(hwnd: HWND) -> i32 {
    is_pinned_app(hwnd).map_or(-1, |x| x as i32)
}
#[no_mangle]
pub extern "C" fn PinApp(hwnd: HWND) -> i32 {
    pin_app(hwnd).map_or(-1, |_| 1)
}
#[no_mangle]
pub extern "C" fn UnPinApp(hwnd: HWND) -> i32 {
    unpin_app(hwnd).map_or(-1, |_| 1)
}
#[no_mangle]
pub extern "C" fn IsWindowOnDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32 {
    is_window_on_desktop(desktop_number, hwnd).map_or(-1, |b| b as i32)
}

#[no_mangle]
pub extern "C" fn CreateDesktop() -> i32 {
    if let Ok(desk) = create_desktop() {
        desk.get_index().map_or(-1, |x| x as i32)
    } else {
        -1
    }
}

#[no_mangle]
pub extern "C" fn RemoveDesktop(remove_desktop_number: i32, fallback_desktop_number: i32) -> i32 {
    if remove_desktop_number == fallback_desktop_number {
        return -1;
    }
    remove_desktop(remove_desktop_number, fallback_desktop_number).map_or(-1, |_| 1)
}

#[no_mangle]
pub extern "C" fn RestartVirtualDesktopAccessor() {
    // ?
}

mod log {
    #[cfg(debug_assertions)]
    extern "system" {
        fn OutputDebugStringW(lpOutputString: windows::core::PCWSTR);
    }

    #[cfg(debug_assertions)]
    pub(crate) fn log_output(s: &str) {
        unsafe {
            println!("{}", s);
            let notepad = format!("{}\0", s).encode_utf16().collect::<Vec<_>>();
            let pw = windows::core::PCWSTR::from_raw(notepad.as_ptr());
            OutputDebugStringW(pw);
        }
    }

    #[cfg(not(debug_assertions))]
    #[inline]
    pub(crate) fn log_output(_s: &str) {}
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_dll_get_desktop_name() {
        // Allocate a buffer for the UTF-8 string
        let utf8_buffer_1024 = [0u8; 1024];
        let utf8_buffer_1024_ptr = utf8_buffer_1024.as_ptr() as *mut u8;
        let res = GetDesktopName(0, utf8_buffer_1024_ptr, 1024);

        let name_cstr = unsafe { std::ffi::CStr::from_ptr(utf8_buffer_1024_ptr as *const i8) };
        let name_str = name_cstr.to_str().unwrap();

        assert_eq!(res, 1);
        assert_eq!(name_str, "Oma");
    }
    #[test]
    fn test_dll_set_desktop_name() {
        let current_desktop_name = get_desktop(0).get_name().unwrap();
        let name = "Testi 😉";
        assert_ne!(current_desktop_name, name);

        let name_cstr = std::ffi::CString::new(name).unwrap();
        let res = SetDesktopName(0, name_cstr.as_ptr() as *mut i8);
        let new_name = get_desktop(0).get_name().unwrap();
        get_desktop(0).set_name(&current_desktop_name).unwrap();
        assert_eq!(new_name, name);
        assert_eq!(res, 1);
    }

    #[test]
    fn test_create_desktop() {
        // Creation works
        let count = GetDesktopCount();
        let new_desk_index = CreateDesktop();
        let new_count = GetDesktopCount();
        assert_eq!(count + 1, new_count);

        // Removing works
        let did_it_work = RemoveDesktop(new_desk_index, 0);
        assert_eq!(did_it_work, 1);
        let after_count = GetDesktopCount();
        assert_eq!(count, after_count);
    }
}
