use std::{
    collections::HashMap,
    sync::{Arc, Mutex},
    thread,
};

use once_cell::sync::Lazy;
use winapi::{shared::windef::HWND, winrt::hstring::HSTRING};
use winvd::{
    create_event_listener, get_desktop_by_guid, get_desktop_by_index, get_desktop_by_window,
    helpers::*, is_pinned_app, is_pinned_window, is_window_on_current_desktop,
    is_window_on_desktop, pin_app, pin_window, unpin_app, unpin_window, DesktopID,
    VirtualDesktopEvent, VirtualDesktopEventSender,
};

#[no_mangle]
pub extern "C" fn GetCurrentDesktopNumber() -> i32 {
    get_current_desktop_number().map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn GetDesktopCount() -> i32 {
    get_desktop_count().map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn GetDesktopIdByNumber(number: i32) -> DesktopID {
    if number < 0 {
        return DesktopID::default();
    }
    get_desktop_by_index(number as usize).map_or(DesktopID::default(), |desktop| desktop.get_id())
}

#[no_mangle]
pub extern "C" fn GetDesktopNumber() -> i32 {
    get_current_desktop_number().map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn GetDesktopNumberById(desktop_id: DesktopID) -> i32 {
    get_desktop_by_guid(&desktop_id).map_or(-1, |x| x.get_index().map_or(-1, |y| y as i32))
}

#[no_mangle]
pub extern "C" fn GetWindowDesktopId(hwnd: HWND) -> DesktopID {
    get_desktop_by_window(hwnd as u32).map_or(DesktopID::default(), |x| x.get_id())
}

#[no_mangle]
pub extern "C" fn GetWindowDesktopNumber(hwnd: HWND) -> i32 {
    get_desktop_by_window(hwnd as u32).map_or(-1, |x| x.get_index().map_or(-1, |y| y as i32))
}

#[no_mangle]
pub extern "C" fn IsWindowOnCurrentVirtualDesktop(hwnd: HWND) -> i32 {
    is_window_on_current_desktop(hwnd as u32).map_or(-1, |x| x as i32)
}

#[no_mangle]
pub extern "C" fn MoveWindowToDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32 {
    move_window_to_desktop_number(hwnd as u32, desktop_number as usize).map_or(-1, |_| 1)
}

#[no_mangle]
pub extern "C" fn GoToDesktopNumber(desktop_number: i32) {
    go_to_desktop_number(desktop_number as usize).unwrap_or_default()
}

#[no_mangle]
pub extern "C" fn SetName(desktop_number: i32, name: HSTRING) {
    // TODO:
    // rename_desktop_number(number, name)
}

#[no_mangle]
pub extern "C" fn GetName(desktop_number: i32, name: HSTRING) {
    // TODO:
    // rename_desktop_number(number, name)
}

static LISTENER_HWNDS: Lazy<Arc<Mutex<HashMap<u32, u32>>>> =
    Lazy::new(|| Arc::new(Mutex::new(HashMap::new())));

static LISTENER_THREAD: Lazy<Arc<Mutex<Option<std::thread::JoinHandle<()>>>>> =
    Lazy::new(|| Arc::new(Mutex::new(None)));

// static LISTENER_SENDER: Lazy<Arc<Mutex<Option<VirtualDesktopEventSender>>>> =
//     Lazy::new(|| Arc::new(Mutex::new(None)));

#[no_mangle]
pub extern "C" fn RegisterPostMessageHook(listener_hwnd: HWND, message_offset: u32) {
    let mut a = LISTENER_HWNDS.lock().unwrap();
    if a.len() == 0 {
        let mut static_thread = LISTENER_THREAD.lock().unwrap();
        let thread_handle = thread::spawn(|| {
            let (sender, receiver) = std::sync::mpsc::channel();
            create_event_listener(VirtualDesktopEventSender::Std(sender)).unwrap();
            receiver.iter().for_each(|msg| match msg {
                VirtualDesktopEvent::DesktopChanged(_old, new) => {
                    let hwnds = LISTENER_HWNDS.lock();
                    if let Ok(hwnds) = hwnds {
                        for (hwnd, offset) in hwnds.iter() {
                            unsafe {
                                winapi::um::winuser::PostMessageW(
                                    *hwnd as _,
                                    *offset,
                                    0,
                                    new.get_index().map_or(-1, |x| x as isize),
                                );
                            }
                        }
                    }
                }
                VirtualDesktopEvent::DesktopCreated(_desk) => {
                    // println!("<- New desktop created {:?}", desk);
                }
                VirtualDesktopEvent::DesktopDestroyed(_desk) => {
                    // println!("<- Desktop destroyed {:?}", desk);
                }
                VirtualDesktopEvent::WindowChanged(_hwnd) => {
                    // println!("<- Window changed {:?}", hwnd);
                }
                VirtualDesktopEvent::DesktopNameChanged(_desk, _name) => {
                    // println!("<- Name of {:?} changed to {}", desk, name);
                }
                VirtualDesktopEvent::DesktopWallpaperChanged(_desk, _name) => {
                    // println!("<- Wallpaper of {:?} changed to {}", desk, name);
                }
                VirtualDesktopEvent::DesktopMoved(_desk, _old, _new) => {
                    // println!("<- Desktop {:?} moved from {} to {}", desk, old, new);
                }
            });
        });
        static_thread.replace(thread_handle);
    }
    a.insert(listener_hwnd as u32, message_offset);
}

#[no_mangle]
pub extern "C" fn UnregisterPostMessageHook(listener_hwnd: HWND) {
    let mut a = LISTENER_HWNDS.lock().unwrap();
    a.remove(&(listener_hwnd as u32));
    if a.len() == 0 {
        // let mut static_thread = LISTENER_THREAD.lock().unwrap();
        // let mut static_sender = LISTENER_SENDER.lock().unwrap();
        // static_sender.take();
        // static_thread.take().unwrap().join().unwrap();
    }
}
#[no_mangle]
pub extern "C" fn IsPinnedWindow(hwnd: HWND) -> i32 {
    is_pinned_window(hwnd as u32).map_or(-1, |x| x as i32)
}
#[no_mangle]
pub extern "C" fn PinWindow(hwnd: HWND) {
    pin_window(hwnd as u32).unwrap_or_default()
}
#[no_mangle]
pub extern "C" fn UnPinWindow(hwnd: HWND) {
    unpin_window(hwnd as u32).unwrap_or_default()
}
#[no_mangle]
pub extern "C" fn IsPinnedApp(hwnd: HWND) -> i32 {
    is_pinned_app(hwnd as u32).map_or(-1, |x| x as i32)
}
#[no_mangle]
pub extern "C" fn PinApp(hwnd: HWND) {
    pin_app(hwnd as u32).unwrap_or_default()
}
#[no_mangle]
pub extern "C" fn UnPinApp(hwnd: HWND) {
    unpin_app(hwnd as u32).unwrap_or_default()
}
#[no_mangle]
pub extern "C" fn IsWindowOnDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32 {
    get_desktop_by_index(desktop_number as usize).map_or(-1, |x| {
        is_window_on_desktop(hwnd as u32, &x).map_or(-1, |b| b as i32)
    })
}
#[no_mangle]
pub extern "C" fn RestartVirtualDesktopAccessor() {
    // ?
}

#[no_mangle]
pub extern "C" fn lib_test() {
    println!("Hello from the library!");
}
/*
* int GetCurrentDesktopNumber()
* int GetDesktopCount()
* GUID GetDesktopIdByNumber(int number) // Returns zeroed GUID with invalid number found
* int GetDesktopNumber(IVirtualDesktop *pDesktop)
* int GetDesktopNumberById(GUID desktopId)
* GUID GetWindowDesktopId(HWND window)
* int GetWindowDesktopNumber(HWND window)
* int IsWindowOnCurrentVirtualDesktop(HWND window)
* BOOL MoveWindowToDesktopNumber(HWND window, int number)
* void GoToDesktopNumber(int number)
* void RegisterPostMessageHook(HWND listener, int messageOffset)
* void UnregisterPostMessageHook(HWND hwnd)
* int IsPinnedWindow(HWND hwnd) // Returns 1 if pinned, 0 if not pinned, -1 if not valid
* void PinWindow(HWND hwnd)
* void UnPinWindow(HWND hwnd)
* int IsPinnedApp(HWND hwnd) // Returns 1 if pinned, 0 if not pinned, -1 if not valid
* void PinApp(HWND hwnd)
* void UnPinApp(HWND hwnd)
* int IsWindowOnDesktopNumber(HWND window, int number) /
* void RestartVirtualDesktopAccessor() // Call this during taskbar created message

*/
