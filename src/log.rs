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
