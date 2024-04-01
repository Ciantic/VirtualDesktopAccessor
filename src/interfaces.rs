//! Interface definitions for the Virtual Desktop API
//!
//! Most of the functions are not tested or used, beware if you try to use these
//! for something else. Notably I know that most out parameters defined as `*mut
//! IMyObject` are incorrect, they probably should be *mut Option<IMyObject>.
//!
//! Generally these are the rules:
//! 1. InOpt = `Option<ComIn<IMyObject>>` or `Option<ManuallyDrop<IMyObject>>`
//! 2. In = `ComIn<IMyObject>` or `ManuallyDrop<IMyObject>`
//! 3. Out = `*mut Option<IMyObject>`
//! 4. OutOpt = `*mut Option<IMyObject>`
//!
//! Last two are same intentionally.
//!
//! ## The summary of COM object lifetime rules:
//!
//! > 1. When a COM object is passed from caller to callee as an input parameter
//! >    to a method, the caller is expected to keep a reference on the object
//! >    for the duration of the method call. The callee shouldn't need to call
//! >    `AddRef` or `Release` for the synchronous duration of that method call.
//! >
//! > 2. When a COM object is passed from callee to caller as an out parameter
//! >    from a method the object is provided to the caller with a reference
//! >    already taken and the caller owns the reference. Which is to say, it is
//! >    the caller's responsibility to call `Release` when they're done with
//! >    the object.
//! >
//! > 3. When making a copy of a COM object pointer you need to call `AddRef`
//! >    and `Release`. The `AddRef` must be called before you call `Release` on
//! >    the original COM object pointer.
//!
//! Rules as [written by David
//! Risney](https://github.com/MicrosoftEdge/WebView2Feedback/issues/2133).
//!
//! If you read the rules carefully, ComIn is most common usecase in Rust
//! API definitions as most parameters are `In` parameters.
#![allow(non_upper_case_globals, clippy::upper_case_acronyms)]

use std::ffi::c_void;
use std::ops::Deref;
use windows::{
    core::{ComInterface, IUnknown, IUnknown_Vtbl, GUID, HRESULT, HSTRING},
    Win32::{Foundation::HWND, UI::Shell::Common::IObjectArray},
};

// Different versions of Windows have slightly different interfaces:
mod build_10240;
mod build_22000;
mod build_dyn;

// We only consume the COM interfaces in a way where we don't depend on the
// exact Windows version:
pub use build_dyn::*;

/// ComIn is a wrapper for COM objects that are passed as input parameters. It
/// allows to keep the life of the COM object for the duration of the function
/// call.
///
/// Imagine following situation:
///
/// First you call an API function that gives COM object as out parameter. And
/// you want to pass it to another function that takes the COM object as an
/// input parameter. If you were to use ManuallyDrop then you'd have to call the
/// drop manually after the second function call.
///
/// E.g.
///
/// ```rust
/// fn get_current_desktop(&mut self, desktop: &mut Option<IVirtualDesktop>) -> HRESULT;
/// fn switch_desktop(&self, desktop: ManuallyDrop<IVirtualDesktop>) -> HRESULT;
///
/// let mut desktop: Option<IVirtualDesktop> = None;
/// get_current_desktop(&mut desktop);
/// if let Some(desktop) = desktop {
///     let input = ManuallyDrop::new(desktop);
///     switch_desktop(input);
///     ManuallyDrop::drop(input);
/// }
/// ```
///
/// To make things safer and easier to use, ComIn is used instead.
///
/// ```rust
/// fn get_current_desktop(&mut self, desktop: &mut Option<IVirtualDesktop>) -> HRESULT;
/// fn switch_desktop(&self, desktop: ComIn<IVirtualDesktop>) -> HRESULT;
///
/// let mut desktop: Option<IVirtualDesktop> = None;
/// if let Some(desktop) = desktop {
///     get_current_desktop(&mut desktop);
///     switch_desktop(ComIn::new(&input));
/// }
/// ```
#[repr(transparent)]
pub struct ComIn<'a, T: ComInterface> {
    data: *mut c_void,
    _phantom: std::marker::PhantomData<&'a T>,
}

impl<'a, T: ComInterface> ComIn<'a, T> {
    pub fn new(t: &'a T) -> Self {
        Self {
            // Copies the raw Inteface pointer
            data: t.as_raw(),
            _phantom: std::marker::PhantomData,
        }
    }
}

impl<'a, T: ComInterface> Deref for ComIn<'a, T> {
    type Target = T;
    fn deref(&self) -> &Self::Target {
        // Safety: A ComInterface type `T` is just a transparent type over a raw pointer
        unsafe { &*(&self.data as *const *mut c_void as *const T) }
    }
}

#[allow(non_upper_case_globals)]
pub const CLSID_ImmersiveShell: GUID = GUID::from_u128(0xC2F03A33_21F5_47FA_B4BB_156362A2F239);

#[allow(dead_code)]
#[allow(non_upper_case_globals)]
pub const CLSID_IVirtualNotificationService: GUID =
    GUID::from_u128(0xA501FDEC_4A09_464C_AE4E_1B9C21B84918);

#[allow(non_upper_case_globals)]
pub const CLSID_VirtualDesktopManagerInternal: GUID =
    GUID::from_u128(0xC5E0CDCA_7B6E_41B2_9FC4_D93975CC467B);

#[allow(non_upper_case_globals)]
pub const CLSID_VirtualDesktopPinnedApps: GUID =
    GUID::from_u128(0xb5a399e7_1c87_46b8_88e9_fc5747b171bd);

type BOOL = i32;
type DWORD = u32;
type INT = i32;
type LPVOID = *mut c_void;
type UINT = u32;
type ULONG = u32;
type WCHAR = u16;
type PCWSTR = *const WCHAR;
type PWSTR = *mut WCHAR;
type ULONGLONG = u64;
type LONG = i32;

type IAsyncCallback = UINT;
type IImmersiveMonitor = UINT;
type IApplicationViewOperation = UINT;
type IApplicationViewPosition = UINT;
type IImmersiveApplication = UINT;
type IApplicationViewChangeListener = UINT;
#[allow(non_camel_case_types)]
type APPLICATION_VIEW_COMPATIBILITY_POLICY = UINT;
#[allow(non_camel_case_types)]
type APPLICATION_VIEW_CLOAK_TYPE = UINT;

#[allow(dead_code)]
#[repr(C)]
pub struct RECT {
    left: LONG,
    top: LONG,
    right: LONG,
    bottom: LONG,
}

#[allow(dead_code)]
#[repr(C)]
pub struct SIZE {
    cx: LONG,
    cy: LONG,
}

// These COM interfaces are not different between different Windows versions:

#[windows_interface::interface("6D5140C1-7436-11CE-8034-00AA006009FA")]
pub unsafe trait IServiceProvider: IUnknown {
    pub unsafe fn query_service(
        &self,
        guid_service: *const GUID,
        riid: *const GUID,
        ppv_object: *mut *mut c_void,
    ) -> HRESULT;
    // unsafe fn remote_query_service(
    //     &self,
    //     guidService: *const DesktopID,
    //     riid: *const IID,
    //     ppvObject: *mut *mut c_void,
    // ) -> HRESULT;
}

#[windows_interface::interface("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
pub unsafe trait IVirtualDesktopManager: IUnknown {
    pub unsafe fn is_window_on_current_desktop(
        &self,
        top_level_window: HWND,
        out_on_current_desktop: *mut bool,
    ) -> HRESULT;
    pub unsafe fn get_desktop_by_window(
        &self,
        top_level_window: HWND,
        out_desktop_id: *mut GUID,
    ) -> HRESULT;
    pub unsafe fn move_window_to_desktop(
        &self,
        top_level_window: HWND,
        desktop_id: *const GUID,
    ) -> HRESULT;
}
