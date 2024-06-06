//! This is a version of the `interface` module that works for multiple
//! different Windows version by using different COM interfaces depending on the
//! version that the program is running on.
//!
//! See the docs for the normal [`crate::interfaces`] module for more info about
//! the bindings themselves.
//!
//! We manage Virtual Desktops using unstable COM interfaces which change their
//! ids and definitions regularly.
//!
//! # References
//!
//! - Interface ids at [VirtualDesktop/src/VirtualDesktop/app.config at 7e37b9848aef681713224dae558d2e51960cf41e · mzomparelli/VirtualDesktop](https://github.com/mzomparelli/VirtualDesktop/blob/7e37b9848aef681713224dae558d2e51960cf41e/src/VirtualDesktop/app.config)
//! - Bindings at [VirtualDesktop/src/VirtualDesktop/Interop at 7e37b9848aef681713224dae558d2e51960cf41e · mzomparelli/VirtualDesktop](https://github.com/mzomparelli/VirtualDesktop/tree/7e37b9848aef681713224dae558d2e51960cf41e/src/VirtualDesktop/Interop)
//!   - These are actually compiled when the app is executed by the `ComInterfaceAssemblyBuilder.CreateAssembly` method at: [VirtualDesktop/src/VirtualDesktop/Interop/ComInterfaceAssemblyBuilder.cs at 7e37b9848aef681713224dae558d2e51960cf41e · mzomparelli/VirtualDesktop](https://github.com/mzomparelli/VirtualDesktop/blob/7e37b9848aef681713224dae558d2e51960cf41e/src/VirtualDesktop/Interop/ComInterfaceAssemblyBuilder.cs#L84-L153)
//! - Bindings at [MScholtes/VirtualDesktop at 6de804dced760778450ae3cd1481f8969f75fb39](https://github.com/MScholtes/VirtualDesktop/tree/6de804dced760778450ae3cd1481f8969f75fb39)

#![allow(non_upper_case_globals, clippy::upper_case_acronyms)]

use std::ffi::c_void;
use std::ops::Deref;
use windows::{
    core::{Interface, IUnknown, IUnknown_Vtbl, GUID, HRESULT, HSTRING},
    Win32::{Foundation::HWND, UI::Shell::Common::IObjectArray},
};

/// This macro allows us to access the version names in other macros.
macro_rules! declare_versions {
    (@inner {
        dollar = {$dollar:tt},
        versions = {$($version:tt),* $(,)?},
    }) => {
        macro_rules! _with_versions {
            ($dollar macro_callback:path $dollar (, $dollar ( $dollar state:tt )* )?) => {
                $dollar macro_callback! {
                    $dollar ($dollar ( $dollar state )* )?
                    versions = {$($version,)*},
                }
            };
        }
        #[allow(unused_imports)]
        use _with_versions as with_versions;
    };
    ($(
        $(#[ $($attr:tt)* ])*
        mod $version:tt $({ $($code:tt)* })? $(; $(<= $semi:tt)?)?
    )*) => {
        $(
            $(#[ $($attr)* ])*
            mod $version $({ $($code)* })? $(; $($semi)?)?
        )*
        declare_versions! {@inner {
            dollar = {$},
            versions = {$($version,)*},
        }}
    };
}
declare_versions!(
    mod build_10240;
    mod build_16299; // IDD change
    mod build_17134; // IDD change
    mod build_19045; // IDD change
    mod build_20348; // Interface change
    mod build_22000; // Interface change
    mod build_22621_2215; // Interface change
    mod build_22621_3155; // IID change
    mod build_22631_2428; // IID change
    mod build_22631_3155; // IID change
);
mod build_dyn;

// We only consume the COM interfaces in a way where we don't depend on the
// exact Windows version:
pub use build_dyn::*;

/// Use when defining a COM interface to allow re-defining it later with a
/// different IID.
///
/// This defines a macro with the same name as the interface that can be invoked
/// with a string literal representing a new IID.
macro_rules! _reusable_com_interface {
    (@inner {
        dollar = {$dollar:tt},
        interface = {$($interface:tt)*},
        export_as = {$name:ident},
        temp_macro_name = {$temp_name:ident},
        initial_iid = {$initial_iid:literal},
    }) => {
        // Allow re-use of the declaration with different COM interface iid:
        macro_rules! $temp_name {
            ($dollar iid:literal) => {
                $crate::interfaces_multi::reusable_com_interface!(
                    MacroOptions {
                        temp_macro_name: $temp_name,
                        iid: $dollar iid,
                    },
                    {
                        $($interface)*
                    }
                );
            }
        }
        #[allow(unused_imports)]
        pub(crate) use $temp_name as $name;

        // Declare the COM interface at least once:
        #[windows_interface::interface($initial_iid)]
        $($interface)*
    };
    // This syntax was chosen so that rustfmt would still work on macro context
    // (rustfmt only formats macros invoked using parenthesis not other
    // brackets.)
    (MacroOptions {
        temp_macro_name: $temp_name:ident,
        iid: $initial_iid:literal $(,)?
    }, {
        $(#[$($attr:tt)*])*
        $pub:vis $(unsafe $(@ $unsafe:ident)?)? trait $name:ident $($interface:tt)*
    }) => {
        $crate::interfaces_multi::reusable_com_interface! {@inner {
            dollar = {$},
            interface = {
                $(#[$($attr)*])*
                $pub $(unsafe $(@ $unsafe)?)? trait $name
                $($interface)*
            },
            export_as = {$name},
            temp_macro_name = {$temp_name},
            initial_iid = {$initial_iid},
        }}
    };
}
// Allow normal imports to work for macro:
use _reusable_com_interface as reusable_com_interface;

/// Type that can be cast into [`ComIn`]
///
/// # Safety
///
/// - Can cast from `*mut c_void` to `Self`. (`Self` is a transparent type over
///   a raw pointer.)
/// - The returned pointer is valid while the reference it was created from is
///   valid.
pub unsafe trait PointerRepr {
    fn as_pointer_repr(&self) -> *mut c_void;
}
unsafe impl<T: Interface> PointerRepr for T {
    fn as_pointer_repr(&self) -> *mut c_void {
        windows::core::Interface::as_raw(self)
    }
}

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
/// ```rust,ignore
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
/// ```rust,ignore
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
pub struct ComIn<'a, T> {
    data: *mut c_void,
    _phantom: std::marker::PhantomData<&'a T>,
}
impl<'a, T: PointerRepr> ComIn<'a, T> {
    pub fn new(t: &'a T) -> Self {
        Self {
            // Copies the raw Inteface pointer
            data: t.as_pointer_repr(),
            _phantom: std::marker::PhantomData,
        }
    }
}
impl<'a, T> ComIn<'a, T> {
    pub fn into_ref(this: &Self) -> &'a T {
        // Safety: A ComInterface type `T` is just a transparent type over a raw pointer
        unsafe { &*(&this.data as *const *mut c_void as *const T) }
    }
}
impl<'a, T> Deref for ComIn<'a, T> {
    type Target = T;
    fn deref(&self) -> &Self::Target {
        Self::into_ref(self)
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
type HMONITOR = isize;

type IAsyncCallback = UINT;
type IImmersiveMonitor = UINT;
type IApplicationViewOperation = UINT;
type IApplicationViewPosition = UINT;
type IShellPositionerPriority = *mut c_void;
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
