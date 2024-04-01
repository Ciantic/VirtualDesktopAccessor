//! Interface abstraction that allow interacting with COM interfaces even when
//! we don't know their IID and what exact methods they have until runtime.

use super::*;

use core::{ffi::c_void, marker::PhantomData};
use windows::{
    core::{ComInterface, Interface, GUID, HRESULT, HSTRING},
    Win32::{
        Foundation::{E_NOTIMPL, HWND},
        UI::Shell::Common::IObjectArray,
    },
};

/// Indicates different Windows versions that have different Virtual Desktop
/// interfaces.
#[derive(Debug, PartialEq, Eq, PartialOrd, Ord, Clone, Copy, Default)]
pub enum WindowsVersion {
    Build10240,
    #[default]
    Build22000,
}
impl WindowsVersion {
    // Aliases used by macros (should match the module names above):
    const build_10240: Self = Self::Build10240;
    const build_22000: Self = Self::Build22000;

    /// Get info about the current Windows version. Only differentiates between
    /// Windows versions that have different virtual desktop interfaces.
    ///
    /// # Determining Windows Version
    ///
    /// We could use the [`GetVersionExW` function
    /// (sysinfoapi.h)](https://learn.microsoft.com/en-us/windows/win32/api/sysinfoapi/nf-sysinfoapi-getversionexw),
    /// but it is deprecated after Windows 8.1. It also changes behavior depending
    /// on what manifest is embedded in the executable.
    ///
    /// That pages links to [Version Helper functions - Win32
    /// apps](https://learn.microsoft.com/en-us/windows/win32/sysinfo/version-helper-apis)
    /// where we are linked to the [`IsWindowsVersionOrGreater` function
    /// (versionhelpers.h)](https://learn.microsoft.com/en-us/windows/win32/api/VersionHelpers/nf-versionhelpers-iswindowsversionorgreater)
    /// and the [`VerifyVersionInfoA` function
    /// (winbase.h)](https://learn.microsoft.com/en-us/windows/win32/api/Winbase/nf-winbase-verifyversioninfoa)
    /// that it uses internally (though the later function is deprecated in Windows
    /// 10).
    ///
    /// We can use `RtlGetVersion` [RtlGetVersion function (wdm.h) - Windows
    /// drivers](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/wdm/nf-wdm-rtlgetversion?redirectedfrom=MSDN)
    /// as mentioned at [c++ - Detecting Windows 10 version - Stack
    /// Overflow](https://stackoverflow.com/questions/36543301/detecting-windows-10-version/36545162#36545162).
    ///
    /// # `windows` API References
    ///
    /// - [GetVersionExW in windows::Win32::System::SystemInformation -
    ///   Rust](https://microsoft.github.io/windows-docs-rs/doc/windows/Win32/System/SystemInformation/fn.GetVersionExW.html)
    ///   - Affected by manifest.
    /// - [RtlGetVersion in windows::Wdk::System::SystemServices -
    ///   Rust](https://microsoft.github.io/windows-docs-rs/doc/windows/Wdk/System/SystemServices/fn.RtlGetVersion.html)
    ///   - Always returns the correct version.
    pub fn get() -> Self {
        static INIT: std::sync::OnceLock<WindowsVersion> = std::sync::OnceLock::new();
        *INIT.get_or_init(|| {
            let mut version: windows::Win32::System::SystemInformation::OSVERSIONINFOW =
                Default::default();
            version.dwOSVersionInfoSize = core::mem::size_of_val(&version) as u32;
            let res = unsafe { windows::Wdk::System::SystemServices::RtlGetVersion(&mut version) };
            if res.is_err() {
                return Default::default();
            }
            if version.dwBuildNumber < 22000 {
                WindowsVersion::Build10240
            } else {
                WindowsVersion::Build22000
            }
        })
    }
}

/// Do an action with the type of the actual COM Interface on this Windows
/// version.
///
/// This is implemented for the more generic COM interfaces that don't know
/// their IID at compile time. Those implementations will put a bound on `F` so
/// that it must accept all concrete COM types that might be used.
pub trait WithVersionedType<F, R> {
    /// Invokes the callback with the COM interface type of this Windows
    /// version. Return `None` if the interface doesn't exist on this platform.
    fn with_versioned_type(callback: F) -> Option<R>;
}
/// A callback that will be invoked with the actual COM interface type for a
/// specific Windows version.
///
/// Implement this when you want to make use of a concrete COM interface type.
pub trait WithVersionedTypeCallback<T: ComInterface, R> {
    fn call(self) -> R;
}

/// Generates support code for a COM interface.
///
/// Syntax is: `as CreatedEnumName for InterfaceName in $(all)? [version_module_1, version_module_2 $(,)?]`
macro_rules! support_interface {
    (@inner {$dollar:tt} as $state:ident for $name:ident in $(all $(@ $all:tt)?)? [$($version:ident),* $(,)?]) => {
        $(
            // assert_eq_size from static_assertions crate
            const _: fn() = || {
                // We need this since we transmute and pointer cast between the two types.
                let _ = core::mem::transmute::<$name, self::$version::$name>;
            };
        )*

        // Maybe enforce that all build versions are supported by this interface:
        #[allow(unreachable_patterns)]
        const _: fn(WindowsVersion) = |version: WindowsVersion| {
            match version {
                $(WindowsVersion::$version => (),)*
                // If there is no "all" word in the macro input then:
                //   all() => true
                // Otherwise:
                //   all(false) => false
                //   any() => false
                //   all(any()) => false
                // And the default macro arm will be hidden.
                #[cfg(all($(any() $($all)?)?))]
                _ => (),
            }
        };

        /// An enum with one variant per Windows version that is supported by
        /// this interface. Use the `from_typed` function to construct this
        /// type.
        #[allow(non_camel_case_types)]
        enum $state<'a> {
            $( $version(ComIn<'a, self::$version::$name>) ),*
        }
        impl<'a> $state<'a> {
            fn from_typed(data: &'a $name) -> Self {
                unsafe { Self::from_raw(&data.0) }
            }
            /// # Safety
            ///
            /// The COM object must implement the expected interface.
            #[allow(unreachable_patterns)]
            unsafe fn from_raw(data: &'a IUnknown) -> Self {
                let win_ver = WindowsVersion::get();
                match win_ver {
                    $(WindowsVersion::$version => $state::$version(core::mem::transmute_copy::<IUnknown, ComIn<'_, _>>(data)),)*
                    _ => unreachable!("Tried to cast into a COM interface that wasn't available for the current Windows version"),
                }
            }
        }
        impl $name {
            /// Convert from a raw pointer to the COM interface.
            ///
            /// # Safety
            ///
            /// The pointer must be an instance of the COM interface indicated
            /// by the `IID` method.
            pub unsafe fn from_raw(ptr: *mut c_void) -> Self {
                Self(IUnknown::from_raw(ptr))
            }
            /// The IID for the COM interface that is supported by this
            /// platform, return a zeroed GUID if the interface isn't supported.
            #[allow(non_snake_case, unreachable_patterns)]
            pub fn IID() -> GUID {
                match WindowsVersion::get() {
                    $(WindowsVersion::$version => self::$version::$name::IID,)*
                    _ => GUID::zeroed(),
                }
            }
        }
        /// Allow direct access to the wrapped COM interface type if required.
        impl<F, R> WithVersionedType<F, R> for $name
        where
            $(
                F: WithVersionedTypeCallback<self::$version::$name, R>,
            )*
        {
            #[allow(unreachable_patterns)]
            fn with_versioned_type(callback: F) -> Option<R> {
                match WindowsVersion::get() {
                    $(WindowsVersion::$version => Some(<F as WithVersionedTypeCallback<self::$version::$name, R>>::call(callback)),)*
                    _ => None,
                }
            }
        }
        $(
            /// Version specific -> Generic interface
            impl From<self::$version::$name> for $name {
                fn from(v: self::$version::$name) -> Self {
                    debug_assert_eq!(
                        WindowsVersion::get(),
                        WindowsVersion::$version,
                        "if we have an COM interface for a specific Windows version then we must already have ensured that it is actually the Windows version the user has"
                    );
                    Self(v.into())
                }
            }
            /// Reference to version specific -> Reference to generic interface
            impl<'a> From<&'a self::$version::$name> for &'a $name {
                fn from(v: &'a self::$version::$name) -> Self {
                    debug_assert_eq!(
                        WindowsVersion::get(),
                        WindowsVersion::$version,
                        "if we have an COM interface for a specific Windows version then we must already have ensured that it is actually the Windows version the user has"
                    );
                    // Safety: both types are just transparent wrappers over a
                    // raw pointer and we don't drop either of them.
                    unsafe {
                        &*(v as *const self::$version::$name as *const $name)
                    }
                }
            }
            /// Fallible conversion from generic interface to version specific
            /// interface.
            impl From<$name> for self::$version::$name {
                fn from(v: $name) -> Self {
                    assert_eq!(WindowsVersion::get(), WindowsVersion::$version);
                    // Safety: interpret the wrapped raw pointer as the specific COM interface.
                    unsafe { core::mem::transmute(v.0) }
                }
            }
            impl<'a> From<&'a $name> for ComIn<'a, self::$version::$name> {
                #[allow(irrefutable_let_patterns)]
                fn from(v: &'a $name) -> Self {
                    if let $state::$version(v) = $state::from_typed(v) {
                        v
                    } else {
                        unreachable!("requested a COM interface for a different Windows version than the one that was installed");
                    }
                }
            }
        )*
        /// Preform the same action for each version of the wrapped COM interface.
        ///
        /// Syntax: GeneralType, |versioned: versioned_mod::VersionedType| block_of_code
        ///
        /// Note: named the same as the interface to allow for easier usage with macros.
        #[allow(unused_macros)]
        macro_rules! $name {
            (
                $dollar this:expr,
                |$dollar arg:ident
                    $dollar (
                        :
                        $dollar module_name:ident
                        ::
                        $dollar arg_ty:ident
                    )?
                |
                $dollar ($dollar body:tt)*
            ) => {
                match $state::from_typed(&$dollar this) {
                    $(
                        $state::$version($dollar arg) => {
                            $dollar (
                                #[allow(unused_imports)]
                                use self::$version as $dollar module_name;
                                #[allow(unused_imports)]
                                use self::$version::$name as $dollar arg_ty;
                            )?
                            $dollar ($dollar body)*
                        },
                    )*
                }
            }
        }
    };
    // Pass an escaped dollar sign to the real macro so that we can construct a
    // new macro later:
    (as $state:ident for $name:ident in $($in:tt)*) => {
        support_interface! { @inner {$} as $state for $name in $($in)* }
    };
}

/// Implement a method by calling the same method on the Windows version
/// dependant COM interface.
macro_rules! forward_call {
    (
        #[forward_for = $name:ident]
        $( #[$attr:meta] )*
        $pub:vis
        $(unsafe $(@ $unsafe:tt)?)?
        fn $fname:ident (
            &$self_:ident $(,)? $( $arg_name:ident : $ArgTy:ty ),* $(,)?
        ) -> $RetTy:ty;
    ) => (
        $( #[$attr] )*
        #[allow(unused_parens)]
        $pub
        $(unsafe $($unsafe)?)?
        fn $fname (
            &$self_, $( $arg_name : $ArgTy ),*
        ) -> $RetTy
        {
            unsafe {
                $name!(
                    $self_,
                    |v| v.$fname( $(
                        Into::into($arg_name)
                    ),*)
                )
            }
        }
    );
    (
        $( #[$attr:meta] )*
        impl $name:ident {
            $($item:item)*
        }
    ) => {
        $(#[$attr])*
        impl $name {
            $(
                #[apply(forward_call)]
                #[forward_for = $name]
                $item
            )*
        }
    };
}

support_interface!(as IApplicationViewInner for IApplicationView in all [build_10240, build_22000]);

#[derive(Debug, Clone, PartialEq, Eq)]
#[repr(transparent)]
pub struct IApplicationView(IUnknown);
impl IApplicationView {
    /* IInspecateble */
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_iids(
        &self,
        out_iid_count: *mut ULONG,
        out_opt_iid_array_ptr: *mut *mut GUID,
    ) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_runtime_class_name(&self, out_opt_class_name: *mut HSTRING) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_trust_level(&self, ptr_trust_level: LPVOID) -> HRESULT;

    /* IApplicationView methods */
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub fn set_focus(&self) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub fn switch_to(&self) -> HRESULT;

    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn try_invoke_back(&self, ptr_async_callback: IAsyncCallback) -> HRESULT;
    pub fn get_thumbnail_window(&self, out_hwnd: &mut HWND) -> HRESULT {
        unsafe { IApplicationView!(self, |i| i.get_thumbnail_window(out_hwnd)) }
    }
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_monitor(&self, out_monitors: *mut *mut IImmersiveMonitor) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_visibility(&self, out_int: LPVOID) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_cloak(
        &self,
        application_view_cloak_type: APPLICATION_VIEW_CLOAK_TYPE,
        unknown: INT,
    ) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_position(
        &self,
        unknowniid: *const GUID,
        unknown_array_ptr: LPVOID,
    ) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_position(&self, view_position: *mut IApplicationViewPosition) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn insert_after_window(&self, window: HWND) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_extended_frame_position(&self, rect: *mut RECT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_app_user_model_id(&self, id: *mut PWSTR) -> HRESULT; // Proc17
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_app_user_model_id(&self, id: PCWSTR) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn is_equal_by_app_user_model_id(&self, id: PCWSTR, out_result: *mut INT)
        -> HRESULT;

    /*** IApplicationView methods ***/
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_view_state(&self, out_state: *mut UINT) -> HRESULT; // Proc20
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_view_state(&self, state: UINT) -> HRESULT; // Proc21
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_neediness(&self, out_neediness: *mut INT) -> HRESULT; // Proc22
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_last_activation_timestamp(&self, out_timestamp: *mut ULONGLONG) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_last_activation_timestamp(&self, timestamp: ULONGLONG) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_virtual_desktop_id(&self, out_desktop_guid: *mut GUID) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_virtual_desktop_id(&self, desktop_guid: *const GUID) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_show_in_switchers(&self, out_show: *mut INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_show_in_switchers(&self, show: INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_scale_factor(&self, out_scale_factor: *mut INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn can_receive_input(&self, out_can: *mut BOOL) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_compatibility_policy_type(
        &self,
        out_policy_type: *mut APPLICATION_VIEW_COMPATIBILITY_POLICY,
    ) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_compatibility_policy_type(
        &self,
        policy_type: APPLICATION_VIEW_COMPATIBILITY_POLICY,
    ) -> HRESULT;

    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_size_constraints(
        &self,
        monitor: *mut IImmersiveMonitor,
        out_size1: *mut SIZE,
        out_size2: *mut SIZE,
    ) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_size_constraints_for_dpi(
        &self,
        dpi: UINT,
        out_size1: *mut SIZE,
        out_size2: *mut SIZE,
    ) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn set_size_constraints_for_dpi(
        &self,
        dpi: *const UINT,
        size1: *const SIZE,
        size2: *const SIZE,
    ) -> HRESULT;

    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn on_min_size_preferences_updated(&self, window: HWND) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn apply_operation(&self, operation: *mut IApplicationViewOperation) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn is_tray(&self, out_is: *mut BOOL) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn is_in_high_zorder_band(&self, out_is: *mut BOOL) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn is_splash_screen_presented(&self, out_is: *mut BOOL) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn flash(&self) -> HRESULT;
    pub unsafe fn get_root_switchable_owner(&self, app_view: *mut IApplicationView) -> HRESULT {
        // proc45
        IApplicationView!(self, |inner| inner
            .get_root_switchable_owner(app_view as *mut _))
    }

    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn enumerate_ownership_tree(&self, objects: *mut IObjectArray) -> HRESULT; // proc46

    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn get_enterprise_id(&self, out_id: *mut PWSTR) -> HRESULT; // proc47
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn is_mirrored(&self, out_is: *mut BOOL) -> HRESULT; //

    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown1(&self, arg: *mut INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown2(&self, arg: *mut INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown3(&self, arg: *mut INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown4(&self, arg: INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown5(&self, arg: *mut INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown6(&self, arg: INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown7(&self) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown8(&self, arg: *mut INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown9(&self, arg: INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown10(&self, arg: INT, arg2: INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown11(&self, arg: INT) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IApplicationView]
    pub unsafe fn unknown12(&self, arg: *mut SIZE) -> HRESULT;
}

support_interface!(as IVirtualDesktopInner for IVirtualDesktop in [build_10240, build_22000]);

#[derive(Debug, Clone, PartialEq, Eq)]
#[repr(transparent)]
pub struct IVirtualDesktop(IUnknown);
impl IVirtualDesktop {
    pub fn is_view_visible(&self, p_view: &IApplicationView, out_bool: &mut u32) -> HRESULT {
        unsafe { IVirtualDesktop!(self, |inner| inner.is_view_visible(p_view.into(), out_bool)) }
    }
    pub fn get_id(&self, out_guid: &mut GUID) -> HRESULT {
        unsafe { IVirtualDesktop!(self, |inner| inner.get_id(out_guid)) }
    }

    pub fn get_name(&self, out_string: &mut HSTRING) -> HRESULT {
        match IVirtualDesktopInner::from_typed(self) {
            IVirtualDesktopInner::build_10240(_) => E_NOTIMPL,
            IVirtualDesktopInner::build_22000(this) => unsafe { this.get_name(out_string) },
        }
    }
    pub fn get_wallpaper(&self, out_string: &mut HSTRING) -> HRESULT {
        match IVirtualDesktopInner::from_typed(self) {
            IVirtualDesktopInner::build_10240(_) => E_NOTIMPL,
            IVirtualDesktopInner::build_22000(this) => unsafe { this.get_wallpaper(out_string) },
        }
    }
}

support_interface!(as IApplicationViewCollectionInner for IApplicationViewCollection in all [build_10240, build_22000]);

#[derive(Debug, Clone, PartialEq, Eq)]
#[repr(transparent)]
pub struct IApplicationViewCollection(IUnknown);
impl IApplicationViewCollection {
    #[apply(forward_call)]
    #[forward_for = IApplicationViewCollection]
    pub unsafe fn get_views(&self, out_views: *mut IObjectArray) -> HRESULT;

    #[apply(forward_call)]
    #[forward_for = IApplicationViewCollection]
    pub unsafe fn get_views_by_zorder(&self, out_views: *mut IObjectArray) -> HRESULT;

    #[apply(forward_call)]
    #[forward_for = IApplicationViewCollection]
    pub unsafe fn get_views_by_app_user_model_id(
        &self,
        id: PCWSTR,
        out_views: *mut IObjectArray,
    ) -> HRESULT;

    pub unsafe fn get_view_for_hwnd(
        &self,
        window: HWND,
        out_view: *mut Option<IApplicationView>,
    ) -> HRESULT {
        IApplicationViewCollection!(self, |inner| inner
            .get_view_for_hwnd(window, out_view as *mut _))
    }

    pub unsafe fn get_view_for_application(
        &self,
        app: IImmersiveApplication,
        out_view: *mut IApplicationView,
    ) -> HRESULT {
        IApplicationViewCollection!(self, |inner| inner
            .get_view_for_application(app, out_view as *mut _))
    }

    pub unsafe fn get_view_for_app_user_model_id(
        &self,
        id: PCWSTR,
        out_view: *mut IApplicationView,
    ) -> HRESULT {
        IApplicationViewCollection!(self, |inner| inner
            .get_view_for_app_user_model_id(id, out_view as *mut _))
    }

    pub fn get_view_in_focus(&self, out_view: &mut Option<IApplicationView>) -> HRESULT {
        unsafe {
            IApplicationViewCollection!(self, |inner| inner
                .get_view_in_focus(out_view as *mut Option<_> as *mut _))
        }
    }

    pub fn try_get_last_active_visible_view(
        &self,
        out_view: &mut Option<IApplicationView>,
    ) -> HRESULT {
        unsafe {
            match IApplicationViewCollectionInner::from_typed(self) {
                IApplicationViewCollectionInner::build_10240(_) => E_NOTIMPL,
                IApplicationViewCollectionInner::build_22000(this) => {
                    this.try_get_last_active_visible_view(out_view as *mut Option<_> as *mut _)
                }
            }
        }
    }

    #[apply(forward_call)]
    #[forward_for = IApplicationViewCollection]
    pub unsafe fn refresh_collection(&self) -> HRESULT;

    pub fn register_for_application_view_changes(
        &self,
        listener: IApplicationViewChangeListener,
        out_id: &mut DWORD,
    ) -> HRESULT {
        unsafe {
            IApplicationViewCollection!(self, |inner| inner
                .register_for_application_view_changes(listener, out_id))
        }
    }

    #[apply(forward_call)]
    #[forward_for = IApplicationViewCollection]
    pub fn unregister_for_application_view_changes(&self, id: DWORD) -> HRESULT;
}

support_interface!(as IVirtualDesktopNotificationInner for IVirtualDesktopNotification in all [build_10240, build_22000]);

#[derive(Debug, Clone, PartialEq, Eq)]
#[repr(transparent)]
pub struct IVirtualDesktopNotification(IUnknown);
impl IVirtualDesktopNotification {
    pub fn as_raw(&self) -> *mut c_void {
        self.0.as_raw()
    }
}
impl<T> From<T> for IVirtualDesktopNotification
where
    T: IVirtualDesktopNotification_Impl,
{
    fn from(value: T) -> Self {
        match WindowsVersion::get() {
            WindowsVersion::Build10240 => build_10240::IVirtualDesktopNotification::from(
                build_10240::VirtualDesktopNotificationAdaptor { inner: value },
            )
            .into(),
            WindowsVersion::Build22000 => build_22000::IVirtualDesktopNotification::from(
                build_22000::VirtualDesktopNotificationAdaptor { inner: value },
            )
            .into(),
        }
    }
}
#[allow(non_camel_case_types)]
pub trait IVirtualDesktopNotification_Impl {
    fn virtual_desktop_created(&self, desktop: &IVirtualDesktop) -> HRESULT;

    fn virtual_desktop_destroy_begin(
        &self,
        desktop_destroyed: &IVirtualDesktop,
        desktop_fallback: &IVirtualDesktop,
    ) -> HRESULT;

    fn virtual_desktop_destroy_failed(
        &self,
        desktop_destroyed: &IVirtualDesktop,
        desktop_fallback: &IVirtualDesktop,
    ) -> HRESULT;

    fn virtual_desktop_destroyed(
        &self,
        desktop_destroyed: &IVirtualDesktop,
        desktop_fallback: &IVirtualDesktop,
    ) -> HRESULT;

    fn virtual_desktop_moved(
        &self,
        desktop: &IVirtualDesktop,
        old_index: i64,
        new_index: i64,
    ) -> HRESULT;

    fn virtual_desktop_name_changed(&self, desktop: &IVirtualDesktop, name: HSTRING) -> HRESULT;

    fn view_virtual_desktop_changed(&self, view: &IApplicationView) -> HRESULT;

    fn current_virtual_desktop_changed(
        &self,
        desktop_old: &IVirtualDesktop,
        desktop_new: &IVirtualDesktop,
    ) -> HRESULT;

    fn virtual_desktop_wallpaper_changed(
        &self,
        desktop: &IVirtualDesktop,
        name: HSTRING,
    ) -> HRESULT;

    fn virtual_desktop_switched(&self, desktop: &IVirtualDesktop) -> HRESULT;

    fn remote_virtual_desktop_connected(&self, desktop: &IVirtualDesktop) -> HRESULT;
}

support_interface!(as IVirtualDesktopNotificationServiceInner for IVirtualDesktopNotificationService in all [build_10240, build_22000]);

#[derive(Debug, Clone, PartialEq, Eq)]
#[repr(transparent)]
pub struct IVirtualDesktopNotificationService(IUnknown);

impl IVirtualDesktopNotificationService {
    pub fn register(
        &self,
        notification: &IVirtualDesktopNotification,
        out_cookie: &mut DWORD,
    ) -> HRESULT {
        unsafe {
            IVirtualDesktopNotificationService!(self, |inner| inner
                .register(notification.as_raw(), out_cookie))
        }
    }

    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopNotificationService]
    pub fn unregister(&self, cookie: u32) -> HRESULT;
}

support_interface!(as IVirtualDesktopManagerInternalInner for IVirtualDesktopManagerInternal in all [build_10240, build_22000]);

#[derive(Debug, Clone, PartialEq, Eq)]
#[repr(transparent)]
pub struct IVirtualDesktopManagerInternal(IUnknown);
impl IVirtualDesktopManagerInternal {
    pub fn get_desktop_count(&self, out_count: &mut UINT) -> HRESULT {
        unsafe { IVirtualDesktopManagerInternal!(self, |i| i.get_desktop_count(out_count)) }
    }

    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopManagerInternal]
    pub fn move_view_to_desktop(
        &self,
        view: &IApplicationView,
        desktop: &IVirtualDesktop,
    ) -> HRESULT;

    pub fn can_move_view_between_desktops(
        &self,
        view: &IApplicationView,
        can_move: &mut i32,
    ) -> HRESULT {
        unsafe {
            IVirtualDesktopManagerInternal!(self, |i| i
                .can_move_view_between_desktops(view.into(), can_move))
        }
    }

    pub fn get_current_desktop(&self, out_desktop: &mut Option<IVirtualDesktop>) -> HRESULT {
        unsafe {
            IVirtualDesktopManagerInternal!(self, |inner| {
                inner.get_current_desktop(out_desktop as *mut Option<_> as *mut Option<_>)
            })
        }
    }

    pub fn get_desktops(&self, out_desktops: &mut Option<IObjectArray>) -> HRESULT {
        unsafe { IVirtualDesktopManagerInternal!(self, |inner| inner.get_desktops(out_desktops)) }
    }

    /// Get next or previous desktop
    ///
    /// Direction values:
    /// 3 = Left direction
    /// 4 = Right direction
    pub fn get_adjacent_desktop(
        &self,
        in_desktop: &IVirtualDesktop,
        direction: UINT,
        out_pp_desktop: &mut Option<IVirtualDesktop>,
    ) -> HRESULT {
        unsafe {
            IVirtualDesktopManagerInternal!(self, |inner| inner.get_adjacent_desktop(
                in_desktop.into(),
                direction,
                out_pp_desktop as *mut Option<_> as *mut Option<_>,
            ))
        }
    }

    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopManagerInternal]
    pub fn switch_desktop(&self, desktop: &IVirtualDesktop) -> HRESULT;

    pub fn create_desktop(&self, out_desktop: &mut Option<IVirtualDesktop>) -> HRESULT {
        unsafe {
            IVirtualDesktopManagerInternal!(self, |inner| inner
                .create_desktop(out_desktop as *mut Option<_> as *mut Option<_>))
        }
    }

    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopManagerInternal]
    pub fn move_desktop(&self, in_desktop: &IVirtualDesktop, index: UINT) -> HRESULT;

    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopManagerInternal]
    pub fn remove_desktop(
        &self,
        destroy_desktop: &IVirtualDesktop,
        fallback_desktop: &IVirtualDesktop,
    ) -> HRESULT;

    pub fn find_desktop(&self, guid: &GUID, out_desktop: &mut Option<IVirtualDesktop>) -> HRESULT {
        unsafe {
            IVirtualDesktopManagerInternal!(self, |inner| {
                inner.find_desktop(guid, out_desktop as *mut Option<_> as *mut Option<_>)
            })
        }
    }

    pub fn get_desktop_switch_include_exclude_views(
        &self,
        desktop: &IVirtualDesktop,
        out_pp_desktops1: &mut Option<IObjectArray>,
        out_pp_desktops2: &mut Option<IObjectArray>,
    ) -> HRESULT {
        unsafe {
            IVirtualDesktopManagerInternal!(self, |inner| inner
                .get_desktop_switch_include_exclude_views(
                    desktop.into(),
                    out_pp_desktops1 as *mut Option<_> as *mut _,
                    out_pp_desktops2 as *mut Option<_> as *mut _
                ))
        }
    }

    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopManagerInternal]
    pub fn set_name(&self, desktop: &IVirtualDesktop, name: HSTRING) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopManagerInternal]
    pub fn set_wallpaper(&self, desktop: &IVirtualDesktop, name: HSTRING) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopManagerInternal]
    pub fn update_wallpaper_for_all(&self, name: HSTRING) -> HRESULT;
}

support_interface!(as IVirtualDesktopPinnedAppsInner for IVirtualDesktopPinnedApps in all [build_10240, build_22000]);

#[derive(Debug, Clone, PartialEq, Eq)]
#[repr(transparent)]
pub struct IVirtualDesktopPinnedApps(IUnknown);

impl IVirtualDesktopPinnedApps {
    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopPinnedApps]
    pub unsafe fn is_app_pinned(&self, app_id: PCWSTR, out_iss: *mut bool) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopPinnedApps]
    pub unsafe fn pin_app(&self, app_id: PCWSTR) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopPinnedApps]
    pub unsafe fn unpin_app(&self, app_id: PCWSTR) -> HRESULT;

    pub fn is_view_pinned(&self, view: &IApplicationView, out_iss: &mut bool) -> HRESULT {
        unsafe {
            IVirtualDesktopPinnedApps!(self, |inner| inner.is_view_pinned(view.into(), out_iss))
        }
    }
    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopPinnedApps]
    pub fn pin_view(&self, view: &IApplicationView) -> HRESULT;
    #[apply(forward_call)]
    #[forward_for = IVirtualDesktopPinnedApps]
    pub fn unpin_view(&self, view: &IApplicationView) -> HRESULT;
}

// Bellow are helper methods that accesses the real COM interfaces. We could
// avoid the need for these helper methods by working with the IUnknown
// interface or implementing ComInterface for our abstraction types with the
// `IUnknown` IDD, even if that might get confusing.

struct IObjectArrayGetAtCallback<'a, T>(&'a IObjectArray, UINT, PhantomData<T>);
impl<COM, T> WithVersionedTypeCallback<COM, Result<T, windows::core::Error>>
    for IObjectArrayGetAtCallback<'_, T>
where
    // The COM interface for this specific Windows version:
    COM: ComInterface,
    // Should be possible to convert it into the more generic type:
    T: From<COM>,
{
    fn call(self) -> Result<T, windows::core::Error> {
        let com: COM = unsafe { self.0.GetAt::<COM>(self.1)? };
        Ok(From::from(com))
    }
}

/// Same as `GetAt` for `IObjectArray` but works even when we don't know the IID
/// of a COM interface at compile time.
#[allow(non_snake_case, private_bounds)]
pub unsafe fn IObjectArrayGetAt<'a, T>(
    object_array: &'a IObjectArray,
    index: UINT,
) -> Result<T, windows::core::Error>
where
    T: WithVersionedType<IObjectArrayGetAtCallback<'a, T>, Result<T, windows::core::Error>>,
{
    T::with_versioned_type(IObjectArrayGetAtCallback(object_array, index, PhantomData))
        .ok_or_else(|| windows::core::Error::from(E_NOTIMPL))?
}
