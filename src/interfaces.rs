/// Interface definitions for the Virtual Desktop API
///
/// Most of the functions are not tested or used, beware if you try to use these
/// for something else. Notably I know that most out parameters defined as `*mut
/// IMyObject` are incorrect, they probably should be *mut Option<IMyObject>.
///
/// Generally these are the rules:
/// 1. InOpt = `Option<ComIn<IMyObject>>` or `Option<ManuallyDrop<IMyObject>>`
/// 2. In = `ComIn<IMyObject>` or `ManuallyDrop<IMyObject>`
/// 3. Out = `*mut Option<IMyObject>`
/// 4. OutOpt = `*mut Option<IMyObject>`
///
/// Last two are same intentionally.
///
/// ## The summary of COM object lifetime rules:
///
/// > 1. When a COM object is passed from caller to callee as an input parameter
/// >    to a method, the caller is expected to keep a reference on the object
/// >    for the duration of the method call. The callee shouldn't need to call
/// >    `AddRef` or `Release` for the synchronous duration of that method call.
/// >
/// > 2. When a COM object is passed from callee to caller as an out parameter
/// >    from a method the object is provided to the caller with a reference
/// >    already taken and the caller owns the reference. Which is to say, it is
/// >    the caller's responsibility to call `Release` when they're done with
/// >    the object.
/// >
/// > 3. When making a copy of a COM object pointer you need to call `AddRef`
/// >    and `Release`. The `AddRef` must be called before you call `Release` on
/// >    the original COM object pointer.
///
/// Rules as [written by David
/// Risney](https://github.com/MicrosoftEdge/WebView2Feedback/issues/2133).
///
/// If you read the rules carefully, ManuallyDrop is most common usecase in Rust
/// API definitions as most parameters are `In` parameters.
#[allow(non_upper_case_globals)]
use std::{ffi::c_void, ops::Deref};
use windows::{
    core::{IUnknown, IUnknown_Vtbl, Vtable, GUID, HRESULT, HSTRING},
    Win32::{Foundation::HWND, UI::Shell::Common::IObjectArray},
};

// Behaves like ManuallyDrop but is kept alive for as long as the given
// reference
#[repr(transparent)]
pub struct ComIn<'a, T: Vtable> {
    data: *mut c_void,
    _phantom: std::marker::PhantomData<&'a T>,
}

impl<'a, T: Vtable> ComIn<'a, T> {
    pub fn new(t: &'a T) -> Self {
        Self {
            // Copies the raw Inteface pointer
            data: t.as_raw(),
            _phantom: std::marker::PhantomData,
        }
    }
}

impl<'a, T: Vtable> Deref for ComIn<'a, T> {
    type Target = T;
    fn deref(&self) -> &Self::Target {
        unsafe { std::mem::transmute(&self.data) }
    }
}

#[allow(non_upper_case_globals)]
pub const CLSID_ImmersiveShell: GUID = GUID {
    data1: 0xC2F03A33,
    data2: 0x21F5,
    data3: 0x47FA,
    data4: [0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39],
};

#[allow(dead_code)]
#[allow(non_upper_case_globals)]
pub const CLSID_IVirtualNotificationService: GUID = GUID {
    data1: 0xA501FDEC,
    data2: 0x4A09,
    data3: 0x464C,
    data4: [0xAE, 0x4E, 0x1B, 0x9C, 0x21, 0xB8, 0x49, 0x18],
};

#[allow(non_upper_case_globals)]
pub const CLSID_VirtualDesktopManagerInternal: GUID = GUID {
    data1: 0xC5E0CDCA,
    data2: 0x7B6E,
    data3: 0x41B2,
    data4: [0x9F, 0xC4, 0xD9, 0x39, 0x75, 0xCC, 0x46, 0x7B],
};

#[allow(non_upper_case_globals)]
pub const CLSID_VirtualDesktopPinnedApps: GUID = GUID {
    data1: 0xb5a399e7,
    data2: 0x1c87,
    data3: 0x46b8,
    data4: [0x88, 0xe9, 0xfc, 0x57, 0x47, 0xb1, 0x71, 0xbd],
};
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
type IImmersiveApplication = UINT;
type IApplicationViewChangeListener = UINT;
#[allow(non_camel_case_types)]
type APPLICATION_VIEW_COMPATIBILITY_POLICY = UINT;
#[allow(non_camel_case_types)]
type APPLICATION_VIEW_CLOAK_TYPE = UINT;

#[allow(dead_code)]
pub struct RECT {
    left: LONG,
    top: LONG,
    right: LONG,
    bottom: LONG,
}

#[allow(dead_code)]
pub struct SIZE {
    cx: LONG,
    cy: LONG,
}

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

#[windows_interface::interface("a5cd92ff-29be-454c-8d04-d82879fb3f1b")]
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

#[windows_interface::interface("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
pub unsafe trait IApplicationView: IUnknown {
    /* IInspecateble */
    pub unsafe fn get_iids(
        &self,
        out_iid_count: *mut ULONG,
        out_opt_iid_array_ptr: *mut *mut GUID,
    ) -> HRESULT;
    pub unsafe fn get_runtime_class_name(&self, out_opt_class_name: *mut HSTRING) -> HRESULT;
    pub unsafe fn get_trust_level(&self, ptr_trust_level: LPVOID) -> HRESULT;

    /* IApplicationView methods */
    pub unsafe fn set_focus(&self) -> HRESULT;
    pub unsafe fn switch_to(&self) -> HRESULT;

    pub unsafe fn try_invoke_back(&self, ptr_async_callback: IAsyncCallback) -> HRESULT;
    pub unsafe fn get_thumbnail_window(&self, out_hwnd: *mut HWND) -> HRESULT;
    pub unsafe fn get_monitor(&self, out_monitors: *mut *mut IImmersiveMonitor) -> HRESULT;
    pub unsafe fn get_visibility(&self, out_int: LPVOID) -> HRESULT;
    pub unsafe fn set_cloak(
        &self,
        application_view_cloak_type: APPLICATION_VIEW_CLOAK_TYPE,
        unknown: INT,
    ) -> HRESULT;
    pub unsafe fn get_position(
        &self,
        unknowniid: *const GUID,
        unknown_array_ptr: LPVOID,
    ) -> HRESULT;
    pub unsafe fn set_position(&self, view_position: *mut IApplicationViewPosition) -> HRESULT;
    pub unsafe fn insert_after_window(&self, window: HWND) -> HRESULT;
    pub unsafe fn get_extended_frame_position(&self, rect: *mut RECT) -> HRESULT;
    pub unsafe fn get_app_user_model_id(&self, id: *mut PWSTR) -> HRESULT; // Proc17
    pub unsafe fn set_app_user_model_id(&self, id: PCWSTR) -> HRESULT;
    pub unsafe fn is_equal_by_app_user_model_id(&self, id: PCWSTR, out_result: *mut INT)
        -> HRESULT;

    /*** IApplicationView methods ***/
    pub unsafe fn get_view_state(&self, out_state: *mut UINT) -> HRESULT; // Proc20
    pub unsafe fn set_view_state(&self, state: UINT) -> HRESULT; // Proc21
    pub unsafe fn get_neediness(&self, out_neediness: *mut INT) -> HRESULT; // Proc22
    pub unsafe fn get_last_activation_timestamp(&self, out_timestamp: *mut ULONGLONG) -> HRESULT;
    pub unsafe fn set_last_activation_timestamp(&self, timestamp: ULONGLONG) -> HRESULT;
    pub unsafe fn get_virtual_desktop_id(&self, out_desktop_guid: *mut GUID) -> HRESULT;
    pub unsafe fn set_virtual_desktop_id(&self, desktop_guid: *const GUID) -> HRESULT;
    pub unsafe fn get_show_in_switchers(&self, out_show: *mut INT) -> HRESULT;
    pub unsafe fn set_show_in_switchers(&self, show: INT) -> HRESULT;
    pub unsafe fn get_scale_factor(&self, out_scale_factor: *mut INT) -> HRESULT;
    pub unsafe fn can_receive_input(&self, out_can: *mut BOOL) -> HRESULT;
    pub unsafe fn get_compatibility_policy_type(
        &self,
        out_policy_type: *mut APPLICATION_VIEW_COMPATIBILITY_POLICY,
    ) -> HRESULT;
    pub unsafe fn set_compatibility_policy_type(
        &self,
        policy_type: APPLICATION_VIEW_COMPATIBILITY_POLICY,
    ) -> HRESULT;

    pub unsafe fn get_size_constraints(
        &self,
        monitor: *mut IImmersiveMonitor,
        out_size1: *mut SIZE,
        out_size2: *mut SIZE,
    ) -> HRESULT;
    pub unsafe fn get_size_constraints_for_dpi(
        &self,
        dpi: UINT,
        out_size1: *mut SIZE,
        out_size2: *mut SIZE,
    ) -> HRESULT;
    pub unsafe fn set_size_constraints_for_dpi(
        &self,
        dpi: *const UINT,
        size1: *const SIZE,
        size2: *const SIZE,
    ) -> HRESULT;

    pub unsafe fn on_min_size_preferences_updated(&self, window: HWND) -> HRESULT;
    pub unsafe fn apply_operation(&self, operation: *mut IApplicationViewOperation) -> HRESULT;
    pub unsafe fn is_tray(&self, out_is: *mut BOOL) -> HRESULT;
    pub unsafe fn is_in_high_zorder_band(&self, out_is: *mut BOOL) -> HRESULT;
    pub unsafe fn is_splash_screen_presented(&self, out_is: *mut BOOL) -> HRESULT;
    pub unsafe fn flash(&self) -> HRESULT;
    pub unsafe fn get_root_switchable_owner(&self, app_view: *mut IApplicationView) -> HRESULT; // proc45
    pub unsafe fn enumerate_ownership_tree(&self, objects: *mut IObjectArray) -> HRESULT; // proc46

    pub unsafe fn get_enterprise_id(&self, out_id: *mut PWSTR) -> HRESULT; // proc47
    pub unsafe fn is_mirrored(&self, out_is: *mut BOOL) -> HRESULT; //

    pub unsafe fn unknown1(&self, arg: *mut INT) -> HRESULT;
    pub unsafe fn unknown2(&self, arg: *mut INT) -> HRESULT;
    pub unsafe fn unknown3(&self, arg: *mut INT) -> HRESULT;
    pub unsafe fn unknown4(&self, arg: INT) -> HRESULT;
    pub unsafe fn unknown5(&self, arg: *mut INT) -> HRESULT;
    pub unsafe fn unknown6(&self, arg: INT) -> HRESULT;
    pub unsafe fn unknown7(&self) -> HRESULT;
    pub unsafe fn unknown8(&self, arg: *mut INT) -> HRESULT;
    pub unsafe fn unknown9(&self, arg: INT) -> HRESULT;
    pub unsafe fn unknown10(&self, arg: INT, arg2: INT) -> HRESULT;
    pub unsafe fn unknown11(&self, arg: INT) -> HRESULT;
    pub unsafe fn unknown12(&self, arg: *mut SIZE) -> HRESULT;
}

#[windows_interface::interface("536D3495-B208-4CC9-AE26-DE8111275BF8")]
pub unsafe trait IVirtualDesktop: IUnknown {
    pub unsafe fn is_view_visible(
        &self,
        p_view: ComIn<IApplicationView>,
        out_bool: *mut u32,
    ) -> HRESULT;
    pub unsafe fn get_id(&self, out_guid: *mut GUID) -> HRESULT;
    pub unsafe fn get_monitor(&self, out_monitor: *mut HMONITOR) -> HRESULT;
    pub unsafe fn get_name(&self, out_string: *mut HSTRING) -> HRESULT;
    pub unsafe fn get_wallpaper(&self, out_string: *mut HSTRING) -> HRESULT;
}

#[windows_interface::interface("1841c6d7-4f9d-42c0-af41-8747538f10e5")]
pub unsafe trait IApplicationViewCollection: IUnknown {
    pub unsafe fn get_views(&self, out_views: *mut IObjectArray) -> HRESULT;

    pub unsafe fn get_views_by_zorder(&self, out_views: *mut IObjectArray) -> HRESULT;

    pub unsafe fn get_views_by_app_user_model_id(
        &self,
        id: PCWSTR,
        out_views: *mut IObjectArray,
    ) -> HRESULT;

    pub unsafe fn get_view_for_hwnd(
        &self,
        window: HWND,
        out_view: *mut Option<IApplicationView>,
    ) -> HRESULT;

    pub unsafe fn get_view_for_application(
        &self,
        app: IImmersiveApplication,
        out_view: *mut IApplicationView,
    ) -> HRESULT;

    pub unsafe fn get_view_for_app_user_model_id(
        &self,
        id: PCWSTR,
        out_view: *mut IApplicationView,
    ) -> HRESULT;

    pub unsafe fn get_view_in_focus(&self, out_view: *mut IApplicationView) -> HRESULT;

    pub unsafe fn try_get_last_active_visible_view(
        &self,
        out_view: *mut IApplicationView,
    ) -> HRESULT;

    pub unsafe fn refresh_collection(&self) -> HRESULT;

    pub unsafe fn register_for_application_view_changes(
        &self,
        listener: IApplicationViewChangeListener,
        out_id: *mut DWORD,
    ) -> HRESULT;

    pub unsafe fn unregister_for_application_view_changes(&self, id: DWORD) -> HRESULT;
}

// NOTE: Currently ComIn is basically ManuallyDrop. I've tried without
// ManuallyDrop and the code starts to act weird if we call Release() on the
// values given by the shell to the IVirtualDesktopNotification.
//
// Normally functions should call IUnknown's Release() on the given pointer
// after they are done with it, but the shell doesn't like that for this
// interface.
#[windows_interface::interface("CD403E52-DEED-4C13-B437-B98380F2B1E8")]
pub unsafe trait IVirtualDesktopNotification: IUnknown {
    pub unsafe fn virtual_desktop_created(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_destroy_begin(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_destroy_failed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_destroyed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_destroyed: ComIn<IVirtualDesktop>,
        desktop_fallback: ComIn<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_is_per_monitor_changed(&self, is_per_monitor: i32) -> HRESULT;

    pub unsafe fn virtual_desktop_moved(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop: ComIn<IVirtualDesktop>,
        old_index: i64,
        new_index: i64,
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_name_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT;

    pub unsafe fn view_virtual_desktop_changed(&self, view: ComIn<IApplicationView>) -> HRESULT;

    pub unsafe fn current_virtual_desktop_changed(
        &self,
        monitors: ComIn<IObjectArray>,
        desktop_old: ComIn<IVirtualDesktop>,
        desktop_new: ComIn<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn virtual_desktop_wallpaper_changed(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT;
}

#[windows_interface::interface("0CD45E71-D927-4F15-8B0A-8FEF525337BF")]
pub unsafe trait IVirtualDesktopNotificationService: IUnknown {
    pub unsafe fn register(
        &self,
        notification: *mut std::ffi::c_void, // *const IVirtualDesktopNotification,
        out_cookie: *mut DWORD,
    ) -> HRESULT;

    pub unsafe fn unregister(&self, cookie: u32) -> HRESULT;
}

#[windows_interface::interface("b2f925b9-5a0f-4d2e-9f4d-2b1507593c10")]
pub unsafe trait IVirtualDesktopManagerInternal: IUnknown {
    pub unsafe fn get_desktop_count(&self, monitor: HMONITOR, out_count: *mut UINT) -> HRESULT;

    pub unsafe fn move_view_to_desktop(
        &self,
        view: ComIn<IApplicationView>,
        desktop: ComIn<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn can_move_view_between_desktops(
        &self,
        view: ComIn<IApplicationView>,
        can_move: *mut i32,
    ) -> HRESULT;

    pub unsafe fn get_current_desktop(
        &self,
        monitor: HMONITOR,
        out_desktop: *mut Option<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn get_all_current_desktops(
        &self,
        out_desktops: *mut Option<IObjectArray>,
    ) -> HRESULT;

    pub unsafe fn get_desktops(
        &self,
        monitor: HMONITOR,
        out_desktops: *mut Option<IObjectArray>,
    ) -> HRESULT;

    /// Get next or previous desktop
    ///
    /// Direction values:
    /// 3 = Left direction
    /// 4 = Right direction
    pub unsafe fn get_adjacent_desktop(
        &self,
        in_desktop: ComIn<IVirtualDesktop>,
        direction: UINT,
        out_pp_desktop: *mut Option<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn switch_desktop(
        &self,
        monitor: HMONITOR,
        desktop: ComIn<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn create_desktop(
        &self,
        monitor: HMONITOR,
        out_desktop: *mut Option<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn move_desktop(
        &self,
        in_desktop: ComIn<IVirtualDesktop>,
        monitor: HMONITOR,
        index: UINT,
    ) -> HRESULT;

    pub unsafe fn remove_desktop(
        &self,
        destroy_desktop: ComIn<IVirtualDesktop>,
        fallback_desktop: ComIn<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn find_desktop(
        &self,
        guid: *const GUID,
        out_desktop: *mut Option<IVirtualDesktop>,
    ) -> HRESULT;

    pub unsafe fn get_desktop_switch_include_exclude_views(
        &self,
        desktop: ComIn<IVirtualDesktop>,
        out_pp_desktops1: *mut IObjectArray,
        out_pp_desktops2: *mut IObjectArray,
    ) -> HRESULT;

    pub unsafe fn set_name(&self, desktop: ComIn<IVirtualDesktop>, name: HSTRING) -> HRESULT;
    pub unsafe fn set_wallpaper(&self, desktop: ComIn<IVirtualDesktop>, name: HSTRING) -> HRESULT;
    pub unsafe fn update_wallpaper_for_all(&self, name: HSTRING) -> HRESULT;
}

#[windows_interface::interface("4ce81583-1e4c-4632-a621-07a53543148f")]
pub unsafe trait IVirtualDesktopPinnedApps: IUnknown {
    pub unsafe fn is_app_pinned(&self, app_id: PCWSTR, out_iss: *mut bool) -> HRESULT;
    pub unsafe fn pin_app(&self, app_id: PCWSTR) -> HRESULT;
    pub unsafe fn unpin_app(&self, app_id: PCWSTR) -> HRESULT;

    pub unsafe fn is_view_pinned(
        &self,
        view: ComIn<IApplicationView>,
        out_iss: *mut bool,
    ) -> HRESULT;
    pub unsafe fn pin_view(&self, view: ComIn<IApplicationView>) -> HRESULT;
    pub unsafe fn unpin_view(&self, view: ComIn<IApplicationView>) -> HRESULT;
}
