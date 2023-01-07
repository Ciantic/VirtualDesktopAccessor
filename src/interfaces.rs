#![allow(non_snake_case)]
#![allow(non_upper_case_globals)]

use crate::{desktopid::DesktopID, hresult::HRESULT, hstring::HSTRING};
use com::com_interface;
use com::{interfaces::IUnknown, sys::CLSID, ComRc, IID};
use std::ffi::c_void;

type HMONITOR = UINT;

pub const CLSID_ImmersiveShell: CLSID = CLSID {
    data1: 0xC2F03A33,
    data2: 0x21F5,
    data3: 0x47FA,
    data4: [0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39],
};

pub const CLSID_IVirtualNotificationService: CLSID = CLSID {
    data1: 0xA501FDEC,
    data2: 0x4A09,
    data3: 0x464C,
    data4: [0xAE, 0x4E, 0x1B, 0x9C, 0x21, 0xB8, 0x49, 0x18],
};

/*
pub const IID_IVirtualDesktopNotification: IID = IID {
    data1: 0xCD403E52,
    data2: 0xDEED,
    data3: 0x4C13,
    data4: [0xB4, 0x37, 0xB9, 0x83, 0x80, 0xF2, 0xB1, 0xE8],
};
*/

pub const CLSID_VirtualDesktopManagerInternal: IID = IID {
    data1: 0xC5E0CDCA,
    data2: 0x7B6E,
    data3: 0x41B2,
    data4: [0x9F, 0xC4, 0xD9, 0x39, 0x75, 0xCC, 0x46, 0x7B],
};

pub const CLSID_VirtualDesktopPinnedApps: IID = IID {
    data1: 0xb5a399e7,
    data2: 0x1c87,
    data3: 0x46b8,
    data4: [0x88, 0xe9, 0xfc, 0x57, 0x47, 0xb1, 0x71, 0xbd],
};

// primitives
pub type HWND = u32;
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
// type HSTRING = LPVOID;

// Ignore following API's:
// type IShellPositionerPriority = UINT;
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

/*
Notepad++ replaces for fn PascalCase -> fn pascal_case
fn ([A-Z])
fn \L$1

fn ([a-z_]+)([A-Z]+)
fn $1_\L$2

Other:
STDMETHOD\((\S+)\)
fn $1

) PURE;
) -> HRESULT;
*/

#[com_interface("6D5140C1-7436-11CE-8034-00AA006009FA")]
pub trait IServiceProvider: IUnknown {
    unsafe fn query_service(
        &self,
        guidService: *const com::sys::GUID,
        riid: *const IID,
        ppvObject: *mut *mut c_void,
    ) -> HRESULT;
    // unsafe fn remote_query_service(
    //     &self,
    //     guidService: *const DesktopID,
    //     riid: *const IID,
    //     ppvObject: *mut *mut c_void,
    // ) -> HRESULT;
}

#[com_interface("a5cd92ff-29be-454c-8d04-d82879fb3f1b")]
pub trait IVirtualDesktopManager: IUnknown {
    unsafe fn is_window_on_current_desktop(
        &self,
        topLevelWindow: HWND,
        outOnCurrentDesktop: *mut bool,
    ) -> HRESULT;
    unsafe fn get_desktop_by_window(
        &self,
        topLevelWindow: HWND,
        outDesktopId: *mut DesktopID,
    ) -> HRESULT;
    unsafe fn move_window_to_desktop(
        &self,
        topLevelWindow: HWND,
        desktopId: *const DesktopID,
    ) -> HRESULT;
}

#[com_interface("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
pub trait IApplicationView: IUnknown {
    /* IInspecateble */
    unsafe fn get_iids(&self, outIidCount: *mut ULONG, outOptIidArrayPtr: *mut *mut IID)
        -> HRESULT;
    unsafe fn get_runtime_class_name(&self, outOptClassName: *mut HSTRING) -> HRESULT;
    unsafe fn get_trust_level(&self, ptrTrustLevel: LPVOID) -> HRESULT;

    /* IApplicationView methods */
    unsafe fn set_focus(&self) -> HRESULT;
    unsafe fn switch_to(&self) -> HRESULT;

    unsafe fn try_invoke_back(&self, ptrAsyncCallback: IAsyncCallback) -> HRESULT;
    unsafe fn get_thumbnail_window(&self, outHwnd: *mut HWND) -> HRESULT;
    unsafe fn get_monitor(&self, outMonitors: *mut *mut IImmersiveMonitor) -> HRESULT;
    unsafe fn get_visibility(&self, outInt: LPVOID) -> HRESULT;
    unsafe fn set_cloak(
        &self,
        applicationViewCloakType: APPLICATION_VIEW_CLOAK_TYPE,
        unknown: INT,
    ) -> HRESULT;
    unsafe fn get_position(&self, unknownIid: *const IID, unknownArrayPtr: LPVOID) -> HRESULT;
    unsafe fn set_position(&self, viewPosition: *mut IApplicationViewPosition) -> HRESULT;
    unsafe fn insert_after_window(&self, window: HWND) -> HRESULT;
    unsafe fn get_extended_frame_position(&self, rect: *mut RECT) -> HRESULT;
    unsafe fn get_app_user_model_id(&self, id: *mut PWSTR) -> HRESULT; // Proc17
    unsafe fn set_app_user_model_id(&self, id: PCWSTR) -> HRESULT;
    unsafe fn is_equal_by_app_user_model_id(&self, id: PCWSTR, outResult: *mut INT) -> HRESULT;

    /*** IApplicationView methods ***/
    unsafe fn get_view_state(&self, outState: *mut UINT) -> HRESULT; // Proc20
    unsafe fn set_view_state(&self, state: UINT) -> HRESULT; // Proc21
    unsafe fn get_neediness(&self, outNeediness: *mut INT) -> HRESULT; // Proc22
    unsafe fn get_last_activation_timestamp(&self, outTimestamp: *mut ULONGLONG) -> HRESULT;
    unsafe fn set_last_activation_timestamp(&self, timestamp: ULONGLONG) -> HRESULT;
    unsafe fn get_virtual_desktop_id(&self, outDesktopGuid: *mut DesktopID) -> HRESULT;
    unsafe fn set_virtual_desktop_id(&self, desktopGuid: *const DesktopID) -> HRESULT;
    unsafe fn get_show_in_switchers(&self, outShow: *mut INT) -> HRESULT;
    unsafe fn set_show_in_switchers(&self, show: INT) -> HRESULT;
    unsafe fn get_scale_factor(&self, outScaleFactor: *mut INT) -> HRESULT;
    unsafe fn can_receive_input(&self, outCan: *mut BOOL) -> HRESULT;
    unsafe fn get_compatibility_policy_type(
        &self,
        outPolicyType: *mut APPLICATION_VIEW_COMPATIBILITY_POLICY,
    ) -> HRESULT;
    unsafe fn set_compatibility_policy_type(
        &self,
        policyType: APPLICATION_VIEW_COMPATIBILITY_POLICY,
    ) -> HRESULT;

    //unsafe fn get_position_priority(&self, THIS_ IShellPositionerPriority**) -> HRESULT; // removed in 1803
    //unsafe fn set_position_priority(&self, THIS_ IShellPositionerPriority*) -> HRESULT; // removed in 1803

    unsafe fn get_size_constraints(
        &self,
        monitor: *mut IImmersiveMonitor,
        outSize1: *mut SIZE,
        outSize2: *mut SIZE,
    ) -> HRESULT;
    unsafe fn get_size_constraints_for_dpi(
        &self,
        dpi: UINT,
        outSize1: *mut SIZE,
        outSize2: *mut SIZE,
    ) -> HRESULT;
    unsafe fn set_size_constraints_for_dpi(
        &self,
        dpi: *const UINT,
        size1: *const SIZE,
        size2: *const SIZE,
    ) -> HRESULT;

    //unsafe fn query_size_constraints_from_app)(&self, THIS PURE; // removed in 1803

    unsafe fn on_min_size_preferences_updated(&self, window: HWND) -> HRESULT;
    unsafe fn apply_operation(&self, operation: *mut IApplicationViewOperation) -> HRESULT;
    unsafe fn is_tray(&self, outIs: *mut BOOL) -> HRESULT;
    unsafe fn is_in_high_zorder_band(&self, outIs: *mut BOOL) -> HRESULT;
    unsafe fn is_splash_screen_presented(&self, outIs: *mut BOOL) -> HRESULT;
    unsafe fn flash(&self) -> HRESULT;
    unsafe fn get_root_switchable_owner(
        &self,
        appView: *mut Option<ComRc<dyn IApplicationView>>,
    ) -> HRESULT; // proc45
    unsafe fn enumerate_ownership_tree(
        &self,
        objects: *mut Option<ComRc<dyn IObjectArray>>,
    ) -> HRESULT; // proc46

    unsafe fn get_enterprise_id(&self, outId: *mut PWSTR) -> HRESULT; // proc47
    unsafe fn is_mirrored(&self, outIs: *mut BOOL) -> HRESULT; //

    unsafe fn unknown1(&self, arg: *mut INT) -> HRESULT;
    unsafe fn unknown2(&self, arg: *mut INT) -> HRESULT;
    unsafe fn unknown3(&self, arg: *mut INT) -> HRESULT;
    unsafe fn unknown4(&self, arg: INT) -> HRESULT;
    unsafe fn unknown5(&self, arg: *mut INT) -> HRESULT;
    unsafe fn unknown6(&self, arg: INT) -> HRESULT;
    unsafe fn unknown7(&self) -> HRESULT;
    unsafe fn unknown8(&self, arg: *mut INT) -> HRESULT;
    unsafe fn unknown9(&self, arg: INT) -> HRESULT;
    unsafe fn unknown10(&self, arg: INT, arg2: INT) -> HRESULT;
    unsafe fn unknown11(&self, arg: INT) -> HRESULT;
    unsafe fn unknown12(&self, arg: *mut SIZE) -> HRESULT;
}

#[com_interface("92ca9dcd-5622-4bba-a805-5e9f541bd8c9")]
pub trait IObjectArray: IUnknown {
    unsafe fn get_count(&self, outPcObjects: *mut UINT) -> HRESULT;
    unsafe fn get_at(&self, uiIndex: UINT, riid: *const IID, outValue: *mut *mut c_void)
        -> HRESULT;
}

#[com_interface("536D3495-B208-4CC9-AE26-DE8111275BF8")]
pub trait IVirtualDesktop: IUnknown {
    unsafe fn is_view_visible(
        &self,
        pView: ComRc<dyn IApplicationView>,
        outBool: *mut u32,
    ) -> HRESULT;
    unsafe fn get_id(&self, outGuid: *mut DesktopID) -> HRESULT;
    unsafe fn get_monitor(&self, outMonitor: *mut HMONITOR) -> HRESULT;
    unsafe fn get_name(&self, outString: *mut HSTRING) -> HRESULT;
    unsafe fn get_wallpaper(&self, outString: *mut HSTRING) -> HRESULT;
}

// #[com_interface("31ebde3f-6ec3-4cbd-b9fb-0ef6d09b41f4")]
// pub trait IVirtualDesktop2: IUnknown {
//     unsafe fn is_view_visible(
//         &self,
//         pView: ComRc<dyn IApplicationView>,
//         outBool: *mut u32,
//     ) -> HRESULT;
//     unsafe fn get_id(&self, outGuid: *mut DesktopID) -> HRESULT;
//     unsafe fn get_name(&self, outName: *mut HSTRING) -> HRESULT;
// }

#[com_interface("1841c6d7-4f9d-42c0-af41-8747538f10e5")]
pub trait IApplicationViewCollection: IUnknown {
    unsafe fn get_views(&self, outViews: *mut Option<ComRc<dyn IObjectArray>>) -> HRESULT;

    unsafe fn get_views_by_zorder(&self, outViews: *mut Option<ComRc<dyn IObjectArray>>)
        -> HRESULT;

    unsafe fn get_views_by_app_user_model_id(
        &self,
        id: PCWSTR,
        outViews: *mut Option<ComRc<dyn IObjectArray>>,
    ) -> HRESULT;

    unsafe fn get_view_for_hwnd(
        &self,
        window: HWND,
        outView: *mut Option<ComRc<dyn IApplicationView>>,
    ) -> HRESULT;

    unsafe fn get_view_for_application(
        &self,
        app: *const IImmersiveApplication,
        outView: *mut Option<ComRc<dyn IApplicationView>>,
    ) -> HRESULT;

    unsafe fn get_view_for_app_user_model_id(
        &self,
        id: PCWSTR,
        outView: *mut Option<ComRc<dyn IApplicationView>>,
    ) -> HRESULT;

    unsafe fn get_view_in_focus(
        &self,
        outView: *mut Option<ComRc<dyn IApplicationView>>,
    ) -> HRESULT;

    unsafe fn try_get_last_active_visible_view(
        &self,
        outView: *mut Option<ComRc<dyn IApplicationView>>,
    ) -> HRESULT;

    unsafe fn refresh_collection(&self) -> HRESULT;

    unsafe fn register_for_application_view_changes(
        &self,
        listener: *const IApplicationViewChangeListener,
        outId: *mut DWORD,
    ) -> HRESULT;

    unsafe fn unregister_for_application_view_changes(&self, id: DWORD) -> HRESULT;
}

#[com_interface("CD403E52-DEED-4C13-B437-B98380F2B1E8")]
pub trait IVirtualDesktopNotification: IUnknown {
    unsafe fn virtual_desktop_created(&self, desktop: ComRc<dyn IVirtualDesktop>) -> HRESULT;

    unsafe fn virtual_desktop_destroy_begin(
        &self,
        desktopDestroyed: ComRc<dyn IVirtualDesktop>,
        desktopFallback: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    unsafe fn virtual_desktop_destroy_failed(
        &self,
        desktopDestroyed: ComRc<dyn IVirtualDesktop>,
        desktopFallback: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    unsafe fn virtual_desktop_destroyed(
        &self,
        desktopDestroyed: ComRc<dyn IVirtualDesktop>,
        desktopFallback: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    unsafe fn virtual_desktop_is_per_monitor_changed(&self, isPerMonitor: bool) -> HRESULT;

    unsafe fn virtual_desktop_moved(
        &self,
        desktop: ComRc<dyn IVirtualDesktop>,
        oldIndex: u64,
        newIndex: u64,
    ) -> HRESULT;

    unsafe fn virtual_desktop_name_changed(
        &self,
        desktop: ComRc<dyn IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT;

    unsafe fn view_virtual_desktop_changed(&self, view: ComRc<dyn IApplicationView>) -> HRESULT;

    unsafe fn current_virtual_desktop_changed(
        &self,
        desktopOld: ComRc<dyn IVirtualDesktop>,
        desktopNew: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;
}

// This is here for completness, however I think this cannot be used at all. One
// can't register a IVirtualDesktopNotification2, it just gives an error when
// given to registration method. This is not finished by Microsoft engineers.
#[com_interface("1ba7cf30-3591-43fa-abfa-4aaf7abeedb7")]
pub trait IVirtualDesktopNotification2: IUnknown {
    unsafe fn virtual_desktop_created(&self, desktop: ComRc<dyn IVirtualDesktop>) -> HRESULT;

    unsafe fn virtual_desktop_destroy_begin(
        &self,
        desktopDestroyed: ComRc<dyn IVirtualDesktop>,
        desktopFallback: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    unsafe fn virtual_desktop_destroy_failed(
        &self,
        desktopDestroyed: ComRc<dyn IVirtualDesktop>,
        desktopFallback: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    unsafe fn virtual_desktop_destroyed(
        &self,
        desktopDestroyed: ComRc<dyn IVirtualDesktop>,
        desktopFallback: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    unsafe fn view_virtual_desktop_changed(&self, view: ComRc<dyn IApplicationView>) -> HRESULT;

    unsafe fn current_virtual_desktop_changed(
        &self,
        desktopOld: ComRc<dyn IVirtualDesktop>,
        desktopNew: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;
    unsafe fn virtual_desktop_renamed(
        &self,
        desktop: ComRc<dyn IVirtualDesktop>,
        newName: HSTRING,
    ) -> HRESULT;
}

#[com_interface("0CD45E71-D927-4F15-8B0A-8FEF525337BF")]
pub trait IVirtualDesktopNotificationService: IUnknown {
    unsafe fn register(
        &self,
        // notification: ComPtr<dyn IVirtualDesktopNotification>,
        notification: ComRc<dyn IVirtualDesktopNotification>,
        // notification: *mut c_void,
        outCookie: *mut DWORD,
    ) -> HRESULT;

    unsafe fn unregister(&self, cookie: DWORD) -> HRESULT;
}

#[com_interface("b2f925b9-5a0f-4d2e-9f4d-2b1507593c10")]
pub trait IVirtualDesktopManagerInternal: IUnknown {
    // Proc3
    unsafe fn get_count(&self, monitor: HMONITOR, outCount: *mut UINT) -> HRESULT;

    // Proc4
    unsafe fn move_view_to_desktop(
        &self,
        view: ComRc<dyn IApplicationView>,
        desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    // Proc5
    unsafe fn can_move_view_between_desktops(
        &self,
        view: ComRc<dyn IApplicationView>,
        canMove: *mut i32,
    ) -> HRESULT;

    // Proc6
    unsafe fn get_current_desktop(
        &self,
        monitor: HMONITOR,
        outDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>,
    ) -> HRESULT;

    // Proc7
    unsafe fn get_all_current_desktops(
        &self,
        outDesktops: *mut Option<ComRc<dyn IObjectArray>>,
    ) -> HRESULT;

    unsafe fn get_desktops(
        &self,
        monitor: HMONITOR,
        outDesktops: *mut Option<ComRc<dyn IObjectArray>>,
    ) -> HRESULT;

    // Proc8
    unsafe fn get_adjacent_desktop(
        &self,
        inDesktop: ComRc<dyn IVirtualDesktop>,
        direction: UINT,
        outDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>,
    ) -> HRESULT;

    // Proc9
    unsafe fn switch_desktop(
        &self,
        monitor: HMONITOR,
        desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    // Proc10
    unsafe fn create_desktop(
        &self,
        monitor: HMONITOR,
        outDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>,
    ) -> HRESULT;

    unsafe fn move_desktop(
        &self,
        inDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>,
        monitor: HMONITOR,
        index: UINT,
    ) -> HRESULT;

    // Proc11
    unsafe fn remove_desktop(
        &self,
        destroyDesktop: ComRc<dyn IVirtualDesktop>,
        fallbackDesktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    // Proc12
    unsafe fn find_desktop(
        &self,
        guid: *const DesktopID,
        outDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>,
    ) -> HRESULT;

    unsafe fn get_desktop_switch_include_exclude_views(
        &self,
        desktop: ComRc<dyn IVirtualDesktop>,
        outPpDesktops1: *mut Option<ComRc<dyn IObjectArray>>,
        outPpDesktops2: *mut Option<ComRc<dyn IObjectArray>>,
    ) -> HRESULT;

    unsafe fn set_name(&self, desktop: ComRc<dyn IVirtualDesktop>, name: HSTRING) -> HRESULT;

    unsafe fn set_wallpaper(&self, desktop: ComRc<dyn IVirtualDesktop>, name: HSTRING) -> HRESULT;

    unsafe fn update_wallpaper_for_all(&self, name: HSTRING) -> HRESULT;

    /*
        virtual HRESULT STDMETHODCALLTYPE CopyDesktopState(
            _In_ IApplicationView* p0,
            _In_ IApplicationView* p1) = 0;

        virtual HRESULT STDMETHODCALLTYPE GetDesktopIsPerMonitor(
            _Out_ BOOL* p0) = 0;

        virtual HRESULT STDMETHODCALLTYPE SetDesktopIsPerMonitor(
            _In_ BOOL p0) = 0;
    */
}

// Notice that engineers at Microsoft have been in hurry, this is basically
// useless for anything else than renaming the desktop! This is because all of
// the signatures still refer to plain old IVirtualDesktop instead of
// IVirtualDesktop2.
/*
#[com_interface("0f3a72b0-4566-487e-9a33-4ed302f6d6ce")]
pub trait IVirtualDesktopManagerInternal2: IUnknown {
    // Proc3
    unsafe fn get_count(&self, outCount: *mut UINT) -> HRESULT;

    // Proc4
    unsafe fn move_view_to_desktop(
        &self,
        view: ComRc<dyn IApplicationView>,
        desktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    // Proc5
    unsafe fn can_move_view_between_desktops(
        &self,
        view: ComRc<dyn IApplicationView>,
        canMove: *mut i32,
    ) -> HRESULT;

    // Proc6
    unsafe fn get_current_desktop(
        &self,
        outDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>,
    ) -> HRESULT;

    // Proc7
    unsafe fn get_desktops(&self, outDesktops: *mut Option<ComRc<dyn IObjectArray>>) -> HRESULT;

    // Proc8
    unsafe fn get_adjacent_desktop(
        &self,
        inDesktop: ComRc<dyn IVirtualDesktop>,
        outDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>,
    ) -> HRESULT;

    // Proc9
    unsafe fn switch_desktop(&self, desktop: ComRc<dyn IVirtualDesktop2>) -> HRESULT;

    // Proc10
    unsafe fn create_desktop(&self, outDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>)
        -> HRESULT;

    // Proc11
    unsafe fn remove_desktop(
        &self,
        destroyDesktop: ComRc<dyn IVirtualDesktop>,
        fallbackDesktop: ComRc<dyn IVirtualDesktop>,
    ) -> HRESULT;

    // Proc12
    unsafe fn find_desktop(
        &self,
        guid: *const DesktopID,
        outDesktop: *mut Option<ComRc<dyn IVirtualDesktop>>,
    ) -> HRESULT;

    // Proc13
    unsafe fn unknown(
        &self,
        desktop: ComRc<dyn IVirtualDesktop>,
        out1: *mut Option<ComRc<dyn IObjectArray>>,
        out2: *mut Option<ComRc<dyn IObjectArray>>,
    ) -> HRESULT;

    // Proc12
    unsafe fn rename_desktop(
        &self,
        inDesktop: ComRc<dyn IVirtualDesktop>,
        name: HSTRING,
    ) -> HRESULT;
}
 */

#[com_interface("4ce81583-1e4c-4632-a621-07a53543148f")]
pub trait IVirtualDesktopPinnedApps: IUnknown {
    unsafe fn is_app_pinned(&self, appId: PCWSTR, outIs: *mut bool) -> HRESULT;
    unsafe fn pin_app(&self, appId: PCWSTR) -> HRESULT;
    unsafe fn unpin_app(&self, appId: PCWSTR) -> HRESULT;

    unsafe fn is_view_pinned(&self, view: ComRc<dyn IApplicationView>, outIs: *mut bool)
        -> HRESULT;
    unsafe fn pin_view(&self, view: ComRc<dyn IApplicationView>) -> HRESULT;
    unsafe fn unpin_view(&self, view: ComRc<dyn IApplicationView>) -> HRESULT;
}
