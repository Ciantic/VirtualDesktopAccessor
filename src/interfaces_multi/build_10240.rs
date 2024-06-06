//! Support for Windows 10.
//!
//! The [`IVirtualDesktopNotification`] and [`IVirtualDesktopManagerInternal`]
//! interfaces do not take "monitor" arguments.
use super::*;

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IApplicationView,
        iid: "9AC0B5C8-1484-4C5B-9533-4134A0F97CEA",
    },
    {
        pub unsafe trait IApplicationView: IUnknown {
            /* IInspecateble */
            pub unsafe fn get_iids(
                &self,
                out_iid_count: *mut ULONG,
                out_opt_iid_array_ptr: *mut *mut GUID,
            ) -> HRESULT;
            pub unsafe fn get_runtime_class_name(
                &self,
                out_opt_class_name: *mut HSTRING,
            ) -> HRESULT;
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
            pub unsafe fn set_position(
                &self,
                view_position: *mut IApplicationViewPosition,
            ) -> HRESULT;
            pub unsafe fn insert_after_window(&self, window: HWND) -> HRESULT;
            pub unsafe fn get_extended_frame_position(&self, rect: *mut RECT) -> HRESULT;
            pub unsafe fn get_app_user_model_id(&self, id: *mut PWSTR) -> HRESULT; // Proc17
            pub unsafe fn set_app_user_model_id(&self, id: PCWSTR) -> HRESULT;
            pub unsafe fn is_equal_by_app_user_model_id(
                &self,
                id: PCWSTR,
                out_result: *mut INT,
            ) -> HRESULT;

            /*** IApplicationView methods ***/
            pub unsafe fn get_view_state(&self, out_state: *mut UINT) -> HRESULT; // Proc20
            pub unsafe fn set_view_state(&self, state: UINT) -> HRESULT; // Proc21
            pub unsafe fn get_neediness(&self, out_neediness: *mut INT) -> HRESULT; // Proc22
            pub unsafe fn get_last_activation_timestamp(
                &self,
                out_timestamp: *mut ULONGLONG,
            ) -> HRESULT;
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
            pub unsafe fn get_position_priority(
                &self,
                out_priority: *mut IShellPositionerPriority,
            ) -> HRESULT;
            pub unsafe fn set_position_priority(
                &self,
                priority: IShellPositionerPriority,
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

            pub unsafe fn query_size_constraints_from_app(&self) -> HRESULT;
            pub unsafe fn on_min_size_preferences_updated(&self, window: HWND) -> HRESULT;
            pub unsafe fn apply_operation(
                &self,
                operation: *mut IApplicationViewOperation,
            ) -> HRESULT;
            pub unsafe fn is_tray(&self, out_is: *mut BOOL) -> HRESULT;
            pub unsafe fn is_in_high_zorder_band(&self, out_is: *mut BOOL) -> HRESULT;
            pub unsafe fn is_splash_screen_presented(&self, out_is: *mut BOOL) -> HRESULT;
            pub unsafe fn flash(&self) -> HRESULT;
            pub unsafe fn get_root_switchable_owner(
                &self,
                app_view: *mut IApplicationView,
            ) -> HRESULT; // proc45
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
    }
);

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IApplicationViewCollection,
        iid: "2C08ADF0-A386-4B35-9250-0FE183476FCC",
    },
    {
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

            pub unsafe fn refresh_collection(&self) -> HRESULT;

            pub unsafe fn register_for_application_view_changes(
                &self,
                listener: IApplicationViewChangeListener,
                out_id: *mut DWORD,
            ) -> HRESULT;

            pub unsafe fn register_for_application_view_position_changes(
                &self,
                listener: IApplicationViewChangeListener,
                out_id: *mut DWORD,
            ) -> HRESULT;

            pub unsafe fn unregister_for_application_view_changes(&self, id: DWORD) -> HRESULT;
        }
    }
);

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IVirtualDesktop,
        iid: "FF72FFDD-BE7E-43FC-9C03-AD81681E88E4",
    },
    {
        pub unsafe trait IVirtualDesktop: IUnknown {
            pub unsafe fn is_view_visible(
                &self,
                p_view: ComIn<IApplicationView>,
                out_bool: *mut u32,
            ) -> HRESULT;
            pub unsafe fn get_id(&self, out_guid: *mut GUID) -> HRESULT;
        }
    }
);

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IVirtualDesktopManagerInternal,
        iid: "F31574D6-B682-4CDC-BD56-1827860ABEC6",
    },
    {
        pub unsafe trait IVirtualDesktopManagerInternal: IUnknown {
            pub unsafe fn get_desktop_count(&self, out_count: *mut UINT) -> HRESULT;

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
                out_desktop: *mut Option<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn get_desktops(&self, out_desktops: *mut Option<IObjectArray>) -> HRESULT;

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

            pub unsafe fn switch_desktop(&self, desktop: ComIn<IVirtualDesktop>) -> HRESULT;

            pub unsafe fn create_desktop(
                &self,
                out_desktop: *mut Option<IVirtualDesktop>,
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
        }
    }
);

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IVirtualDesktopNotification,
        iid: "C179334C-4295-40D3-BEA1-C654D965605A",
    },
    {
        pub unsafe trait IVirtualDesktopNotification: IUnknown {
            pub unsafe fn virtual_desktop_created(
                &self,
                desktop: ComIn<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn virtual_desktop_destroy_begin(
                &self,
                desktop_destroyed: ComIn<IVirtualDesktop>,
                desktop_fallback: ComIn<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn virtual_desktop_destroy_failed(
                &self,
                desktop_destroyed: ComIn<IVirtualDesktop>,
                desktop_fallback: ComIn<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn virtual_desktop_destroyed(
                &self,
                desktop_destroyed: ComIn<IVirtualDesktop>,
                desktop_fallback: ComIn<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn view_virtual_desktop_changed(
                &self,
                view: ComIn<IApplicationView>,
            ) -> HRESULT;

            pub unsafe fn current_virtual_desktop_changed(
                &self,
                desktop_old: ComIn<IVirtualDesktop>,
                desktop_new: ComIn<IVirtualDesktop>,
            ) -> HRESULT;
        }

        /// Implements an unstable interface that is only valid for a single Windows
        /// version using a more stable trait that works for all Windows versions.
        #[windows::core::implement(IVirtualDesktopNotification)]
        pub struct VirtualDesktopNotificationAdaptor<T>
        where
            T: build_dyn::IVirtualDesktopNotification_Impl,
        {
            pub inner: T,
        }
        impl<T> IVirtualDesktopNotification_Impl for VirtualDesktopNotificationAdaptor<T>
        where
            T: build_dyn::IVirtualDesktopNotification_Impl,
        {
            unsafe fn current_virtual_desktop_changed(
                &self,
                desktop_old: ComIn<IVirtualDesktop>,
                desktop_new: ComIn<IVirtualDesktop>,
            ) -> HRESULT {
                self.inner
                    .current_virtual_desktop_changed(desktop_old.into(), desktop_new.into())
            }

            unsafe fn virtual_desktop_created(&self, desktop: ComIn<IVirtualDesktop>) -> HRESULT {
                self.inner.virtual_desktop_created(desktop.into())
            }

            unsafe fn virtual_desktop_destroy_begin(
                &self,
                desktop_destroyed: ComIn<IVirtualDesktop>,
                desktop_fallback: ComIn<IVirtualDesktop>,
            ) -> HRESULT {
                self.inner.virtual_desktop_destroy_begin(
                    desktop_destroyed.into(),
                    desktop_fallback.into(),
                )
            }

            unsafe fn virtual_desktop_destroy_failed(
                &self,
                desktop_destroyed: ComIn<IVirtualDesktop>,
                desktop_fallback: ComIn<IVirtualDesktop>,
            ) -> HRESULT {
                self.inner.virtual_desktop_destroy_failed(
                    desktop_destroyed.into(),
                    desktop_fallback.into(),
                )
            }

            unsafe fn virtual_desktop_destroyed(
                &self,
                desktop_destroyed: ComIn<IVirtualDesktop>,
                desktop_fallback: ComIn<IVirtualDesktop>,
            ) -> HRESULT {
                self.inner
                    .virtual_desktop_destroyed(desktop_destroyed.into(), desktop_fallback.into())
            }

            unsafe fn view_virtual_desktop_changed(
                &self,
                view: ComIn<IApplicationView>,
            ) -> HRESULT {
                self.inner.view_virtual_desktop_changed(view.into())
            }
        }
    }
);

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IVirtualDesktopNotificationService,
        iid: "0CD45E71-D927-4F15-8B0A-8FEF525337BF",
    },
    {
        pub unsafe trait IVirtualDesktopNotificationService: IUnknown {
            pub unsafe fn register(
                &self,
                notification: *mut std::ffi::c_void, // *const IVirtualDesktopNotification,
                out_cookie: *mut DWORD,
            ) -> HRESULT;

            pub unsafe fn unregister(&self, cookie: u32) -> HRESULT;
        }
    }
);

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IVirtualDesktopPinnedApps,
        iid: "4CE81583-1E4C-4632-A621-07A53543148F",
    },
    {
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
    }
);
