//! New for this version is that "monitor" arguments are no longer passed to the
//! [`IVirtualDesktopNotification`] and [`IVirtualDesktopManagerInternal`]
//! interfaces.
//!
//! # References
//!
//! - [New changes for 22621 · Ciantic/VirtualDesktopAccessor@c918946](https://github.com/Ciantic/VirtualDesktopAccessor/commit/c918946421c42a7f022abdf8b4672a4b3ddf2f35#diff-e073b55bb9e1746ee6b3a029ba955df76b3d61e774585c23535bd0b4967d9e18)
//! - <https://github.com/mzomparelli/VirtualDesktop/tree/7e37b9848aef681713224dae558d2e51960cf41e/src/VirtualDesktop/Interop/Build22621_2215>
use super::*;

// These interfaces haven't changed since previous version:
build_10240::IApplicationView!("372E1D3B-38D3-42E4-A15B-8AB2B178F513");
build_10240::IApplicationViewCollection!("1841C6D7-4F9D-42C0-AF41-8747538F10E5");
build_10240::IVirtualDesktopNotificationService!("0cd45e71-d927-4f15-8b0a-8fef525337bf");
build_10240::IVirtualDesktopPinnedApps!("4CE81583-1E4C-4632-A621-07A53543148F");

// But these interfaces have different methods:

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IVirtualDesktop,
        iid: "3F07F4BE-B107-441A-AF0F-39D82529072C",
    },
    {
        pub unsafe trait IVirtualDesktop: IUnknown {
            pub unsafe fn is_view_visible(
                &self,
                p_view: ComIn<IApplicationView>,
                out_bool: *mut u32,
            ) -> HRESULT;
            pub unsafe fn get_id(&self, out_guid: *mut GUID) -> HRESULT;
            // get_monitor removed
            pub unsafe fn get_name(&self, out_string: *mut HSTRING) -> HRESULT;
            pub unsafe fn get_wallpaper(&self, out_string: *mut HSTRING) -> HRESULT;
            // This method is new:
            pub unsafe fn is_remote(&self, out_is_remote: *mut i32) -> HRESULT;
        }
    }
);

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IVirtualDesktopManagerInternal,
        iid: "A3175F2D-239C-4BD2-8AA0-EEBA8B0B138E",
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

            pub unsafe fn move_desktop(
                &self,
                in_desktop: ComIn<IVirtualDesktop>,
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

            pub unsafe fn set_name(
                &self,
                desktop: ComIn<IVirtualDesktop>,
                name: HSTRING,
            ) -> HRESULT;
            pub unsafe fn set_wallpaper(
                &self,
                desktop: ComIn<IVirtualDesktop>,
                name: HSTRING,
            ) -> HRESULT;
            pub unsafe fn update_wallpaper_for_all(&self, name: HSTRING) -> HRESULT;

            pub unsafe fn copy_desktop_state(
                &self,
                view0: ComIn<IApplicationView>,
                view1: ComIn<IApplicationView>,
            ) -> HRESULT;

            // These methods are new:

            pub unsafe fn create_remote_desktop(
                &self,
                name: HSTRING,
                out_desktop: *mut Option<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn switch_remote_desktop(&self, desktop: ComIn<IVirtualDesktop>) -> HRESULT;

            pub unsafe fn switch_desktop_with_animation(
                &self,
                desktop: ComIn<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn get_last_active_desktop(
                &self,
                out_desktop: *mut Option<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn wait_for_animation_to_complete(&self) -> HRESULT;
        }
    }
);

reusable_com_interface!(
    MacroOptions {
        temp_macro_name: _IVirtualDesktopNotification,
        iid: "B287FA1C-7771-471A-A2DF-9B6B21F0D675",
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

            pub unsafe fn virtual_desktop_moved(
                &self,
                desktop: ComIn<IVirtualDesktop>,
                old_index: i64,
                new_index: i64,
            ) -> HRESULT;

            pub unsafe fn virtual_desktop_name_changed(
                &self,
                desktop: ComIn<IVirtualDesktop>,
                name: HSTRING,
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

            pub unsafe fn virtual_desktop_wallpaper_changed(
                &self,
                desktop: ComIn<IVirtualDesktop>,
                name: HSTRING,
            ) -> HRESULT;

            pub unsafe fn virtual_desktop_switched(
                &self,
                desktop: ComIn<IVirtualDesktop>,
            ) -> HRESULT;

            pub unsafe fn remote_virtual_desktop_connected(
                &self,
                desktop: ComIn<IVirtualDesktop>,
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

            unsafe fn virtual_desktop_moved(
                &self,
                desktop: ComIn<IVirtualDesktop>,
                old_index: i64,
                new_index: i64,
            ) -> HRESULT {
                self.inner
                    .virtual_desktop_moved(desktop.into(), old_index, new_index)
            }

            unsafe fn virtual_desktop_name_changed(
                &self,
                desktop: ComIn<IVirtualDesktop>,
                name: HSTRING,
            ) -> HRESULT {
                self.inner
                    .virtual_desktop_name_changed(desktop.into(), name)
            }

            unsafe fn view_virtual_desktop_changed(
                &self,
                view: ComIn<IApplicationView>,
            ) -> HRESULT {
                self.inner.view_virtual_desktop_changed(view.into())
            }

            unsafe fn current_virtual_desktop_changed(
                &self,
                desktop_old: ComIn<IVirtualDesktop>,
                desktop_new: ComIn<IVirtualDesktop>,
            ) -> HRESULT {
                self.inner
                    .current_virtual_desktop_changed(desktop_old.into(), desktop_new.into())
            }

            unsafe fn virtual_desktop_wallpaper_changed(
                &self,
                desktop: ComIn<IVirtualDesktop>,
                name: HSTRING,
            ) -> HRESULT {
                self.inner
                    .virtual_desktop_wallpaper_changed(desktop.into(), name)
            }

            unsafe fn virtual_desktop_switched(&self, desktop: ComIn<IVirtualDesktop>) -> HRESULT {
                self.inner.virtual_desktop_switched(desktop.into())
            }

            unsafe fn remote_virtual_desktop_connected(
                &self,
                desktop: ComIn<IVirtualDesktop>,
            ) -> HRESULT {
                self.inner.remote_virtual_desktop_connected(desktop.into())
            }
        }
    }
);
