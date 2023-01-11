use crate::{
    desktopid::DesktopID, get_desktop_name, get_index_by_desktop, set_desktop_name, Error,
};
use std::fmt::Debug;

#[derive(Copy, Clone, PartialEq)]
pub struct Desktop {
    pub(crate) id: DesktopID,
}

impl Debug for Desktop {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Desktop({:?})", self.id.0)
    }
}

impl Desktop {
    pub(crate) fn empty() -> Desktop {
        Desktop {
            id: DesktopID::default(),
        }
    }

    /// Name of the desktop
    pub fn get_name(&self) -> Result<String, Error> {
        get_desktop_name(&self)
    }

    /// Set the name of the desktop
    pub fn set_name(&self, name: &str) -> Result<(), Error> {
        set_desktop_name(self, name)
    }

    /// Index of the desktop among other
    pub fn get_index(&self) -> Result<usize, Error> {
        get_index_by_desktop(&self)
    }

    pub fn get_id(&self) -> DesktopID {
        self.id
    }
}
