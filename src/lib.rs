//! winvd - crate for accessing the Windows Virtual Desktop API
//!
//! All functions taking `Into<Desktop>` can take either a index or a GUID.
//!
//! # Examples
//! * Get first desktop name by index `get_desktop(0).get_name()`
//! * Get second desktop name by index `get_desktop(1).get_name()`
//! * Get desktop name by GUID `get_desktop(GUID(123...)).get_name()`
//! * Switch to fifth desktop by index `switch_desktop(4)`
//! * Get third desktop name `get_desktop(2).get_name()`
mod comobjects;
mod desktop;
mod events;
mod interfaces;
mod listener;
mod log;

#[cfg(feature = "integration-tests")]
#[cfg(test)]
mod tests;

pub use comobjects::Error;
pub use desktop::*;
pub use events::*;
pub use listener::DesktopEventThread;
pub type Result<T> = std::result::Result<T, Error>;

#[macro_use]
extern crate macro_rules_attribute;
