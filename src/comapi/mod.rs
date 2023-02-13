use crate::Error;

mod interfaces;
mod raw;

pub mod desktop;
pub mod numbered;
pub mod windowing;

#[cfg(test)]
mod experiments;

pub type Result<T> = std::result::Result<T, Error>;
