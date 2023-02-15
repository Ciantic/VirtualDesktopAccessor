use crate::Error;

mod interfaces;
mod raw;

pub mod desktop;

#[cfg(test)]
mod experiments;

pub type Result<T> = std::result::Result<T, Error>;
