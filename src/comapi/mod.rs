mod interfaces;
// mod raw;
mod raw2;

pub use raw2::Error;
pub mod desktop;

// #[cfg(test)]
// mod experiments;

pub type Result<T> = std::result::Result<T, Error>;
