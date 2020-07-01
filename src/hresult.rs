use std::fmt::Debug;

/// This is wrapper for handling HRESULT values.
///
/// Value is printed in hexadecimal format for convinience, this is usually the
/// format it's given in MSDN.
#[derive(PartialEq, PartialOrd, Clone, Copy)]
pub struct HRESULT(i32);

impl HRESULT {
    /// Is any failure?
    #[inline]
    pub fn failed(&self) -> bool {
        self.0 < 0
    }

    /// Test if the HRESULT failed with certain value, e.g.
    /// hresult.failed_with(0x800706BA)
    #[inline]
    pub fn failed_with(&self, u: u32) -> bool {
        self.0 == u as i32
    }

    /// Indicates not a failure
    #[inline]
    pub fn ok() -> HRESULT {
        HRESULT(0)
    }

    /// Create value
    #[inline]
    pub fn from_i32(v: i32) -> HRESULT {
        HRESULT(v)
    }
}

impl Default for HRESULT {
    fn default() -> Self {
        HRESULT(0)
    }
}

impl Debug for HRESULT {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "HRESULT(0x{:X})", self.0)
    }
}

impl From<i32> for HRESULT {
    fn from(item: i32) -> Self {
        HRESULT::from_i32(item)
    }
}
