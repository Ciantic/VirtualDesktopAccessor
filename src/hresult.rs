use std::fmt::Debug;

use crate::Error;

/// This is wrapper for handling HRESULT values.
///
/// Value is printed in hexadecimal format for convinience, this is usually the
/// format it's given in MSDN. Similarily this can be pattern matched using
/// hexadecimal format: HRESULT(0x800706BA)
#[derive(PartialEq, PartialOrd, Clone, Copy)]
#[repr(C)]
pub struct HRESULT(pub u32);

impl HRESULT {
    /// Is any failure?
    #[inline]
    pub fn failed(&self) -> bool {
        (self.0 as i32) < 0
    }

    /// Indicates not a failure
    #[inline]
    pub fn ok() -> HRESULT {
        HRESULT(0)
    }

    pub fn as_result(&self) -> Result<(), Error> {
        if self.failed() {
            if self.0 == 0x800706BA {
                return Err(Error::ServiceNotConnected);
            }
            Err(Error::ComError(self.clone()))
        } else {
            Ok(())
        }
    }

    #[inline]
    pub fn panic_if_failed(&self) {
        if self.failed() {
            panic!("HRESULT failed: {:?}", self);
        }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_ok() {
        assert_eq!(HRESULT::ok().failed(), false);
    }

    #[test]
    fn test_failure() {
        assert_eq!(HRESULT(0x800706BA).failed(), true);
    }
}
