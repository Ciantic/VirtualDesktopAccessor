use std::fmt::Debug;

#[derive(PartialEq, PartialOrd, Clone, Copy)]
pub struct HRESULT(i32);

impl HRESULT {
    #[inline]
    pub fn failed(&self) -> bool {
        self < &HRESULT::from_i32(0)
    }
    #[inline]
    pub fn failed_with(&self, u: u32) -> bool {
        self.0 == u as i32
    }
    #[inline]
    pub fn ok() -> HRESULT {
        HRESULT(0)
    }
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
