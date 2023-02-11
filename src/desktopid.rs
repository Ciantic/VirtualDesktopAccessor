/// cbindgen:field-names=[Id]
#[derive(PartialEq, Debug, Clone, Copy)]
#[repr(C)]
pub struct DesktopID(pub(crate) com::sys::GUID);

impl Default for DesktopID {
    fn default() -> Self {
        DesktopID(com::sys::GUID {
            data1: 0,
            data2: 0,
            data3: 0,
            data4: [0, 0, 0, 0, 0, 0, 0, 0],
        })
    }
}

impl DesktopID {
    pub fn new(data: (u32, u16, u16, [u8; 8])) -> Self {
        DesktopID(com::sys::GUID {
            data1: data.0,
            data2: data.1,
            data3: data.2,
            data4: data.3,
        })
    }
    pub fn get_data(&self) -> (u32, u16, u16, [u8; 8]) {
        (self.0.data1, self.0.data2, self.0.data3, self.0.data4)
    }
}
