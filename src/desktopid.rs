#[derive(PartialEq, Debug, Clone, Copy)]
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
