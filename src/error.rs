use crate::HRESULT;

#[derive(Debug, PartialEq, Clone)]
pub enum Error {
    /// Window is not found
    WindowNotFound,

    /// Desktop with given ID is not found
    DesktopNotFound,

    /// Unable to create service, ensure that explorer.exe is running
    ServiceNotCreated,

    /// Some unhandled COM error
    ComError(HRESULT),

    /// This should not happen, this means that successful COM call allocated a
    /// null pointer, in this case it is an error in the COM service, or it's
    /// usage.
    ComAllocatedNullPtr,
}

impl From<HRESULT> for Error {
    fn from(hr: HRESULT) -> Self {
        Error::ComError(hr)
    }
}
impl From<HRESULT> for Result<(), Error> {
    fn from(item: HRESULT) -> Self {
        if !item.failed() {
            Ok(())
        } else {
            Err(item.into())
        }
    }
}
