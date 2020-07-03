use crate::HRESULT;

#[derive(Debug, PartialEq, Clone)]
pub enum Error {
    /// Window is not found
    WindowNotFound,

    /// Desktop with given ID is not found
    DesktopNotFound,

    /// Unable to get the service provider, this is raised for example when
    /// explorer.exe is not running.
    ComClassNotRegistered,

    /// COM apartment is not initialized, please use appropriate constructor
    /// `VirtualDesktopService::create_with_com` or initialize with direct call
    /// to winapi function `CoInitialize`.
    ComNotInitialized,

    /// When RPC server is not available, this is an indication that explorer
    /// needs to be restarted
    ComRpcUnavailable,

    /// Some COM result error
    ComError(HRESULT),

    /// This should not happen
    NullPtr,

    /// This should not happen
    PoisonError,
}

impl From<HRESULT> for Error {
    fn from(hr: HRESULT) -> Self {
        match hr {
            HRESULT(0x80040154) => Error::ComClassNotRegistered,
            HRESULT(0x800401F0) => Error::ComNotInitialized,
            HRESULT(0x800706BA) => Error::ComRpcUnavailable,
            _ => Error::ComError(hr),
        }
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
