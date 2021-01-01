use std::{error::Error, fmt};

/// These are considered common windows errors
///
/// https://docs.microsoft.com/en-us/windows/win32/seccrypto/common-hresult-values
#[derive(Debug)]
pub enum WinError {
    InvalidArg,
    OutOfMemory,
    Unexpected,
    Abort,
    AccessDenied,
    Fail,
    Handle,
    NoInterface,
    NotImpl,
    Pointer,
}

impl fmt::Display for WinError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            WinError::InvalidArg => {
                write!(f, "One or more arguments are not valid")
            }
            WinError::OutOfMemory => {
                write!(f, "Failed to allocate necessary memory")
            }
            WinError::Unexpected => {
                write!(f, "Unexpected failure")
            }
            WinError::Abort => {
                write!(f, "Operation aborted")
            }
            WinError::AccessDenied => {
                write!(f, "General access denied error")
            }
            WinError::Fail => {
                write!(f, "Unspecified failure")
            }
            WinError::Handle => {
                write!(f, "Handle that is not valid")
            }
            WinError::NoInterface => {
                write!(f, "No such interface supported")
            }
            WinError::NotImpl => {
                write!(f, "Not implemented")
            }
            WinError::Pointer => {
                write!(f, "Pointer that is not valid")
            }
        }
    }
}
impl Error for WinError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        None
    }
}
