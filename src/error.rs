use std::{error::Error, fmt};

/// These are considered common windows errors
///
/// https://docs.microsoft.com/en-us/windows/win32/seccrypto/common-hresult-values
#[derive(Debug)]
pub enum WinError {
    InvalidArg(String),
    OutOfMemory,
    Unexpected,
    Abort,
    AccessDenied,
    Fail,
    Handle,
    NoInterface,
    NotImpl,
    Pointer(String),
    UnknownError(String),
}

impl fmt::Display for WinError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            WinError::InvalidArg(description) => {
                write!(f, "One or more arguments are not valid: {}", description)
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
            WinError::Pointer(description) => {
                write!(f, "Pointer that is not valid: {}", description)
            }
            WinError::UnknownError(description) => {
                write!(f, "{}", description)
            }
        }
    }
}
impl Error for WinError {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        None
    }
}
