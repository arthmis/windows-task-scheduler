use std::{error::Error, fmt};

#[derive(Debug)]
pub enum TaskError {
    WinError(WinError),
    ComError(ComError),
    TaskServiceError(TaskServiceError),
    Error(String),
}

impl From<WinError> for TaskError {
    fn from(error: WinError) -> Self {
        TaskError::WinError(error)
    }
}
impl From<ComError> for TaskError {
    fn from(error: ComError) -> Self {
        TaskError::ComError(error)
    }
}
impl From<TaskServiceError> for TaskError {
    fn from(error: TaskServiceError) -> Self {
        TaskError::TaskServiceError(error)
    }
}
impl From<String> for TaskError {
    fn from(error: String) -> Self {
        TaskError::Error(error)
    }
}
#[derive(Debug)]
pub enum ComError {
    /// A previous call to CoInitializeEx specified the concurrency model
    /// for this thread as multithread apartment (MTA). This could also
    /// indicate that a change from neutral-threaded apartment to single-threaded
    /// apartment has occurred.
    RpcChangedMode,
    /// CoInitializeSecurity has already been called.
    RpcTooLate,
    /// The asAuthSvc parameter was not NULL, and none of the authentication
    /// services in the list could be registered. Check the results saved
    /// in asAuthSvc for authentication service–specific error codes.
    NoGoodSecurityPackages,
    /// A specified class is not registered in the registration database.
    /// Also can indicate that the type of server you requested in the CLSCTX
    ///enumeration is not registered or the values for the server types in the
    /// registry are corrupt.
    RegdbClassNotReg,
    /// This class cannot be created as part of an aggregate.
    ClassNoAggregation,
    /// The specified class does not implement the requested interface, or
    /// the controlling IUnknown does not expose the requested interface.
    NoInterface,
    GeneralError(WinError),
}

impl From<WinError> for ComError {
    fn from(error: WinError) -> Self {
        ComError::GeneralError(error)
    }
}

impl fmt::Display for ComError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            ComError::RpcChangedMode => {
                write!(f, "A previous call to CoInitializeEx specified the concurrency model for this thread as multithread apartment (MTA). This could also indicate that a change from neutral-threaded apartment to single-threaded apartment has occurred.")
            }
            ComError::RpcTooLate => {
                write!(f, "CoInitializeSecurity has already been called.")
            }
            ComError::NoGoodSecurityPackages => {
                write!(f, "The asAuthSvc parameter was not NULL, and none of the authentication services in the list could be registered. Check the results saved in asAuthSvc for authentication service–specific error codes.")
            }
            ComError::GeneralError(common_error) => {
                write!(f, "{}", common_error.to_string())
            }
            ComError::RegdbClassNotReg => {
                write!(f, "A specified class is not registered in the registration database. Also can indicate that the type of server you requested in the CLSCTX enumeration is not registered or the values for the server types in the registry are corrupt.")
            }
            ComError::ClassNoAggregation => {
                write!(f, "This class cannot be created as part of an aggregate.")
            }
            ComError::NoInterface => {
                write!(f, "The specified class does not implement the requested interface, or the controlling IUnknown does not expose the requested interface.")
            }
        }
    }
}
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
#[derive(Debug)]
pub enum TaskServiceError {
    /// Access is denied to connect to the Task Scheduler service.
    AccessDenied,
    /// The Task Scheduler service is not running.
    SchedulerServiceNotRunning,

    /// This error is returned in the following situations:
    /// The computer name specified in the serverName parameter does not exist.
    /// When you are trying to connect to a Windows Server 2003 or Windows XP computer, and the remote computer does not have the File and Printer Sharing firewall exception enabled or the Remote Registry service is not running.
    /// When you are trying to connect to a Windows Vista computer, and the remote computer does not have the Remote Scheduled Tasks Management firewall exception enabled and the File and Printer Sharing firewall exception enabled, or the Remote Registry service is not running.
    BadNetPath,
    /// The user, password, or domain parameters cannot be specified when connecting to a remote Windows XP or Windows Server 2003 computer from a Windows Vista computer.
    NotSupported,
    /// Common errors for windows [`WinError`]
    WinError(WinError),
    /// [`ComError`]
    ComError(ComError),
}
