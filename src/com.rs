use std::{fmt, ptr};

use log::error;
use winapi::{
    shared::{
        rpcdce::{RPC_C_AUTHN_LEVEL_PKT_PRIVACY, RPC_C_IMP_LEVEL_IMPERSONATE},
        winerror::FAILED,
    },
    um::{
        combaseapi::{self, CoInitializeSecurity, CoUninitialize},
        objbase,
    },
};

use crate::error::WinError;
/// These are potential errors when initializing com and its security levels
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

pub(crate) struct Com;

impl Com {
    // think about letting the user set the com multithreaded parameter
    // and implementing a default that chooses the single threaded version
    // I might split initialization and security initialization into
    // a builder function; for now this is fine
    #[allow(clippy::new_ret_no_self)]
    pub fn initialize() -> Result<Self, ComError> {
        unsafe {
            // pvReserved is a reserved paramter and must be null
            let hr = combaseapi::CoInitializeEx(
                // pvReserved is a reserved paramter and must be null
                ptr::null_mut(),
                // The concurrency model and initialization options for the thread. Values for this parameter are taken from the COINIT enumeration. Any combination of values from COINIT can be used, except that the COINIT_APARTMENTTHREADED and COINIT_MULTITHREADED flags cannot both be set. The default is COINIT_MULTITHREADED.
                objbase::COINIT_APARTMENTTHREADED,
            );
            if FAILED(hr) {
                error!("Com initialization failed: {:X}", hr);
                match hr {
                    _RPC_E_CHANGED_MODE => return Err(ComError::RpcChangedMode),
                    _ => unreachable!(),
                }
            }

            // set general COM security levels
            let hr = CoInitializeSecurity(
                ptr::null_mut(),
                -1,
                ptr::null_mut(),
                ptr::null_mut(),
                RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
                RPC_C_IMP_LEVEL_IMPERSONATE,
                ptr::null_mut(),
                0,
                // this must be null, it is reserved parameter
                ptr::null_mut(),
            );
            if FAILED(hr) {
                error!("Com security initialization failed: {:X}", hr);
                match hr {
                    RPC_E_TOO_LATE => return Err(ComError::RpcTooLate),
                    RPC_E_NO_GOOD_SECURITY_PACKAGES => {
                        return Err(ComError::NoGoodSecurityPackages)
                    }
                    E_OUT_OF_MEMORY => return Err(ComError::GeneralError(WinError::OutOfMemory)),
                }
            }
        }
        Ok(Self)
    }
}
impl Drop for Com {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}
