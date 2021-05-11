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

use crate::error::{ComError, TaskError, WinError};
/// These are potential errors when initializing com and its security levels

pub(crate) struct Com;

impl Com {
    // think about letting the user set the com multithreaded parameter
    // and implementing a default that chooses the single threaded version
    // I might split initialization and security initialization into
    // a builder function; for now this is fine
    #[allow(clippy::new_ret_no_self)]
    pub fn initialize() -> Result<Self, TaskError> {
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
                    _RPC_E_CHANGED_MODE => return Err(TaskError::from(ComError::RpcChangedMode)),
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
                    RPC_E_TOO_LATE => return Err(TaskError::from(ComError::RpcTooLate)),
                    RPC_E_NO_GOOD_SECURITY_PACKAGES => {
                        return Err(TaskError::from(ComError::NoGoodSecurityPackages))
                    }
                    E_OUT_OF_MEMORY => return Err(TaskError::from(WinError::OutOfMemory)),
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
