use std::ptr;

use bindings::Windows::Win32::TaskScheduler::IIdleSettings;
use chrono::Duration;
use log::error;
// use winapi::{shared::winerror::FAILED, um::taskschd::IIdleSettings};

use crate::{error::WinError, task_settings::TaskSettings, to_win_str};

/// Specifies how the Task Scheduler performs tasks when the computer is in an idle condition. For information about idle conditions, see Task Idle Conditions.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-iidlesettings
pub(crate) struct IdleSettings<'a> {
    idle_settings: &'a mut IIdleSettings,
}

impl<'a> IdleSettings<'a> {
    /// Creates IdleSettings from TaskSettings which will allow setting task scheduler
    /// behavior when the system is idle
    pub(crate) fn new(settings: &TaskSettings) -> Result<Self, WinError> {
        unsafe {
            // set the idle settings for the task
            let mut idle_settings: *mut IIdleSettings = ptr::null_mut();
            let mut hr = settings
                .settings
                .get_IdleSettings(&mut idle_settings as *mut *mut IIdleSettings);
            if hr.is_err() {
                error!("Cannot get idle settings information: {:X}", hr);
                return Err(WinError::UnknownError(format!(
                    "Cannot get idle settings information: {:X}",
                    hr
                )));
            }
            Ok(Self {
                idle_settings: &mut *idle_settings,
            })
        }
    }

    /// Gets or sets a value that indicates the amount of time that the Task Scheduler will wait for an idle condition to occur. If no value is specified for this property, then the Task Scheduler service will wait indefinitely for an idle condition to occur.
    /// This will wait 5 minutes as of right now but will be customizable in the future
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-iidlesettings-put_waittimeout
    // TODO: eventually this should take the duration and translate it to the string
    // time they use which is found on https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-waittimeout-idlesettingstype-element
    pub(crate) fn put_wait_timeout(&self, _wait_time: Duration) -> Result<(), WinError> {
        unsafe {
            let hr = self
                .idle_settings
                .put_WaitTimeout(to_win_str("PT5M").as_mut_ptr());
            if FAILED(hr) {
                error!("Cannot put idle setting information: {:X}", hr);
                return Err(WinError::UnknownError(format!(
                    "Cannot put idle setting information: {:X}",
                    hr
                )));
            }
        }
        Ok(())
    }
}

impl<'a> Drop for IdleSettings<'a> {
    fn drop(&mut self) {
        unsafe {
            self.idle_settings.Release();
        }
    }
}
