use std::ptr;

use bindings::Windows::Win32::{Automation::BSTR, TaskScheduler::IIdleSettings};
use chrono::Duration;
use log::error;
// use winapi::{shared::winerror::FAILED, um::taskschd::IIdleSettings};

use crate::{error::WinError, task_settings::TaskSettings, to_win_str};

/// Specifies how the Task Scheduler performs tasks when the computer is in an idle condition. For information about idle conditions, see Task Idle Conditions.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-iidlesettings
pub(crate) struct IdleSettings(pub(crate) IIdleSettings);

impl IdleSettings {
    /// Creates IdleSettings from TaskSettings which will allow setting task scheduler
    /// behavior when the system is idle
    pub(crate) fn new(settings: IIdleSettings) -> Self {
        Self(settings)
    }

    /// Gets or sets a value that indicates the amount of time that the Task Scheduler will wait for an idle condition to occur. If no value is specified for this property, then the Task Scheduler service will wait indefinitely for an idle condition to occur.
    /// This will wait 5 minutes as of right now but will be customizable in the future
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-iidlesettings-put_waittimeout
    // TODO: eventually this should take the duration and translate it to the string
    // time they use which is found on https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-waittimeout-idlesettingstype-element
    pub(crate) fn put_wait_timeout(&self, _wait_time: Duration) -> Result<(), windows::Error> {
        unsafe { self.0.put_WaitTimeout(BSTR::from("PT5M")).ok() }
    }
}
