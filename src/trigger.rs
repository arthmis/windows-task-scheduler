use std::{ffi::c_void, ptr};

use log::error;
use winapi::{
    shared::winerror::FAILED,
    um::taskschd::{ITimeTrigger, ITrigger, TASK_TRIGGER_TIME},
    Interface,
};

use crate::{
    error::WinError,
    time_trigger::TimeTrigger,
    trigger_collection::{TaskTriggerType, TriggerCollection},
};

/// Provides the common properties that are inherited by all trigger objects.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-itrigger
pub(crate) struct TriggerInterface<'a> {
    trigger: &'a mut ITrigger,
}

impl<'a> TriggerInterface<'a> {
    pub(crate) fn query_interface(&self) -> Result<TimeTrigger, WinError> {
        unsafe {
            let mut time_trigger: *mut ITimeTrigger = ptr::null_mut();
            let mut hr = self.trigger.QueryInterface(
                Box::into_raw(Box::new(ITimeTrigger::uuidof())),
                &mut time_trigger as *mut *mut ITimeTrigger as *mut *mut c_void,
            );
            if FAILED(hr) {
                error!("QueryInterface call failed for ITimeTrigger: {:X}", hr);
                // root_task_folder.Release();
                return Err(WinError::UnknownError(format!(
                    "QueryInterface call failed for ITimeTrigger: {:X}",
                    hr
                )));
            }
            Ok(TimeTrigger::new(time_trigger))
        }
    }
}
impl<'a> TimeTrigger<'a> {
    unsafe fn new(time_trigger: *mut ITimeTrigger) -> Self {
        Self {
            time_trigger: &mut *time_trigger,
        }
    }
}

impl<'a> Drop for TriggerInterface<'a> {
    fn drop(&mut self) {
        unsafe {
            self.trigger.Release();
        }
    }
}
