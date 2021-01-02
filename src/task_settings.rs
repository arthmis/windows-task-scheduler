use std::ptr;

use log::error;
use winapi::{
    shared::{
        winerror::FAILED,
        wtypes::{VARIANT_BOOL, VARIANT_TRUE},
    },
    um::taskschd::ITaskSettings,
};

use crate::task::Task;

/// Provides the settings that the Task Scheduler service uses to perform the task.
pub(crate) struct TaskSettings<'a> {
    pub(crate) settings: &'a mut ITaskSettings,
}

impl<'a> TaskSettings<'a> {
    /// Creates TaskSettings to define settings for Task Scheduler.
    pub(crate) fn new(task: &Task) -> Self {
        unsafe {
            // create the settings for the task
            let mut settings: *mut ITaskSettings = ptr::null_mut();
            let mut hr = task
                .task
                .get_Settings(&mut settings as *mut *mut ITaskSettings);
            if FAILED(hr) {
                error!("Cannot get settings pointer: {:X}", hr);
                unreachable!();
            }

            Self {
                settings: &mut *settings,
            }
        }
    }

    // TODO figure VARIANT_BOOL
    // I have to figure out what that even means and how to translate it
    // to Rust
    /// TODO documentation
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itasksettings-put_startwhenavailable
    pub(crate) fn put_start_when_available(&self, variant: VARIANT_BOOL) {
        unsafe {
            let hr = self.settings.put_StartWhenAvailable(VARIANT_TRUE);
            if FAILED(hr) {
                error!("Cannot put setting information: {:X}", hr);
                unreachable!();
            }
        }
    }
}

impl<'a> Drop for TaskSettings<'a> {
    fn drop(&mut self) {
        unsafe {
            self.settings.Release();
        }
    }
}
