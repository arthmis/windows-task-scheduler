use std::ptr;

use bindings::Windows::Win32::TaskScheduler::{IIdleSettings, ITaskSettings};
use log::error;

/// Provides the settings that the Task Scheduler service uses to perform the task.
pub(crate) struct TaskSettings(pub(crate) ITaskSettings);

impl TaskSettings {
    /// Creates TaskSettings to define settings for Task Scheduler.
    pub(crate) fn new(task_settings: ITaskSettings) -> Self {
        Self(task_settings)
    }

    // TODO figure VARIANT_BOOL
    // I have to figure out what that even means and how to translate it
    // to Rust
    /// TODO documentation
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itasksettings-put_startwhenavailable
    pub(crate) fn put_start_when_available(&self, variant: i16) -> Result<(), windows::Error> {
        unsafe { self.0.put_StartWhenAvailable(variant).ok() }
    }

    /// Gets or sets the information that specifies how the Task Scheduler performs tasks when the computer is in an idle condition. For information about idle conditions, see Task Idle Conditions.
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itasksettings-get_idlesettings
    pub(crate) fn get_idle_settings(&self) -> Result<IIdleSettings, windows::Error> {
        unsafe {
            let mut idle_settings = None;
            let res = self.0.get_IdleSettings(&mut idle_settings).ok();
            match res {
                Ok(_) => Ok(idle_settings.unwrap()),
                Err(err) => Err(err),
            }
        }
    }
}
