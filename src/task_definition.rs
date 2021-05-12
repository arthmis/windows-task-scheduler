use std::ptr;

use bindings::Windows::Win32::TaskScheduler::{
    IPrincipal, IRegistrationInfo, ITaskDefinition, ITaskSettings, ITriggerCollection,
};
use log::error;

pub(crate) struct TaskDefinition(pub(crate) ITaskDefinition);

impl TaskDefinition {
    /// Create a new task definition
    pub(crate) fn new(task_definition: ITaskDefinition) -> Self {
        Self(task_definition)
    }

    // this shouldn't fail, I hope
    /// Gets or sets the registration information used to describe a task, such as the description of the task, the author of the task, and the date the task is registered.
    /// This property is read/write
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_registrationinfo
    pub(crate) fn get_registration_info(&self) -> Result<IRegistrationInfo, windows::Error> {
        let mut registration_info = None;
        unsafe {
            let res = self.0.get_RegistrationInfo(&mut registration_info).ok();
            match res {
                Ok(_) => Ok(registration_info.unwrap()),
                Err(error) => Err(error),
            }
        }
    }

    // /// Gets or sets the principal for the task that provides the security credentials for the task.
    // ///
    // /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_principal
    pub(crate) fn get_principal(&self) -> Result<IPrincipal, windows::Error> {
        let mut principal = None;
        unsafe {
            let res = self.0.get_Principal(&mut principal).ok();
            match res {
                Ok(_) => Ok(principal.unwrap()),
                Err(err) => Err(err),
            }
        }
    }

    // /// Gets or sets the settings that define how the Task Scheduler service performs the task.
    // /// This property is read/write.
    // ///
    // /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_settings
    pub(crate) fn get_settings(&self) -> Result<ITaskSettings, windows::Error> {
        unsafe {
            let mut task_settings = None;
            let res = self.0.get_Settings(&mut task_settings).ok();
            match res {
                Ok(_) => Ok(task_settings.unwrap()),
                Err(err) => Err(err),
            }
        }
    }

    /// Gets or sets a collection of triggers used to start a task.
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_triggers
    pub(crate) fn get_triggers(&self) -> Result<ITriggerCollection, windows::Error> {
        let mut trigger_collection = None;
        unsafe {
            let res = self.0.get_Triggers(&mut trigger_collection).ok();
            match res {
                Ok(_) => Ok(trigger_collection.unwrap()),
                Err(err) => Err(err),
            }
        }
    }
}
