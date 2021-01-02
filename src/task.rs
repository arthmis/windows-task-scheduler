use std::ptr;

use log::error;
use winapi::{shared::winerror::FAILED, um::taskschd::ITaskDefinition};

use crate::{
    error::WinError, principal::Principal, registration_info::RegistrationInfo,
    task_service::TaskService, task_settings::TaskSettings,
};

pub(crate) struct Task<'a> {
    pub(crate) task: &'a mut ITaskDefinition,
}

impl<'a> Task<'a> {
    /// Create a new task definition
    pub(crate) fn new(task_service: &TaskService) -> Result<Self, WinError> {
        unsafe {
            // Pass in a reference to a NULL ITaskDefinition interface pointer. Referencing a non-NULL pointer can cause a memory leak because the pointer will be overwritten.
            // The returned ITaskDefinition pointer must be released after it is used.
            let mut task: *mut ITaskDefinition = ptr::null_mut();
            // the flags argument is reserved for future use, so it will be 0
            let mut hr = task_service
                .task_service
                .NewTask(0, &mut task as *mut *mut ITaskDefinition);
            if FAILED(hr) {
                error!(
                    "Failed to CoCreate an instance of the TaskService class: {:X}",
                    hr
                );
                match hr {
                    E_POINTER => return Err(WinError::Pointer("NULL was passed in to the ppDefinition parameter. Pass in a reference to a NULL ITaskDefinition interface pointer.".to_string())),
                    E_INVALIDARG => return Err(WinError::InvalidArg("A nonzero value was passed into the flags paramter".to_string())),
                }
            }
            Ok(Self { task: &mut *task })
        }
    }

    // this shouldn't fail, I hope
    /// Gets or sets the registration information used to describe a task, such as the description of the task, the author of the task, and the date the task is registered.
    /// This property is read/write
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_registrationinfo
    pub(crate) fn get_registration_info(&self) -> RegistrationInfo {
        RegistrationInfo::new(self)
    }

    /// Gets or sets the principal for the task that provides the security credentials for the task.
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_principal
    pub(crate) fn get_principal(&self) -> Principal {
        Principal::new(self)
    }

    /// Gets or sets the settings that define how the Task Scheduler service performs the task.
    /// This property is read/write.
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_settings
    pub(crate) fn get_settings(&self) -> TaskSettings {
        TaskSettings::new(self)
    }
}

impl<'a> Task<'a> {
    pub(crate) unsafe fn as_ptr(&self) -> *const ITaskDefinition {
        self.task as *const ITaskDefinition
    }
}

impl<'a> Drop for Task<'a> {
    fn drop(&mut self) {
        unsafe {
            self.task.Release();
        }
    }
}
