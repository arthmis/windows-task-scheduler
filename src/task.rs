use log::error;
use winapi::{shared::winerror::FAILED, um::taskschd::ITaskDefinition};

use crate::{error::WinError, task_service::TaskService};

pub(crate) struct Task<'a> {
    pub(crate) task: &'a mut ITaskDefinition,
}

impl<'a> Task<'a> {
    pub(crate) fn new(task_service: &TaskService) -> Result<Self, WinError> {
        unsafe {
            // Pass in a reference to a NULL ITaskDefinition interface pointer. Referencing a non-NULL pointer can cause a memory leak because the pointer will be overwritten.
            // The returned ITaskDefinition pointer must be released after it is used.
            let mut task: *mut ITaskDefinition = core::ptr::null_mut();
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
