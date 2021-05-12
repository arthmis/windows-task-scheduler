use std::ptr;

use bindings::Windows::Win32::TaskScheduler;
use bindings::Windows::Win32::{
    Com::{self, CoCreateInstance},
    TaskScheduler::ITaskService,
};
use log::error;
use windows::IntoParam;
use windows::{Guid, Interface};

use crate::{
    error::{ComError, TaskError, TaskServiceError, WinError},
    to_win_str,
};

pub(crate) struct TaskService(pub(crate) ITaskService);
// pub(crate) task_service: &'a mut ITaskService,

impl TaskService {
    pub(crate) fn new() -> Self {
        unsafe {
            // Create an instance of the task service
            // this isn't properly documented, however these are pointers to GUIDs for these particular
            // classes. winapi has uuidof method to get the guid
            // I figured it out because of this https://docs.microsoft.com/en-us/archive/msdn-magazine/2007/october/windows-with-c-task-scheduler-2-0
            let mut task_service: ITaskService = CoCreateInstance(
                &TaskScheduler::TaskScheduler as *const Guid,
                // TaskScheduler::ITaskScheduler,
                None,
                Com::CLSCTX::CLSCTX_INPROC_SERVER,
            )
            .unwrap();
            Self(task_service)
        }
    }
}

// impl Drop for TaskService {
//     fn drop(&mut self) {
//         unsafe {
//             self.0.Release();
//         }
//     }
// }
