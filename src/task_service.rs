use std::{convert::TryFrom, ptr};

use bindings::Windows::Win32::{
    Automation::BSTR,
    TaskScheduler::{self, ITaskDefinition, ITaskFolder},
};
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

    pub(crate) fn connect(&self) -> Result<(), windows::Error> {
        unsafe { self.0.Connect(None, None, None, None).ok() }
    }

    pub(crate) fn get_folder(&self) -> Result<ITaskFolder, windows::Error> {
        let mut task_folder = None;
        unsafe {
            let err = self
                .0
                .GetFolder(
                    BSTR::try_from("\\".to_string()).unwrap(),
                    // &None as *mut Option<ITaskFolder>,
                    &mut task_folder,
                )
                .ok();
            match err {
                Ok(_) => Ok(task_folder.unwrap()),
                Err(error) => Err(error),
            }
        }
        // let task_folder = task_folder.unwrap();
    }
    pub(crate) fn new_task(&self) -> Result<ITaskDefinition, windows::Error> {
        let mut task_definition = None;
        unsafe {
            let res = self.0.NewTask(0, &mut task_definition).ok();
            match res {
                Ok(_) => Ok(task_definition.unwrap()),
                Err(error) => Err(error),
            }
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
