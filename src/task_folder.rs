use std::convert::TryFrom;

use bindings::Windows::Win32::{
    Automation::BSTR,
    TaskScheduler::{
        IRegisteredTask, ITaskDefinition, ITaskFolder, TASK_CREATION, TASK_LOGON_TYPE,
    },
};
use log::error;

use crate::{task_service::TaskService, to_win_str};

pub(crate) struct TaskFolder(pub(crate) ITaskFolder);

impl TaskFolder {
    /// Gets a folder of registered tasks
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskservice-getfolder
    pub(crate) fn new(task_folder: ITaskFolder) -> Self {
        Self(task_folder)
    }
    /// Deletes a task from the folder
    ///
    /// The task name is the name that was specified when the task was registered
    /// The '.' cannot be used to specify the current task folder and the '..'
    /// cannot be used to specify the parent task folder in the path
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskfolder-deletetask
    pub(crate) fn delete_task(&self, task_name: &str) -> Result<(), windows::Error> {
        const flags: i32 = 0;
        unsafe {
            self.0
                .DeleteTask(BSTR::try_from(task_name).unwrap(), flags)
                .ok();
        }
        Ok(())
    }

    pub(crate) fn register_task(
        &self,
        task_name: &str,
        task_definition: ITaskDefinition,
    ) -> Result<IRegisteredTask, windows::Error> {
        let mut registered_task = None;
        unsafe {
            let err = self
                .0
                .RegisterTaskDefinition(
                    BSTR::from(task_name),
                    task_definition.clone(),
                    TASK_CREATION::TASK_CREATE_OR_UPDATE.0,
                    None,
                    None,
                    TASK_LOGON_TYPE::TASK_LOGON_INTERACTIVE_TOKEN,
                    None,
                    &mut registered_task,
                )
                .ok();
            match err {
                Ok(_) => Ok(registered_task.unwrap()),
                Err(error) => Err(error),
            }
        }
    }
}
