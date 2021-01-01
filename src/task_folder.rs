use log::error;
use winapi::{shared::winerror::FAILED, um::taskschd::ITaskFolder};

use crate::to_win_str;

pub(crate) struct TaskFolder<'a> {
    pub(crate) task_folder: &'a mut ITaskFolder,
}

impl<'a> TaskFolder<'a> {
    /// This should only be used with TaskService::get_folder()
    pub(crate) unsafe fn new(folder: *mut ITaskFolder) -> Self {
        // get the pointer to the root task folder. This folder will hold the
        // new task that is registered
        // let mut root_task_folder: *mut ITaskFolder = core::ptr::null_mut();
        // let hr = task_service.GetFolder(
        //     to_win_str("\\").as_mut_ptr(),
        //     &mut root_task_folder as *mut *mut ITaskFolder,
        // );
        // if FAILED(hr) {
        //     println!("Cannot get root folder pointer: {:X}", hr);
        //     task_service.Release();
        //     return;
        // }
        let task_folder = folder.as_mut().unwrap();
        Self { task_folder }
    }
    /// Deletes a task from the folder
    ///
    /// The task name is the name that was specified when the task was registered
    /// The '.' cannot be used to specify the current task folder and the '..'
    /// cannot be used to specify the parent task folder in the path
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskfolder-deletetask
    pub(crate) fn delete_task(&self, task_name: &str) -> Result<(), String> {
        const flags: i32 = 0;
        let mut task_name_win_str = to_win_str(task_name);
        unsafe {
            // flags are not supported so it will be 0
            let hr = self
                .task_folder
                .DeleteTask(task_name_win_str.as_mut_ptr(), flags);
            if FAILED(hr) {
                error!(
                    "There was an issue deleting task with name: {}\nError code was: {:X}",
                    task_name, hr
                );
                return Err(format!(
                    "There was an issue deleting tas with name: {}\nError code was: {:X}",
                    task_name, hr
                ));
            }
        }
        Ok(())
    }
}

impl<'a> Drop for TaskFolder<'a> {
    fn drop(&mut self) {
        unsafe {
            self.task_folder.Release();
        }
    }
}
