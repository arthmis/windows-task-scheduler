use std::ptr;

use log::error;
use winapi::{
    ctypes::c_void,
    shared::{winerror::FAILED, wtypesbase::CLSCTX_INPROC_SERVER},
    um::{
        combaseapi::CoCreateInstance,
        oaidl::VARIANT,
        taskschd::{ITaskDefinition, ITaskFolder, ITaskService, TaskScheduler},
    },
    Class, Interface,
};

use crate::{com::ComError, error::WinError, task::Task, task_folder::TaskFolder, to_win_str};

// pub struct TaskService<'a> {
//     task_service: &'a mut ITaskService,
// }
pub(crate) struct TaskService<'a> {
    pub(crate) task_service: &'a mut ITaskService,
}

impl<'a> TaskService<'a> {
    pub(crate) fn new() -> Result<Self, TaskServiceError> {
        unsafe {
            // Create an instance of the task service
            // this isn't properly documented, however these are pointers to GUIDs for these particular
            // classes. winapi has uuidof method to get the guid
            // I figured it out because of this https://docs.microsoft.com/en-us/archive/msdn-magazine/2007/october/windows-with-c-task-scheduler-2-0
            let CLSID_TaskScheduler = Box::into_raw(Box::new(TaskScheduler::uuidof()));
            let IID_ITaskService = Box::into_raw(Box::new(ITaskService::uuidof()));
            let mut task_service: *mut ITaskService = core::ptr::null_mut();
            let hr = CoCreateInstance(
                CLSID_TaskScheduler,
                ptr::null_mut(),
                CLSCTX_INPROC_SERVER,
                IID_ITaskService,
                // have to figure out how this pointer casting works
                &mut task_service as *mut *mut ITaskService as *mut *mut c_void,
            );

            if FAILED(hr) {
                error!("Failed to create an instance of TaskService: {:X}", hr);
                match hr {
                    REGDB_E_CLASSNOTREG => {
                        return Err(TaskServiceError::ComError(ComError::RegdbClassNotReg))
                    }
                    CLASS_E_NOAGGREGATION => {
                        return Err(TaskServiceError::ComError(ComError::ClassNoAggregation))
                    }
                    E_NOINTERFACE => return Err(TaskServiceError::ComError(ComError::NoInterface)),
                    E_POINTER => {
                        return Err(TaskServiceError::WinError(WinError::Pointer(
                            "The ppv parameter is NULL".to_string(),
                        )))
                    }
                }
            }
            let task_service = Self {
                // I can safely dereference this pointer here because I handle the
                // error where the pointer might be null
                task_service: &mut *task_service,
            };
            match task_service.connect() {
                Ok(_) => Ok(task_service),
                Err(error) => Err(error),
            }
            // Ok(task_service)
        }
    }

    // fn connect(_server_name: VARIANT, _user: VARIANT, domain: VARIANT, _password: VARIANT) {}
    // the above is the actual function signature i want to eventually have
    // TODO figure out how the VARIANT union works. Variant is actually a union
    /// Connects to the task service
    ///
    /// This will use a zeroed union by default for the actual api. In the future
    /// it will hopefully be possible to accept the appropriate parameters.
    #[must_use = "This function needs to be called before the other methods can be used"]
    fn connect(&self) -> Result<(), TaskServiceError> {
        // connect to the task service
        let variant: VARIANT = Default::default();
        unsafe {
            let hr = self
                .task_service
                .Connect(variant, variant, variant, variant);

            if FAILED(hr) {
                error!("ITaskService::Connect failed: {:X}", hr);
                match hr {
                    E_ACCESS_DENIED => return Err(TaskServiceError::AccessDenied),
                    SCHED_E_SERVICE_NOT_RUNNING => {
                        return Err(TaskServiceError::SchedulerServiceNotRunning)
                    }
                    ERROR_BAD_NET_PATH => return Err(TaskServiceError::BadNetPath),
                    ERROR_NOT_SUPPORTED => return Err(TaskServiceError::NotSupported),
                    E_OUT_OF_MEMORY => {
                        return Err(TaskServiceError::WinError(WinError::OutOfMemory))
                    }
                }
            }
        }
        Ok(())
    }

    /// This doesn't really specify what errors can be returned
    /// The path to the folder to retrieve. Do not use a backslash following the last folder name in the path. The root task folder is specified with a backslash (). An example of a task folder path, under the root task folder, is \MyTaskFolder. The '.' character cannot be used to specify the current task folder and the '..' characters cannot be used to specify the parent task folder in the path.
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskservice-getfolder
    // I'll have to figure what do with it in the future
    pub(crate) fn get_folder(&self) -> Result<TaskFolder, String> {
        TaskFolder::new(self)
    }

    /// Returns an empty task definition object to be filled in with settings
    // and properties and then registered using TaskFolder::register_task_definition
    pub(crate) fn new_task(&self) -> Result<Task, WinError> {
        Task::new(self)
    }
}

impl<'a> Drop for TaskService<'a> {
    fn drop(&mut self) {
        unsafe {
            self.task_service.Release();
        }
    }
}
#[derive(Debug)]
pub enum TaskServiceError {
    /// Access is denied to connect to the Task Scheduler service.
    AccessDenied,
    /// The Task Scheduler service is not running.
    SchedulerServiceNotRunning,

    /// This error is returned in the following situations:
    /// The computer name specified in the serverName parameter does not exist.
    /// When you are trying to connect to a Windows Server 2003 or Windows XP computer, and the remote computer does not have the File and Printer Sharing firewall exception enabled or the Remote Registry service is not running.
    /// When you are trying to connect to a Windows Vista computer, and the remote computer does not have the Remote Scheduled Tasks Management firewall exception enabled and the File and Printer Sharing firewall exception enabled, or the Remote Registry service is not running.
    BadNetPath,
    /// The user, password, or domain parameters cannot be specified when connecting to a remote Windows XP or Windows Server 2003 computer from a Windows Vista computer.
    NotSupported,
    /// Common errors for windows [`WinError`]
    WinError(WinError),
    /// [`ComError`]
    ComError(ComError),
}
