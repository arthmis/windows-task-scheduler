#![allow(warnings)]
use combaseapi::CoCreateInstance;
use core::default::Default;
use iter::once;
use log::error;
use std::ffi::OsStr;
use std::{error::Error, fmt, path::Path, ptr, unreachable};
use std::{iter, os::windows::ffi::OsStrExt};
use winapi::{
    ctypes::c_void,
    shared::{winerror::FAILED, wtypesbase::CLSCTX_INPROC_SERVER},
    um::{
        combaseapi::{self},
        oaidl::VARIANT,
        taskschd::{ITaskFolder, ITaskService, TaskScheduler},
    },
    Class, Interface,
};
use winapi::{
    shared::wtypes::VARIANT_TRUE,
    um::taskschd::{
        IAction, IActionCollection, IExecAction, IIdleSettings, IPrincipal, IRegisteredTask,
        IRegistrationInfo, ITaskDefinition, ITaskSettings, ITimeTrigger, ITrigger,
        ITriggerCollection, TASK_ACTION_EXEC, TASK_CREATE_OR_UPDATE, TASK_LOGON_INTERACTIVE_TOKEN,
        TASK_TRIGGER_TIME,
    },
};
// use task_scheduler::Task;

mod com;
mod error;
mod task_service;

pub use chrono::DateTime;
pub use chrono::Duration;
pub use chrono::Utc;
use com::Com;
use task_service::TaskService;

/// Turns a string into a windows string
fn to_win_str(string: &str) -> Vec<u16> {
    OsStr::new(string).encode_wide().chain(once(0)).collect()
}

pub fn schedule_task(
    task_name: &str,
    task_path: &Path,
    start_time: DateTime<Utc>,
    end_time: DateTime<Utc>,
) {
    // validate that the task_path is not a folder and so on
    // validate that the end time is after the start times
    unsafe {
        let _com = Com::initialize().unwrap();

        // name for task
        let mut task_name = to_win_str(task_name);
        // path for notepad program
        let mut exe_path = to_win_str(task_path.to_str().unwrap());

        // Create an instance of the task service
        // this isn't properly documented, however these are pointers to GUIDs for these particular
        // classes. winapi has uuidof method to get the guid
        // I figured it out because of this https://docs.microsoft.com/en-us/archive/msdn-magazine/2007/october/windows-with-c-task-scheduler-2-0
        // here the got the uuid of the task scheduler using a function and the object
        let CLSID_TaskScheduler = Box::into_raw(Box::new(TaskScheduler::uuidof()));
        let IID_ITaskService = Box::into_raw(Box::new(ITaskService::uuidof()));
        let mut task_service: *mut ITaskService = core::ptr::null_mut();
        let mut hr = CoCreateInstance(
            CLSID_TaskScheduler,
            ptr::null_mut(),
            CLSCTX_INPROC_SERVER,
            IID_ITaskService,
            // have to figure out how this works
            &mut task_service as *mut *mut ITaskService as *mut *mut c_void,
        );

        if FAILED(hr) {
            println!("Failed to create an instance of ITaskService: {:X}", hr);
            return;
        }
        let task_service = task_service.as_mut().unwrap();

        // connect to the task service
        let variant: VARIANT = Default::default();
        hr = task_service.Connect(variant, variant, variant, variant);

        if FAILED(hr) {
            println!("ITaskService::Connect failed: {:X}", hr);
            task_service.Release();
            return;
        }

        // get the pointer to the root task folder. This folder will hold the
        // new task that is registered
        let mut root_task_folder: *mut ITaskFolder = core::ptr::null_mut();
        hr = task_service.GetFolder(
            to_win_str("\\").as_mut_ptr(),
            &mut root_task_folder as *mut *mut ITaskFolder,
        );
        if FAILED(hr) {
            println!("Cannot get root folder pointer: {:X}", hr);
            task_service.Release();
            return;
        }
        let root_task_folder = root_task_folder.as_mut().unwrap();

        // if the same task exists, remove it
        root_task_folder.DeleteTask(task_name.as_mut_ptr(), 0);

        // Create the task definition object to create the task
        let mut task: *mut ITaskDefinition = core::ptr::null_mut();
        hr = task_service.NewTask(0, &mut task as *mut *mut ITaskDefinition);
        if FAILED(hr) {
            println!(
                "Failed to CoCreate an instance of the TaskService class: {:X}",
                hr
            );
            root_task_folder.Release();
            return;
        }
        let task = task.as_mut().unwrap();
        // COM clean up. Pointer is no longer used
        task_service.Release();

        // Get the registration info for setting the identification
        let mut registration_info: *mut IRegistrationInfo = core::ptr::null_mut();
        hr = task.get_RegistrationInfo(&mut registration_info as *mut *mut IRegistrationInfo);

        if FAILED(hr) {
            println!("Cannot get identification pointer: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let mut registration_info = registration_info.as_mut().unwrap();
        hr = registration_info.put_Author(to_win_str("author name").as_mut_ptr());
        if FAILED(hr) {
            println!("Cannot put identification info: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }

        // Create the principal for the task - these credentials are
        // overwritten with the credentials passed to RegisterTaskDefinitio1n
        let mut principal: *mut IPrincipal = ptr::null_mut();
        hr = task.get_Principal(&mut principal as *mut *mut IPrincipal);
        if FAILED(hr) {
            println!("Cannot get principal pointer: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let mut principal = principal.as_mut().unwrap();

        // set up principal logon type to interactive logon
        hr = principal.put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
        if FAILED(hr) {
            println!("Cannot put principal info: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }

        // create the settings for the task
        let mut settings: *mut ITaskSettings = ptr::null_mut();
        hr = task.get_Settings(&mut settings as *mut *mut ITaskSettings);
        if FAILED(hr) {
            println!("Cannot get settings pointer: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }

        let settings = settings.as_mut().unwrap();

        // set settings values for the task
        hr = settings.put_StartWhenAvailable(VARIANT_TRUE);
        settings.Release();
        if FAILED(hr) {
            println!("Cannot put setting information: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }

        // set the idle settings for the task
        let mut idle_settings: *mut IIdleSettings = ptr::null_mut();
        hr = settings.get_IdleSettings(&mut idle_settings as *mut *mut IIdleSettings);
        if FAILED(hr) {
            println!("Cannot get idle setting information: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let idle_settings = idle_settings.as_mut().unwrap();

        idle_settings.put_WaitTimeout(to_win_str("PT5M").as_mut_ptr());
        idle_settings.Release();
        if FAILED(hr) {
            println!("Cannot put idle setting information: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }

        // get the trigger collection to insert the time trigger
        let mut trigger_collection: *mut ITriggerCollection = ptr::null_mut();
        hr = task.get_Triggers(&mut trigger_collection as *mut *mut ITriggerCollection);
        if FAILED(hr) {
            println!("Cannot get trigger collection: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let trigger_collection = trigger_collection.as_mut().unwrap();

        // add the time trigger to the task
        let mut trigger: *mut ITrigger = ptr::null_mut();
        hr = trigger_collection.Create(TASK_TRIGGER_TIME, &mut trigger as *mut *mut ITrigger);
        if FAILED(hr) {
            println!("Cannot create trigger: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let trigger = trigger.as_mut().unwrap();

        let mut time_trigger: *mut ITimeTrigger = ptr::null_mut();
        hr = trigger.QueryInterface(
            Box::into_raw(Box::new(ITimeTrigger::uuidof())),
            &mut time_trigger as *mut *mut ITimeTrigger as *mut *mut c_void,
        );
        trigger.Release();
        if FAILED(hr) {
            println!("QueryInterface call failed for ITimeTrigger: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let time_trigger = time_trigger.as_mut().unwrap();

        hr = time_trigger.put_Id(to_win_str("Trigger1").as_mut_ptr());
        if FAILED(hr) {
            println!("Cannot put trigger ID: {:X}", hr);
        }

        // hr = time_trigger.put_EndBoundary(to_win_str("2021-01-01T12:12:00").as_mut_ptr());
        hr = time_trigger.put_EndBoundary(to_win_str(&end_time.to_rfc3339()).as_mut_ptr());
        if FAILED(hr) {
            println!("Cannot put end boundary on trigger: {:X}", hr);
        }

        //  Set the task to start at a certain time. The time
        //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
        //  For example, the start boundary below
        //  is January 1st 2005 at 12:05
        // hr = time_trigger.put_StartBoundary(to_win_str("2021-01-01T12:11:00").as_mut_ptr());
        hr = time_trigger.put_StartBoundary(to_win_str(&start_time.to_rfc3339()).as_mut_ptr());
        time_trigger.Release();
        if FAILED(hr) {
            println!("Cannot add start boundary to trigger: {:x}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }

        // Add an action to the task. This task will execute notepad.exe
        let mut action_collection: *mut IActionCollection = ptr::null_mut();
        hr = task.get_Actions(&mut action_collection as *mut *mut IActionCollection);
        if FAILED(hr) {
            println!("Cannot get Task collection pointer: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let action_collection = &mut *action_collection;

        // create the action, specifying that it is an executable
        let mut action: *mut IAction = ptr::null_mut();
        action_collection.Create(TASK_ACTION_EXEC, &mut action as *mut *mut IAction);
        action_collection.Release();
        if FAILED(hr) {
            println!("Cannot create the action: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let action = &mut *action;

        let mut exec_action: *mut IExecAction = ptr::null_mut();
        // Query Interface for the executable task pointer
        hr = action.QueryInterface(
            Box::into_raw(Box::new(IExecAction::uuidof())),
            &mut exec_action as *mut *mut IExecAction as *mut *mut c_void,
        );
        if FAILED(hr) {
            println!("Cannot put action path: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }
        let exec_action = &mut *exec_action;

        // Set the path of the executable to notepad.exe
        hr = exec_action.put_Path(exe_path.as_mut_ptr());
        exec_action.Release();
        if FAILED(hr) {
            println!("Cannot put action path: {:X}", hr);
        }

        // Save the task in the root folder
        let mut registered_task: *mut IRegisteredTask = ptr::null_mut();
        hr = root_task_folder.RegisterTaskDefinition(
            task_name.as_mut_ptr(),
            task,
            TASK_CREATE_OR_UPDATE as i32,
            variant,
            variant,
            TASK_LOGON_INTERACTIVE_TOKEN,
            // this was supposed to be _variant_t(L"") in C++
            variant,
            registered_task as *mut *mut IRegisteredTask,
        );
        if FAILED(hr) {
            println!("Error saving the Task: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            return;
        }

        println!("Success! Task successfully registered");

        // Clean up
        root_task_folder.Release();
        task.Release();
        registered_task.as_mut().unwrap().Release();
    }
}

// #[cfg(test)]
// mod tests {
//     #[test]
//     fn it_works() {
//         assert_eq!(2 + 2, 4);
//     }
// }
