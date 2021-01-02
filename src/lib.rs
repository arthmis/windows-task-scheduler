#![allow(warnings)]
use combaseapi::CoCreateInstance;
use core::default::Default;
use iter::once;
use log::error;
use principal::TaskLogon;
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
mod principal;
mod registration_info;
mod task;
mod task_folder;
mod task_service;
mod task_settings;

/// Re-exported from chrono for convenience
pub use chrono::DateTime;
/// Re-exported from chrono for convenience
pub use chrono::Duration;
/// Re-exported from chrono for convenience
pub use chrono::Utc;
/// Small wrapper over some of the com base apis
use com::Com;
/// Wrapper over ITaskService class
use task_service::TaskService;

/// Turns a string into a windows string
fn to_win_str(string: &str) -> Vec<u16> {
    OsStr::new(string).encode_wide().chain(once(0)).collect()
}

/// Use this function to schedule a task
/// For now the task will only be to start an executable
/// If the start time is not after the time this function is called
/// or the end time is before the start time then this function will fail
/// The task name can be anything you want, but it cannot start with a "."
pub fn schedule_task(
    task_name: &str,
    task_path: &Path,
    start_time: DateTime<Utc>,
    end_time: DateTime<Utc>,
) {
    // validate that the task_path is not a folder and so on
    // validate that the end time is after the start times
    let _com = Com::initialize().unwrap();

    // path for notepad program
    let mut exe_path = to_win_str(task_path.to_str().unwrap());

    // Create an instance of the task service
    let task_service = TaskService::new().unwrap();

    // Get the root task folder. This folder will hold the
    // new task that is registered
    let root_task_folder = task_service.get_folder().unwrap();

    // if the same task exists, remove it
    root_task_folder.delete_task(task_name).unwrap();

    // Create the task definition object to create the task
    let mut task = task_service.new_task().unwrap();

    // Get the registration info for setting the identification
    // and put the author
    let registration_info = task.get_registration_info();
    registration_info.put_author("author name");

    // Create the principal for the task - these credentials are
    // overwritten with the credentials passed to RegisterTaskDefinition
    let principal = task.get_principal();
    principal.put_logon_type(TaskLogon::InteractiveToken);

    // create the settings for the task
    let settings = task.get_settings();
    settings.put_start_when_available(VARIANT_TRUE);

    unsafe {
        // set the idle settings for the task
        let mut idle_settings: *mut IIdleSettings = ptr::null_mut();
        let mut hr = settings
            .settings
            .get_IdleSettings(&mut idle_settings as *mut *mut IIdleSettings);
        if FAILED(hr) {
            println!("Cannot get idle setting information: {:X}", hr);
            // root_task_folder.Release();
            // task.Release();
            return;
        }
        let idle_settings = idle_settings.as_mut().unwrap();

        idle_settings.put_WaitTimeout(to_win_str("PT5M").as_mut_ptr());
        idle_settings.Release();
        if FAILED(hr) {
            println!("Cannot put idle setting information: {:X}", hr);
            // root_task_folder.Release();
            // task.Release();
            return;
        }

        // get the trigger collection to insert the time trigger
        let mut trigger_collection: *mut ITriggerCollection = ptr::null_mut();
        hr = task
            .task
            .get_Triggers(&mut trigger_collection as *mut *mut ITriggerCollection);
        if FAILED(hr) {
            println!("Cannot get trigger collection: {:X}", hr);
            // root_task_folder.Release();
            // task.Release();
            return;
        }
        let trigger_collection = trigger_collection.as_mut().unwrap();

        // add the time trigger to the task
        let mut trigger: *mut ITrigger = ptr::null_mut();
        hr = trigger_collection.Create(TASK_TRIGGER_TIME, &mut trigger as *mut *mut ITrigger);
        if FAILED(hr) {
            println!("Cannot create trigger: {:X}", hr);
            // root_task_folder.Release();
            // task.Release();
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
            // root_task_folder.Release();
            // task.Release();
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
            // root_task_folder.Release();
            // task.Release();
            return;
        }

        // Add an action to the task. This task will execute notepad.exe
        let mut action_collection: *mut IActionCollection = ptr::null_mut();
        hr = task
            .task
            .get_Actions(&mut action_collection as *mut *mut IActionCollection);
        if FAILED(hr) {
            println!("Cannot get Task collection pointer: {:X}", hr);
            // root_task_folder.Release();
            // task.Release();
            return;
        }
        let action_collection = &mut *action_collection;

        // create the action, specifying that it is an executable
        let mut action: *mut IAction = ptr::null_mut();
        action_collection.Create(TASK_ACTION_EXEC, &mut action as *mut *mut IAction);
        action_collection.Release();
        if FAILED(hr) {
            println!("Cannot create the action: {:X}", hr);
            // root_task_folder.Release();
            // task.Release();
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
            // root_task_folder.Release();
            // task.Release();
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
        let variant = Default::default();
        let mut registered_task: *mut IRegisteredTask = ptr::null_mut();
        hr = root_task_folder.task_folder.RegisterTaskDefinition(
            to_win_str(task_name).as_mut_ptr(),
            // &mut task.task as *const ITaskDefinition,
            task.as_ptr(),
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
            // root_task_folder.Release();
            // task.Release();
            return;
        }

        println!("Success! Task successfully registered");

        // Clean up
        // root_task_folder.Release();
        // task.Release();
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
