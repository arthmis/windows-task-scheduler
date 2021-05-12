#![allow(warnings)]
use bindings::Windows::Win32::{
    Automation::BSTR,
    TaskScheduler::{
        IExecAction, ITaskFolder, ITaskService, ITimeTrigger, TASK_ACTION_TYPE, TASK_CREATION,
        TASK_LOGON_TYPE, TASK_TRIGGER_TYPE, TASK_TRIGGER_TYPE2,
    },
};
use bindings::Windows::Win32::{Automation::VARIANT, Com::CoCreateInstance};
use core::default::Default;
use error::{TaskError, WinError};
use iter::once;
use log::error;
use windows::{Interface, IntoParam};
// use principal::TaskLogon;
use std::{convert::TryFrom, error::Error, fmt, path::Path, ptr, unreachable};
use std::{ffi::OsStr, path::PathBuf};
use std::{iter, os::windows::ffi::OsStrExt};
// use trigger_collection::{TaskTriggerType, TriggerCollection};

// use task_scheduler::Task;

mod com;
mod error;
// mod idle_settings;
// mod principal;
// mod registration_info;
// mod task;
// mod task_folder;
mod task_service;
// mod task_settings;
// mod trigger_collection;

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

// pub struct Task {
//     exe_path: PathBuf,
//     task_name: String,
//     triggers: TaskTriggers,
// }

// impl Task {
pub fn execute(task_path: PathBuf, task_name: &str) {
    let _com = Com::initialize().unwrap();

    let task_service = TaskService::new();
    unsafe {
        task_service.0.Connect(None, None, None, None).unwrap();

        // task_service.0.GetFolder(to_win_str("\\").as_mut_ptr())
        // let mut task_folder: *mut ITaskFolder = std::ptr::null_mut();
        let mut task_folder = None;
        task_service
            .0
            .GetFolder(
                BSTR::try_from("\\".to_string()).unwrap(),
                // &None as *mut Option<ITaskFolder>,
                &mut task_folder,
            )
            .unwrap();
        let task_folder = task_folder.unwrap();

        // delete tasks if it exists
        let mut err = task_folder.DeleteTask(BSTR::try_from(task_name).unwrap(), 0);
        println!("{}", err.message());

        let mut task_definition = None;
        task_service.0.NewTask(0, &mut task_definition).unwrap();
        let task_definition = task_definition.unwrap();

        let mut registration_info = None;
        task_definition
            .get_RegistrationInfo(&mut registration_info)
            .unwrap();
        let registration_info = registration_info.unwrap();
        registration_info.put_Author(BSTR::from("Author")).unwrap();

        let mut principal = None;
        task_definition.get_Principal(&mut principal).unwrap();
        let principal = principal.unwrap();

        principal
            .put_LogonType(TASK_LOGON_TYPE::TASK_LOGON_INTERACTIVE_TOKEN)
            .unwrap();

        let mut task_settings = None;
        task_definition.get_Settings(&mut task_settings).unwrap();
        let task_settings = task_settings.unwrap();

        const VARIANT_TRUE: i16 = -1;
        task_settings.put_StartWhenAvailable(VARIANT_TRUE).unwrap();

        let mut idle_settings = None;
        task_settings.get_IdleSettings(&mut idle_settings).unwrap();
        let idle_settings = idle_settings.unwrap();

        idle_settings.put_WaitTimeout(BSTR::from("PT5M")).unwrap();

        let mut trigger_collection = None;
        task_definition.get_Triggers(&mut trigger_collection);
        let trigger_collection = trigger_collection.unwrap();

        let mut trigger = None;
        trigger_collection
            .Create(TASK_TRIGGER_TYPE2::TASK_TRIGGER_TIME, &mut trigger)
            .unwrap();
        let trigger = trigger.unwrap();

        let time_trigger = trigger.cast::<ITimeTrigger>().unwrap();
        // for trigger in time_triggers {
        //     let trigger_collection = TriggerCollection::new(&task_definition)?;
        //     let trigger = trigger_collection.create(TaskTriggerType::SpecificTime(trigger))?;
        // }
        time_trigger.put_Id(BSTR::from("Trigger1")).unwrap();
        let start = Utc::now() + chrono::Duration::seconds(10);
        let end = Utc::now() + chrono::Duration::seconds(60);
        time_trigger
            .put_EndBoundary(BSTR::from(end.to_rfc3339()))
            .unwrap();
        time_trigger
            .put_StartBoundary(BSTR::from(start.to_rfc3339()))
            .unwrap();

        let mut action_collection = None;
        task_definition.get_Actions(&mut action_collection).unwrap();
        let action_collection = action_collection.unwrap();

        let mut action = None;
        action_collection
            .Create(TASK_ACTION_TYPE::TASK_ACTION_EXEC, &mut action)
            .unwrap();
        let action = action.unwrap();

        let exec_action = action.cast::<IExecAction>().unwrap();
        exec_action
            .put_Path(BSTR::from(task_path.to_str().unwrap()))
            .unwrap();

        let mut registered_task = None;
        let err = task_folder.RegisterTaskDefinition(
            BSTR::from(task_name),
            task_definition.clone(),
            TASK_CREATION::TASK_CREATE_OR_UPDATE.0,
            None,
            None,
            TASK_LOGON_TYPE::TASK_LOGON_INTERACTIVE_TOKEN,
            None,
            &mut registered_task,
        );
        println!("{}", err.message());
    }

    // path for notepad program
    // let mut exe_path = to_win_str(actions.0[0].to_str().unwrap());
}

/// Use this function to schedule a task
/// For now the task will only be to start an executable
/// If the start time is not after the time this function is called
/// or the end time is before the start time then this function will fail
/// The task name can be anything you want, but it cannot start with a "."
pub fn execute_task(
    actions: Actions,
    task_name: String,
    triggers: TaskTriggers,
) -> Result<(), TaskError> {
    dbg!(&actions.0[0], &task_name);
    // validate that the task_path is not a folder and so on
    // validate that the end time is after the start times

    let _com = Com::initialize()?;

    // path for notepad program
    let mut exe_path = to_win_str(actions.0[0].to_str().unwrap());

    // Create an instance of the task service
    // let task_service = TaskService::new()?;

    // Get the root task folder. This folder will hold the
    // new task that is registered
    // let root_task_folder = task_service.get_folder()?;

    // // if the same task exists, remove it
    // // root_task_folder.delete_task(&task_name).unwrap();

    // // Create the task definition object to create the task
    // let mut task_definition = task_service.new_task()?;

    // // Get the registration info for setting the identification
    // // and put the author
    // let registration_info = task_definition.get_registration_info();
    // registration_info.put_author("author name");

    // // Create the principal for the task - these credentials are
    // // overwritten with the credentials passed to RegisterTaskDefinition
    // let principal = task_definition.get_principal();
    // principal.put_logon_type(TaskLogon::InteractiveToken);

    // // create the settings for the task
    // let settings = task_definition.get_settings();
    // settings.put_start_when_available(VARIANT_TRUE);

    // // set the idle settings for the task
    // let idle_settings = settings.get_idle_settings()?;
    // idle_settings.put_wait_timeout(Duration::minutes(5));

    // let trigger_collection = task_definition.get_triggers()?;
    // // let trigger_collection = TriggerCollection::new(&task_definition).unwrap();
    // // let trigger = trigger_collection
    // //     .create(TaskTriggerType::TaskTriggerTime)
    // //     .unwrap();

    // unsafe {
    //     // sets the daily triggers
    //     if let Some(daily_triggers) = triggers.daily {
    //         for daily_trigger in daily_triggers {
    //             let trigger_collection = TriggerCollection::new(&task_definition)?;
    //             let trigger = trigger_collection.create(TaskTriggerType::Daily(daily_trigger))?;
    //         }
    //     }

    //     // sets triggers that happen at specific times
    //     if let Some(time_triggers) = triggers.time {
    //         for trigger in time_triggers {
    //             let trigger_collection = TriggerCollection::new(&task_definition)?;
    //             let trigger = trigger_collection.create(TaskTriggerType::SpecificTime(trigger))?;
    //         }
    //     }

    //     // Add an action to the task. This task will execute notepad.exe
    //     let mut action_collection: *mut IActionCollection = ptr::null_mut();
    //     let mut hr = task_definition
    //         .task
    //         .get_Actions(&mut action_collection as *mut *mut IActionCollection);
    //     if FAILED(hr) {
    //         println!("Cannot get Task collection pointer: {:X}", hr);
    //         // root_task_folder.Release();
    //         // task.Release();
    //         // return;
    //     }
    //     let action_collection = &mut *action_collection;

    //     // create the action, specifying that it is an executable
    //     let mut action: *mut IAction = ptr::null_mut();
    //     action_collection.Create(TASK_ACTION_EXEC, &mut action as *mut *mut IAction);
    //     action_collection.Release();
    //     if FAILED(hr) {
    //         println!("Cannot create the action: {:X}", hr);
    //         return Err(TaskError::from(WinError::UnknownError(format!(
    //             "Cannot create the action: {:X}",
    //             hr
    //         ))));
    //     }
    //     let action = &mut *action;

    //     let mut exec_action: *mut IExecAction = ptr::null_mut();
    //     // Query Interface for the executable task pointer
    //     hr = action.QueryInterface(
    //         Box::into_raw(Box::new(IExecAction::uuidof())),
    //         &mut exec_action as *mut *mut IExecAction as *mut *mut c_void,
    //     );
    //     if FAILED(hr) {
    //         println!("Cannot get IExecAction interface: {:X}", hr);
    //         return Err(TaskError::from(WinError::UnknownError(format!(
    //             "Cannot get IExecAction interface: {:X}",
    //             hr
    //         ))));
    //     }
    //     let exec_action = &mut *exec_action;

    //     // Set the path of the executable to notepad.exe
    //     hr = exec_action.put_Path(exe_path.as_mut_ptr());
    //     exec_action.Release();
    //     if FAILED(hr) {
    //         println!("Cannot put action path: {:X}", hr);
    //         return Err(TaskError::from(WinError::UnknownError(format!(
    //             "Cannot put action path: {:X}",
    //             hr
    //         ))));
    //     }

    //     // Save the task in the root folder
    //     let variant = Default::default();
    //     let mut registered_task: *mut IRegisteredTask = ptr::null_mut();
    //     hr = root_task_folder.task_folder.RegisterTaskDefinition(
    //         to_win_str(&task_name).as_mut_ptr(),
    //         // &mut task.task as *const ITaskDefinition,
    //         task_definition.as_ptr(),
    //         TASK_CREATE_OR_UPDATE as i32,
    //         variant,
    //         variant,
    //         TASK_LOGON_INTERACTIVE_TOKEN,
    //         // this was supposed to be _variant_t(L"") in C++
    //         variant,
    //         registered_task as *mut *mut IRegisteredTask,
    //     );
    //     if FAILED(hr) {
    //         println!("Error saving the Task: {:X}", hr);
    //         // root_task_folder.Release();
    //         // task.Release();
    //         // return;
    //     }

    //     println!("Success! Task successfully registered");

    //     // Clean up
    //     // root_task_folder.Release();
    //     // task.Release();
    //     registered_task.as_mut().unwrap().Release();
    // }

    Ok(())
}

#[derive(Debug)]
pub struct Actions(Vec<PathBuf>);

/// There can only be up to 32 actions per task
impl Actions {
    pub fn new(path: PathBuf) -> Self {
        if !&path.is_file() {
            // return Err(TaskError::Error(
            //     "The path does not point to a file".to_string(),
            // ));
            panic!("This should absolutely be a file");
        }
        Self(vec![path])
    }
}

#[derive(Debug)]
pub struct TaskTriggers {
    daily: Option<Vec<DailyTrigger>>,
    event: Option<Vec<EventTrigger>>,
    idle: Option<IdleTrigger>,
    registration: Option<RegistrationTrigger>,
    time: Option<Vec<SpecificTimeTrigger>>,
    logon: Option<LogonTrigger>,
    boot: Option<BootTrigger>,
    monthly: Option<Vec<MonthlyTrigger>>,
    weekly: Option<Vec<WeeklyTrigger>>,
}

impl TaskTriggers {
    pub fn new(builder: TaskTriggersBuilder) -> Self {
        if builder.number_of_triggers == 0 {
            panic!("There needs to be at least one trigger");
        }
        Self {
            daily: builder.daily,
            event: builder.event,
            idle: builder.idle,
            registration: builder.registration,
            time: builder.specific_times,
            logon: builder.logon,
            boot: builder.boot,
            monthly: builder.monthly,
            weekly: builder.weekly,
        }
    }
}
/// A task can only have up to 48 triggers
/// If you go over that limit then this will panic!
#[derive(Debug)]
pub struct TaskTriggersBuilder {
    number_of_triggers: u8,
    daily: Option<Vec<DailyTrigger>>,
    event: Option<Vec<EventTrigger>>,
    idle: Option<IdleTrigger>,
    registration: Option<RegistrationTrigger>,
    specific_times: Option<Vec<SpecificTimeTrigger>>,
    logon: Option<LogonTrigger>,
    boot: Option<BootTrigger>,
    monthly: Option<Vec<MonthlyTrigger>>,
    weekly: Option<Vec<WeeklyTrigger>>,
}

const MAX_TRIGGERS: u8 = 48;

impl TaskTriggersBuilder {
    pub fn new() -> Self {
        Self {
            number_of_triggers: 0,
            daily: None,
            event: None,
            idle: None,
            registration: None,
            specific_times: None,
            logon: None,
            boot: None,
            monthly: None,
            weekly: None,
        }
    }

    pub fn with_daily(mut self, daily: DailyTrigger) -> Self {
        if self.number_of_triggers >= MAX_TRIGGERS {
            panic!("You can only have up to 48 triggers on a task");
        }

        match self.daily {
            Some(ref mut daily_triggers) => daily_triggers.push(daily),
            None => {
                let mut daily_triggers = Vec::new();
                daily_triggers.push(daily);
                self.daily = Some(daily_triggers);
            }
        }
        self.number_of_triggers += 1;
        self
    }

    pub fn with_specific_time(mut self, specific_time: SpecificTimeTrigger) -> Self {
        if self.number_of_triggers >= MAX_TRIGGERS {
            panic!("You can only have up to 48 triggers on a task");
        }

        match self.specific_times {
            Some(ref mut specific_time_triggers) => specific_time_triggers.push(specific_time),
            None => {
                let mut specific_time_triggers = Vec::new();
                specific_time_triggers.push(specific_time);
                self.specific_times = Some(specific_time_triggers);
            }
        }
        self.number_of_triggers += 1;
        self
    }

    pub fn build(self) -> TaskTriggers {
        TaskTriggers::new(self)
    }
}

#[derive(Clone, Debug)]
pub struct DailyTrigger {
    // validate that end time is after start time
    start_time: Option<DateTime<Utc>>,
    end_time: Option<DateTime<Utc>>,
    interval: Option<u16>,
    id: String,
}
impl DailyTrigger {
    /// The start time will be Utc::now() and the end time will be 100 years
    /// afterwards. The default interval will be every day if not specified
    pub fn new(id: String) -> Self {
        Self {
            id,
            start_time: None,
            end_time: None,
            interval: None,
        }
    }

    /// The start time is also the time the task will be executed daily
    /// If this isn't set then the start time will be set with Utc::now()
    pub fn with_start_time(mut self, start: DateTime<Utc>) -> Self {
        self.start_time = Some(start);
        self
    }
    /// End time specifies the date that the task will stop activating.
    /// If this isn't set then the end date will be ~100 years after Utc::now()
    pub fn with_end_time(mut self, end: DateTime<Utc>) -> Self {
        self.end_time = Some(end);
        self
    }

    /// Interval is unfortunately represented as i16 in windows
    /// so the max interval can only be up to i16::MAX
    /// This probably shouldn't matter in practice
    pub fn with_interval(mut self, interval: u16) -> Self {
        self.interval = Some(interval);
        self
    }
}
#[derive(Debug)]
pub struct EventTrigger {}
#[derive(Debug)]
pub struct IdleTrigger {}
#[derive(Debug)]
pub struct RegistrationTrigger {}
#[derive(Debug)]
pub struct SpecificTimeTrigger {
    id: String,
    time: DateTime<Utc>,
    deactivate_date: Option<DateTime<Utc>>,
}
impl SpecificTimeTrigger {
    pub fn new(id: String, time: DateTime<Utc>) -> Self {
        Self {
            id,
            time,
            deactivate_date: None,
        }
    }

    pub fn deactivate_date(mut self, time: DateTime<Utc>) -> Self {
        self.deactivate_date = Some(time);
        self
    }
}
#[derive(Debug)]
pub struct LogonTrigger {}
#[derive(Debug)]
pub struct BootTrigger {}
#[derive(Debug)]
pub struct MonthlyTrigger {}
#[derive(Debug)]
pub struct WeeklyTrigger {}

// #[cfg(test)]
// mod tests {
//     #[test]
//     fn it_works() {
//         assert_eq!(2 + 2, 4);
//     }
// }
