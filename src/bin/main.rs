#![allow(warnings)]

use chrono::Duration;
use combaseapi::{CoCreateInstance, CoInitializeSecurity};
use core::default::Default;
use iter::once;
use std::{ffi::OsStr, path::PathBuf, ptr};
use std::{iter, os::windows::ffi::OsStrExt};
use task_scheduler::{execute_task, Actions, DailyTrigger, TaskTriggersBuilder, TimeTrigger, Utc};
use winapi::{
    shared::wtypes::VARIANT_TRUE,
    um::{
        oaidl::VARIANT,
        taskschd::{
            IAction, IActionCollection, IExecAction, IIdleSettings, IPrincipal, IRegisteredTask,
            IRegistrationInfo, ITaskDefinition, ITaskFolder, ITaskSettings, ITimeTrigger, ITrigger,
            ITriggerCollection, TASK_ACTION_EXEC, TASK_CREATE_OR_UPDATE,
            TASK_LOGON_INTERACTIVE_TOKEN, TASK_TRIGGER_TIME,
        },
    },
};
// use task_scheduler::Task;
use winapi::{
    ctypes::c_void,
    um::{combaseapi, objbase},
};
use winapi::{shared::rpcdce::RPC_C_AUTHN_LEVEL_PKT_PRIVACY, um::taskschd::ITaskService};
use winapi::{shared::winerror::FAILED, um::taskschd::TaskScheduler};
use winapi::{
    shared::{
        guiddef::GUID, rpcdce::RPC_C_IMP_LEVEL_IMPERSONATE, wtypesbase::CLSCTX_INPROC_SERVER,
    },
    Interface,
};
use winapi::{um::oaidl::VARIANT_n1, Class};

fn to_win_str(string: &str) -> Vec<u16> {
    OsStr::new(string).encode_wide().chain(once(0)).collect()
}
fn main() {
    let task_path = PathBuf::from("C:\\Windows\\System32\\notepad.exe");
    let task_name = "Open Notepad";
    let now = Utc::now();

    let daily_trigger = DailyTrigger::new("my daily trigger".to_string())
        .with_start_time(Utc::now() + Duration::seconds(3))
        .with_end_time(Utc::now() + Duration::minutes(1));

    let time_trigger = TimeTrigger::new(
        "My Time Trigger".to_string(),
        Utc::now() + Duration::seconds(1),
    );
    let other_time_trigger = TimeTrigger::new(
        "My Other Time Trigger".to_string(),
        Utc::now() + Duration::seconds(10),
    );

    let triggers = TaskTriggersBuilder::new();
    // let triggers = triggers.with_daily(daily_trigger).build();
    let triggers = triggers
        .with_specific_time(time_trigger)
        .with_specific_time(other_time_trigger)
        .build();
    let actions = Actions::new(task_path);

    execute_task(actions, task_name.to_owned(), triggers);
}
