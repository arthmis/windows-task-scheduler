use std::{ffi::c_void, ptr};

use chrono::{Duration, Utc};
use log::error;
use winapi::{
    shared::winerror::FAILED,
    um::taskschd::{
        IDailyTrigger, ITimeTrigger, ITrigger, ITriggerCollection, TASK_TRIGGER_BOOT,
        TASK_TRIGGER_CUSTOM_TRIGGER_01, TASK_TRIGGER_DAILY, TASK_TRIGGER_EVENT, TASK_TRIGGER_IDLE,
        TASK_TRIGGER_LOGON, TASK_TRIGGER_MONTHLY, TASK_TRIGGER_MONTHLYDOW,
        TASK_TRIGGER_REGISTRATION, TASK_TRIGGER_SESSION_STATE_CHANGE, TASK_TRIGGER_TIME,
        TASK_TRIGGER_WEEKLY,
    },
    Interface,
};

use crate::{
    error::WinError, task::TaskDefinition, to_win_str, trigger::TriggerInterface, DailyTrigger,
    TimeTrigger,
};

/// Provides the methods that are used to add to, remove from, and get the triggers of a task.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-itriggercollection
pub(crate) struct TriggerCollection<'a> {
    trigger_collection: &'a mut ITriggerCollection,
}

impl<'a> TriggerCollection<'a> {
    /// Gets trigger collection using provided TaskDefinition
    pub(crate) fn new(task: &TaskDefinition) -> Result<Self, WinError> {
        unsafe {
            let mut trigger_collection: *mut ITriggerCollection = ptr::null_mut();
            let hr = task
                .task
                .get_Triggers(&mut trigger_collection as *mut *mut ITriggerCollection);
            if FAILED(hr) {
                error!("Cannot get trigger collection: {:X}", hr);
                return Err(WinError::UnknownError(format!(
                    "cannot get trigger collection: {:X}",
                    hr
                )));
            }
            Ok(Self {
                trigger_collection: &mut *trigger_collection,
            })
        }
    }
}

/// Wrappers around ITriggerCollection interface methods
impl<'a> TriggerCollection<'a> {
    /// Creates a new trigger for the task.
    /// This will create the trigger depending on the type and then
    /// execute all of its parameters
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itriggercollection-create
    pub(crate) fn create(&self, task_trigger_type: TaskTriggerType) -> Result<(), WinError> {
        unsafe {
            match task_trigger_type {
                TaskTriggerType::Daily(ref trigger) => {
                    let mut trigger_interface: *mut ITrigger = ptr::null_mut();
                    let hr = self.trigger_collection.Create(
                        // task_trigger_type as u32,
                        task_trigger_type.as_type(),
                        &mut trigger_interface as *mut *mut ITrigger,
                    );
                    if FAILED(hr) {
                        println!("Cannot create trigger: {:X}", hr);
                        return Err(WinError::UnknownError(format!(
                            "Cannot create trigger: {:X}",
                            hr
                        )));
                    }
                    let trigger_interface = &mut *trigger_interface;

                    let mut daily_trigger: *mut IDailyTrigger = ptr::null_mut();
                    let mut hr = trigger_interface.QueryInterface(
                        Box::into_raw(Box::new(IDailyTrigger::uuidof())),
                        &mut daily_trigger as *mut *mut IDailyTrigger as *mut *mut c_void,
                    );
                    // match on each part of the daily struct, this will contain
                    // all the configuration patterns, using this I can execute the
                    // relevant functions on IDailyTrigger
                    if FAILED(hr) {
                        println!("QueryInterface call failed for ITimeTrigger: {:X}", hr);
                        // root_task_folder.Release();
                        // task.Release();
                        // return;
                    }
                    let daily_trigger = daily_trigger.as_mut().unwrap();

                    // let mut hr = daily_trigger.put_Id(to_win_str("Trigger1").as_mut_ptr());
                    let mut hr = daily_trigger.put_Id(to_win_str(&trigger.id).as_mut_ptr());
                    if FAILED(hr) {
                        println!("Cannot put trigger ID: {:X}", hr);
                    }

                    //  Set the task to start at a certain time. The time
                    //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
                    //  For example, the start boundary below
                    //  is January 1st 2005 at 12:05
                    // hr = time_trigger.put_StartBoundary(to_win_str("2021-01-01T12:11:00").as_mut_ptr());
                    let start_time = if let Some(start_time) = trigger.start_time {
                        start_time
                    } else {
                        Utc::now()
                    };
                    hr = daily_trigger
                        .put_StartBoundary(to_win_str(&start_time.to_rfc3339()).as_mut_ptr());
                    if FAILED(hr) {
                        println!("Cannot add start boundary to trigger: {:x}", hr);
                        // root_task_folder.Release();
                        // task.Release();
                    }
                    // hr = time_trigger.put_EndBoundary(to_win_str("2021-01-01T12:12:00").as_mut_ptr());
                    let end_time = if let Some(end_time) = trigger.end_time {
                        end_time
                    } else {
                        Utc::now() + Duration::weeks(52 * 100)
                    };
                    hr = daily_trigger
                        .put_EndBoundary(to_win_str(&end_time.to_rfc3339()).as_mut_ptr());
                    if FAILED(hr) {
                        println!("Cannot put end boundary on trigger: {:X}", hr);
                    }

                    let interval = if let Some(interval) = trigger.interval {
                        interval
                    } else {
                        1
                    };
                    daily_trigger.put_DaysInterval(interval as i16);
                    daily_trigger.Release();
                }
                TaskTriggerType::SpecificTime(ref specific_time_trigger) => {
                    let mut trigger_interface: *mut ITrigger = ptr::null_mut();
                    let hr = self.trigger_collection.Create(
                        task_trigger_type.as_type(),
                        &mut trigger_interface as *mut *mut ITrigger,
                    );
                    if FAILED(hr) {
                        println!("Cannot create trigger: {:X}", hr);
                        return Err(WinError::UnknownError(format!(
                            "Cannot create trigger: {:X}",
                            hr
                        )));
                    }
                    let trigger_interface = &mut *trigger_interface;

                    let mut time_trigger: *mut ITimeTrigger = ptr::null_mut();
                    let mut hr = trigger_interface.QueryInterface(
                        Box::into_raw(Box::new(ITimeTrigger::uuidof())),
                        &mut time_trigger as *mut *mut ITimeTrigger as *mut *mut c_void,
                    );
                    // match on each part of the daily struct, this will contain
                    // all the configuration patterns, using this I can execute the
                    // relevant functions on IDailyTrigger
                    if FAILED(hr) {
                        println!("QueryInterface call failed for ITimeTrigger: {:X}", hr);
                        // root_task_folder.Release();
                        // task.Release();
                        // return;
                    }
                    let time_trigger = time_trigger.as_mut().unwrap();

                    // let mut hr = daily_trigger.put_Id(to_win_str("Trigger1").as_mut_ptr());
                    hr = time_trigger.put_Id(to_win_str(&specific_time_trigger.id).as_mut_ptr());
                    if FAILED(hr) {
                        println!("Cannot put trigger ID: {:X}", hr);
                    }

                    //  Set the task to start at a certain time. The time
                    //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
                    //  For example, the start boundary below
                    //  is January 1st 2005 at 12:05
                    // hr = time_trigger.put_StartBoundary(to_win_str("2021-01-01T12:11:00").as_mut_ptr());
                    let start_time = specific_time_trigger.time;
                    hr = time_trigger
                        .put_StartBoundary(to_win_str(&start_time.to_rfc3339()).as_mut_ptr());
                    if FAILED(hr) {
                        println!("Cannot add start boundary to trigger: {:x}", hr);
                        // root_task_folder.Release();
                        // task.Release();
                    }
                    // hr = time_trigger.put_EndBoundary(to_win_str("2021-01-01T12:12:00").as_mut_ptr());
                    let end_time = if let Some(end_time) = specific_time_trigger.deactivate_date {
                        end_time
                    } else {
                        Utc::now() + Duration::weeks(52 * 100)
                    };
                    hr = time_trigger
                        .put_EndBoundary(to_win_str(&end_time.to_rfc3339()).as_mut_ptr());
                    if FAILED(hr) {
                        println!("Cannot put end boundary on trigger: {:X}", hr);
                    }

                    time_trigger.Release();
                }
                _ => unimplemented!("These will be implemented some day"),
            }
        }

        Ok(())
    }
}

impl<'a> Drop for TriggerCollection<'a> {
    fn drop(&mut self) {
        unsafe {
            self.trigger_collection.Release();
        }
    }
}

// impl<'a> TriggerInterface<'a> {
//     /// Creates new trigger interface from trigger collection
//     ///
//     /// Unsafety: The trigger parameter should point to valid ITrigger data
//     unsafe fn new(trigger: *mut ITrigger) -> Self {
//         TriggerInterface {
//             trigger: &mut *trigger,
//         }
//     }
// }

/// When the task will be triggered
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/ne-taskschd-task_trigger_type2
pub enum TaskTriggerType {
    /// Triggers the task when a specific event occurs.
    Event,
    /// Triggers the task at a specific time of day.
    SpecificTime(TimeTrigger),
    /// Triggers the task on a daily schedule. For example, the task starts at a specific time every day, every-other day, every third day, and so on.
    Daily(DailyTrigger),
    /// Triggers the task on a weekly schedule. For example, the task starts at 8:00 AM on a specific day every week or other week.
    Weekly,
    /// Triggers the task on a monthly schedule. For example, the task starts on specific days of specific months.
    Monthly,
    /// Triggers the task on a monthly day-of-week schedule. For example, the task starts on a specific days of the week, weeks of the month, and months of the year.
    MonthlyDow,
    /// Triggers the task when the computer goes into an idle state.
    Idle,
    /// Triggers the task when the task is registered.
    Registration,
    /// Triggers the task when the computer boots.
    Boot,
    /// Triggers the task when a specific user logs on.
    Logon,
    /// Triggers the task when a specific session state changes.
    SessionStateChange,
    // I will have to investigate it a little
    /// This doesn't have docs currently
    CustomTrigger01,
}

impl TaskTriggerType {
    fn as_type(&self) -> u32 {
        match self {
            TaskTriggerType::Event => TASK_TRIGGER_EVENT,
            TaskTriggerType::SpecificTime(_) => TASK_TRIGGER_TIME,
            TaskTriggerType::Daily(_) => TASK_TRIGGER_DAILY,
            TaskTriggerType::Weekly => TASK_TRIGGER_WEEKLY,
            TaskTriggerType::Monthly => TASK_TRIGGER_MONTHLY,
            TaskTriggerType::MonthlyDow => TASK_TRIGGER_MONTHLYDOW,
            TaskTriggerType::Idle => TASK_TRIGGER_IDLE,
            TaskTriggerType::Registration => TASK_TRIGGER_REGISTRATION,
            TaskTriggerType::Boot => TASK_TRIGGER_BOOT,
            TaskTriggerType::Logon => TASK_TRIGGER_LOGON,
            TaskTriggerType::SessionStateChange => TASK_TRIGGER_SESSION_STATE_CHANGE,
            TaskTriggerType::CustomTrigger01 => TASK_TRIGGER_CUSTOM_TRIGGER_01,
        }
    }
}
