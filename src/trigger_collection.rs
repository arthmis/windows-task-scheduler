use std::{ffi::c_void, ptr};

use bindings::Windows::Win32::TaskScheduler::{
    ITrigger, ITriggerCollection, TASK_TRIGGER, TASK_TRIGGER_TYPE2,
};
use chrono::{Duration, Utc};
use log::error;

use crate::{DailyTrigger, SpecificTimeTrigger};

/// Provides the methods that are used to add to, remove from, and get the triggers of a task.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-itriggercollection
pub(crate) struct TriggerCollection(pub(crate) ITriggerCollection);

impl TriggerCollection {
    /// Gets trigger collection using provided TaskDefinition
    pub(crate) fn new(collection: ITriggerCollection) -> Self {
        Self(collection)
    }
}

/// Wrappers around ITriggerCollection interface methods
impl TriggerCollection {
    /// Creates a new trigger for the task.
    /// This will create the trigger depending on the type and then
    /// execute all of its parameters
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itriggercollection-create
    pub(crate) fn create(
        &self,
        task_trigger_type: TASK_TRIGGER_TYPE2,
    ) -> Result<ITrigger, windows::Error> {
        // match task_trigger_type {
        // TaskTriggerType::Daily(ref trigger) => {
        //     let mut trigger_interface: *mut ITrigger = ptr::null_mut();
        //     let hr = unsafe {
        //         self.trigger_collection.Create(
        //             // task_trigger_type as u32,
        //             task_trigger_type.as_type(),
        //             &mut trigger_interface as *mut *mut ITrigger,
        //         )
        //     };
        //     if FAILED(hr) {
        //         println!("Cannot create trigger: {:X}", hr);
        //         return Err(WinError::UnknownError(format!(
        //             "Cannot create trigger: {:X}",
        //             hr
        //         )));
        //     }
        //     let trigger_interface = &mut *trigger_interface;

        //     let mut daily_trigger: *mut IDailyTrigger = ptr::null_mut();
        //     let mut hr = trigger_interface.QueryInterface(
        //         Box::into_raw(Box::new(IDailyTrigger::uuidof())),
        //         &mut daily_trigger as *mut *mut IDailyTrigger as *mut *mut c_void,
        //     );
        //     // match on each part of the daily struct, this will contain
        //     // all the configuration patterns, using this I can execute the
        //     // relevant functions on IDailyTrigger
        //     if FAILED(hr) {
        //         println!("QueryInterface call failed for ITimeTrigger: {:X}", hr);
        //         // root_task_folder.Release();
        //         // task.Release();
        //         // return;
        //     }
        //     let daily_trigger = daily_trigger.as_mut().unwrap();

        //     // let mut hr = daily_trigger.put_Id(to_win_str("Trigger1").as_mut_ptr());
        //     let mut hr = daily_trigger.put_Id(to_win_str(&trigger.id).as_mut_ptr());
        //     if FAILED(hr) {
        //         println!("Cannot put trigger ID: {:X}", hr);
        //     }

        //     //  Set the task to start at a certain time. The time
        //     //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
        //     //  For example, the start boundary below
        //     //  is January 1st 2005 at 12:05
        //     // hr = time_trigger.put_StartBoundary(to_win_str("2021-01-01T12:11:00").as_mut_ptr());
        //     let start_time = if let Some(start_time) = trigger.start_time {
        //         start_time
        //     } else {
        //         Utc::now()
        //     };
        //     hr = daily_trigger
        //         .put_StartBoundary(to_win_str(&start_time.to_rfc3339()).as_mut_ptr());
        //     if FAILED(hr) {
        //         println!("Cannot add start boundary to trigger: {:x}", hr);
        //         // root_task_folder.Release();
        //         // task.Release();
        //     }
        //     // hr = time_trigger.put_EndBoundary(to_win_str("2021-01-01T12:12:00").as_mut_ptr());
        //     let end_time = if let Some(end_time) = trigger.end_time {
        //         end_time
        //     } else {
        //         Utc::now() + Duration::weeks(52 * 100)
        //     };
        //     hr = daily_trigger.put_EndBoundary(to_win_str(&end_time.to_rfc3339()).as_mut_ptr());
        //     if FAILED(hr) {
        //         println!("Cannot put end boundary on trigger: {:X}", hr);
        //     }

        //     let interval = if let Some(interval) = trigger.interval {
        //         interval
        //     } else {
        //         1
        //     };
        //     daily_trigger.put_DaysInterval(interval as i16);
        //     daily_trigger.Release();
        // }
        let mut trigger = None;
        unsafe {
            let res = self.0.Create(task_trigger_type, &mut trigger).ok();
            match res {
                Ok(_) => Ok(trigger.unwrap()),
                Err(err) => Err(err),
            }
        }
        //     TaskTriggerType::SpecificTime(ref specific_time_trigger) => {
        //         let mut trigger_interface: *mut ITrigger = ptr::null_mut();
        //         let hr = unsafe {
        //             self.trigger_collection.Create(
        //                 task_trigger_type.as_type(),
        //                 &mut trigger_interface as *mut *mut ITrigger,
        //             )
        //         };
        //         if FAILED(hr) {
        //             println!("Cannot create trigger: {:X}", hr);
        //             return Err(TaskError::from(WinError::UnknownError(format!(
        //                 "Cannot create trigger with type SpecificTime\nError code is: {:X}",
        //                 hr
        //             ))));
        //         }
        //         let trigger_interface = unsafe { TriggerInterface::new(trigger_interface) };

        //         let mut time_trigger: *mut ITimeTrigger = ptr::null_mut();
        //         let mut hr = unsafe {
        //             trigger_interface.trigger.QueryInterface(
        //                 Box::into_raw(Box::new(ITimeTrigger::uuidof())),
        //                 &mut time_trigger as *mut *mut ITimeTrigger as *mut *mut c_void,
        //             )
        //         };
        //         if FAILED(hr) {
        //             println!("QueryInterface call failed for ITimeTrigger: {:X}", hr);
        //             return Err(TaskError::from(WinError::UnknownError(format!(
        //                 "QueryInterface call failed for ITimeTrigger: {:X}",
        //                 hr,
        //             ))));
        //         }
        //         let time_trigger = unsafe { TimeTrigger::new(time_trigger) };

        //         let hr = unsafe {
        //             time_trigger
        //                 .time_trigger
        //                 .put_Id(to_win_str(&specific_time_trigger.id).as_mut_ptr())
        //         };
        //         if FAILED(hr) {
        //             println!("Cannot put trigger ID: {:X}", hr);
        //             return Err(TaskError::from(WinError::UnknownError(format!(
        //                 "Cannot put trigger ID: {:X}",
        //                 hr,
        //             ))));
        //         }

        //         //  Set the task to start at a certain time. The time
        //         //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
        //         let start_time = specific_time_trigger.time;
        //         let hr = unsafe {
        //             time_trigger
        //                 .time_trigger
        //                 .put_StartBoundary(to_win_str(&start_time.to_rfc3339()).as_mut_ptr())
        //         };
        //         if FAILED(hr) {
        //             println!("Cannot add start boundary to trigger: {:x}", hr);
        //             return Err(TaskError::from(WinError::UnknownError(format!(
        //                 "Cannot add start boundary to trigger: {:X}",
        //                 hr,
        //             ))));
        //         }
        //         let end_time = if let Some(end_time) = specific_time_trigger.deactivate_date {
        //             end_time
        //         } else {
        //             Utc::now() + Duration::weeks(52 * 100)
        //         };
        //         let hr = unsafe {
        //             time_trigger
        //                 .time_trigger
        //                 .put_EndBoundary(to_win_str(&end_time.to_rfc3339()).as_mut_ptr())
        //         };
        //         if FAILED(hr) {
        //             println!("Cannot put end boundary on trigger: {:X}", hr);
        //             return Err(TaskError::from(WinError::UnknownError(format!(
        //                 "Cannot add end boundary to trigger: {:X}",
        //                 hr,
        //             ))));
        //         }
        //     }
        //     _ => unimplemented!("These will be implemented some day"),
        // }

        // Ok(())
    }
}
/// When the task will be triggered
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/ne-taskschd-task_trigger_type2
#[derive(Debug)]
pub enum TaskTriggerType {
    /// Triggers the task when a specific event occurs.
    Event,
    /// Triggers the task at a specific time of day.
    SpecificTime(SpecificTimeTrigger),
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

// impl TaskTriggerType {
//     fn as_type(&self) -> u32 {
//         match self {
//             TaskTriggerType::Event => TASK_TRIGGER_EVENT,
//             TaskTriggerType::SpecificTime(_) => TASK_TRIGGER_TIME,
//             TaskTriggerType::Daily(_) => TASK_TRIGGER_DAILY,
//             TaskTriggerType::Weekly => TASK_TRIGGER_WEEKLY,
//             TaskTriggerType::Monthly => TASK_TRIGGER_MONTHLY,
//             TaskTriggerType::MonthlyDow => TASK_TRIGGER_MONTHLYDOW,
//             TaskTriggerType::Idle => TASK_TRIGGER_IDLE,
//             TaskTriggerType::Registration => TASK_TRIGGER_REGISTRATION,
//             TaskTriggerType::Boot => TASK_TRIGGER_BOOT,
//             TaskTriggerType::Logon => TASK_TRIGGER_LOGON,
//             TaskTriggerType::SessionStateChange => TASK_TRIGGER_SESSION_STATE_CHANGE,
//             TaskTriggerType::CustomTrigger01 => TASK_TRIGGER_CUSTOM_TRIGGER_01,
//         }
//     }
// }

// struct TimeTrigger<'a> {
//     time_trigger: &'a mut ITimeTrigger,
// }

// impl<'a> TimeTrigger<'a> {
//     unsafe fn new(time_trigger: *mut ITimeTrigger) -> Self {
//         Self {
//             time_trigger: &mut *time_trigger,
//         }
//     }
// }
