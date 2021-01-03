use std::ptr;

use log::error;
use winapi::{
    shared::winerror::FAILED,
    um::taskschd::{ITrigger, ITriggerCollection, TASK_TRIGGER_TIME},
};

use crate::{error::WinError, task::Task, trigger::Trigger};

/// Provides the methods that are used to add to, remove from, and get the triggers of a task.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-itriggercollection
pub(crate) struct TriggerCollection<'a> {
    trigger_collection: &'a mut ITriggerCollection,
}

impl<'a> TriggerCollection<'a> {
    /// Gets trigger collection using provided TaskDefinition
    pub(crate) fn new(task: &Task) -> Result<Self, WinError> {
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
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itriggercollection-create
    pub(crate) fn create(&self, task_trigger_type: TaskTriggerType) -> Result<Trigger, WinError> {
        unsafe {
            let mut trigger: *mut ITrigger = ptr::null_mut();
            let hr = self
                .trigger_collection
                .Create(task_trigger_type as u32, &mut trigger as *mut *mut ITrigger);
            if FAILED(hr) {
                println!("Cannot create trigger: {:X}", hr);
                return Err(WinError::UnknownError(format!(
                    "Cannot create trigger: {:X}",
                    hr
                )));
            }
            Ok(Trigger {
                trigger: &mut *trigger,
            })
        }
    }
}

impl<'a> Drop for TriggerCollection<'a> {
    fn drop(&mut self) {
        unsafe {
            self.trigger_collection.Release();
        }
    }
}

/// When the task will be triggered
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/ne-taskschd-task_trigger_type2
#[repr(u32)]
pub enum TaskTriggerType {
    /// Triggers the task when a specific event occurs.
    TaskTriggerEvent,
    /// Triggers the task at a specific time of day.
    TaskTriggerTime,
    /// Triggers the task on a daily schedule. For example, the task starts at a specific time every day, every-other day, every third day, and so on.
    TaskTriggerDaily,
    /// Triggers the task on a weekly schedule. For example, the task starts at 8:00 AM on a specific day every week or other week.
    TaskTriggerWeekly,
    /// Triggers the task on a monthly schedule. For example, the task starts on specific days of specific months.
    TaskTriggerMonthly,
    /// Triggers the task on a monthly day-of-week schedule. For example, the task starts on a specific days of the week, weeks of the month, and months of the year.
    TaskTriggerMonthlydow,
    /// Triggers the task when the computer goes into an idle state.
    TaskTriggerIdle,
    /// Triggers the task when the task is registered.
    TaskTriggerRegistration,
    /// Triggers the task when the computer boots.
    TaskTriggerBoot,
    /// Triggers the task when a specific user logs on.
    TaskTriggerLogon,
    /// Triggers the task when a specific session state changes.
    TaskTriggerSessionStateChange,
    // I will have to investigate it a little
    /// This doesn't have docs currently
    TaskTriggerCustomTrigger01,
}
