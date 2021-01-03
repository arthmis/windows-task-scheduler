use std::ptr;

use winapi::{
    shared::winerror::FAILED,
    um::taskschd::{ITrigger, TASK_TRIGGER_TIME},
};

use crate::{
    error::WinError,
    trigger_collection::{TaskTriggerType, TriggerCollection},
};

/// Provides the common properties that are inherited by all trigger objects.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-itrigger
pub(crate) struct Trigger<'a> {
    pub(crate) trigger: &'a mut ITrigger,
}

impl<'a> Trigger<'a> {
    /// Creates new trigger interface from trigger collection
    pub(crate) fn new(
        trigger_collection: &'a TriggerCollection,
        task_trigger_type: TaskTriggerType,
    ) -> Result<Self, WinError> {
        trigger_collection.create(task_trigger_type)
    }
}

impl<'a> Trigger<'a> {
    // pub(crate) fn query_interface(&self) {
    //     let mut time_trigger: *mut ITimeTrigger = ptr::null_mut();
    //     let mut hr = trigger.trigger.QueryInterface(
    //         Box::into_raw(Box::new(ITimeTrigger::uuidof())),
    //         &mut time_trigger as *mut *mut ITimeTrigger as *mut *mut c_void,
    //     );
    //     if FAILED(hr) {
    //         println!("QueryInterface call failed for ITimeTrigger: {:X}", hr);
    //         // root_task_folder.Release();
    //         // task.Release();
    //         return;
    //     }
    //     let time_trigger = time_trigger.as_mut().unwrap();
    // }
}

impl<'a> Drop for Trigger<'a> {
    fn drop(&mut self) {
        unsafe {
            self.trigger.Release();
        }
    }
}
