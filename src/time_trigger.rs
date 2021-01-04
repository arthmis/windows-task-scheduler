use winapi::um::taskschd::ITimeTrigger;

use crate::trigger::TriggerInterface;

pub(crate) struct TimeTrigger<'a> {
    pub(super) time_trigger: &'a mut ITimeTrigger,
}

impl<'a> TimeTrigger<'a> {}
