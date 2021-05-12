fn main() {
    windows::build!(
        // Windows::Win32::TaskScheduler::{ITaskService, TaskScheduler, ITaskFolder, ITaskDefinition, IRegistrationInfo, IPrincipal, ITaskSettings, IIdleSettings, ITriggerCollection, ITrigger, ITimeTrigger, IActionCollection, IAction, IExecAction},
        Windows::Win32::TaskScheduler::*,
        Windows::Win32::Automation::VARIANT,
        Windows::Win32::Com::{
            CoCreateInstance, CoInitializeEx, COINIT, CoUninitialize,
        },
    );
}
