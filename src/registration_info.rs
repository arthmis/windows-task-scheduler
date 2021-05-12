use std::unreachable;

use bindings::Windows::Win32::{Automation::BSTR, TaskScheduler::IRegistrationInfo};
use log::error;

// This could be a trait, since it's supposed to be an interface and
// not actually a type
/// Provides the administrative information that can be used to describe the task. This information includes details such as a description of the task, the author of the task, the date the task is registered, and the security descriptor of the task.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-iregistrationinfo
pub(crate) struct RegistrationInfo(pub(crate) IRegistrationInfo);

impl RegistrationInfo {
    /// Gets or sets the registration information used to describe a task, such as the description of the task, the author of the task, and the date the task is registered.
    /// This property is read/write
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_registrationinfo
    pub(crate) fn new(registration_info: IRegistrationInfo) -> Self {
        // Get the registration info for setting the identification
        Self(registration_info)
    }

    // I will probably think about making this not read or write
    // since there is a get_author and put_author
    /// Gets or sets the author of the task.
    /// This property is read/write
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-iregistrationinfo-put_author
    pub(crate) fn put_author(&self, author: &str) -> Result<(), windows::Error> {
        let author = BSTR::from(author);
        unsafe { self.0.put_Author(author).ok() }
    }
}
