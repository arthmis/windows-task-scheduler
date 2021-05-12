use std::{ptr, unreachable};

use bindings::Windows::Win32::TaskScheduler::IPrincipal;
use log::error;

/// Provides the security credentials for a principal. These security credentials define the security context for the tasks that are associated with the principal.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_principal
pub(crate) struct Principal(pub(crate) IPrincipal);

impl Principal {
    /// Gets the principal for the task that provides the security credentials for the task.
    pub(crate) fn new(principal: IPrincipal) -> Self {
        Self(principal)
    }

    /// Gets or sets the security logon method that is required to run the tasks that are associated with the principal.
    /// This property is read/write.
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-iprincipal-put_logontype
    pub(crate) fn put_logon_type(&self, task_logon_kind: TaskLogon) -> Result<(), windows::Error> {
        unsafe {
            // set up principal logon type to interactive logon
            // let hr = self.principal.put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
            self.0.put_LogonType((task_logon_kind as i32).into()).ok()
        }
    }
}

/// Defines what logon technique is required to run a task.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/ne-taskschd-task_logon_type
#[derive(Debug)]
#[repr(i32)]
pub(crate) enum TaskLogon {
    /// The logon method is not specified. Used for non-NT credentials.
    None = 0,
    /// Use a password for logging on the user. The password must be supplied at registration time.
    Password = 1,
    /// The service will log the user on using Service For User (S4U), and the task will run in a non-interactive desktop. When an S4U logon is used, no password is stored by the system and there is no access to either the network or to encrypted files.
    S4U = 2,
    /// User must already be logged on. The task will be run only in an existing interactive session.
    InteractiveToken = 3,
    /// Group activation. The groupId field specifies the group.
    Group = 4,
    /// Indicates that a Local System, Local Service, or Network Service account is being used as a security context to run the task.
    ServiceAccount = 5,
    /// Not in use; currently identical to Password
    InteractiveTokenOrPassword = 6,
}
