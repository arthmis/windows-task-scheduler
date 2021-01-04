use std::unreachable;

use log::error;
use winapi::{shared::winerror::FAILED, um::taskschd::IRegistrationInfo};

use crate::{task::TaskDefinition, to_win_str};

// This could be a trait, since it's supposed to be an interface and
// not actually a type
/// Provides the administrative information that can be used to describe the task. This information includes details such as a description of the task, the author of the task, the date the task is registered, and the security descriptor of the task.
///
/// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nn-taskschd-iregistrationinfo
pub(crate) struct RegistrationInfo<'a> {
    registration_info: &'a mut IRegistrationInfo,
}

impl<'a> RegistrationInfo<'a> {
    /// Gets or sets the registration information used to describe a task, such as the description of the task, the author of the task, and the date the task is registered.
    /// This property is read/write
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-itaskdefinition-get_registrationinfo
    pub(crate) fn new(task: &TaskDefinition) -> Self {
        unsafe {
            // Get the registration info for setting the identification
            let mut registration_info: *mut IRegistrationInfo = core::ptr::null_mut();
            let mut hr = task
                .task
                .get_RegistrationInfo(&mut registration_info as *mut *mut IRegistrationInfo);

            if FAILED(hr) {
                error!("Cannot get identification pointer: {:X}", hr);
                // this should be unreachable because the msdn doc claims
                // that function doesn't return anything, although it returns
                // HRESULT with winapi
                unreachable!();
            }
            Self {
                registration_info: &mut *registration_info,
            }
        }
    }
    // I will probably think about making this not read or write
    // since there is a get_author and put_author
    /// Gets or sets the author of the task.
    /// This property is read/write
    ///
    /// https://docs.microsoft.com/en-us/windows/win32/api/taskschd/nf-taskschd-iregistrationinfo-put_author
    pub(crate) fn put_author(&self, author: &str) {
        let author_win_str = to_win_str(author);

        unsafe {
            let hr = self
                .registration_info
                .put_Author(to_win_str("author name").as_mut_ptr());
            // this should not fail, the docs says this function doesn't return anything
            if FAILED(hr) {
                error!("Cannot put identification info: {:X}", hr);
                unreachable!();
            }
        }
    }
}

impl<'a> Drop for RegistrationInfo<'a> {
    fn drop(&mut self) {
        unsafe {
            self.registration_info.Release();
        }
    }
}
