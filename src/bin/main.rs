#![allow(warnings)]

use combaseapi::{CoCreateInstance, CoInitializeSecurity, CoUninitialize};
use core::default::Default;
use iter::once;
use std::{ffi::OsStr, ptr};
use std::{iter, os::windows::ffi::OsStrExt};
use winapi::{
    shared::wtypes::VARIANT_TRUE,
    um::{
        oaidl::VARIANT,
        taskschd::{
            IIdleSettings, IPrincipal, IRegistrationInfo, ITaskDefinition, ITaskFolder,
            ITaskSettings, ITimeTrigger, ITrigger, ITriggerCollection,
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
    unsafe {
        // initialize COM
        let mut hr = combaseapi::CoInitializeEx(ptr::null_mut(), objbase::COINIT_APARTMENTTHREADED);
        if FAILED(hr) {
            println!("Com initialization failed: {:X}", hr);
            return;
        }

        // set general COM security levels
        hr = CoInitializeSecurity(
            ptr::null_mut(),
            -1,
            ptr::null_mut(),
            ptr::null_mut(),
            RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            ptr::null_mut(),
            0,
            ptr::null_mut(),
        );

        if FAILED(hr) {
            println!("Com security initialization failed: {:X}", hr);
            combaseapi::CoUninitialize();
            return;
        }

        // name for task
        let mut task_name = to_win_str("Trigger Notepad");
        // path for notepad program
        let exe_path = to_win_str("C:\\Windows\\System32\\notepad.exe");

        // Create an instance of the task service
        // this isn't properly documented, however these are pointers to GUIDs for these particular
        // classes. winapi has uuidof method to get the guid
        // I figured it out because of this https://docs.microsoft.com/en-us/archive/msdn-magazine/2007/october/windows-with-c-task-scheduler-2-0
        // here the got the uuid of the task scheduler using a function and the object
        let CLSID_TaskScheduler = Box::into_raw(Box::new(TaskScheduler::uuidof()));
        let IID_ITaskService = Box::into_raw(Box::new(ITaskService::uuidof()));
        let mut task_service: *mut ITaskService = core::ptr::null_mut();
        hr = CoCreateInstance(
            CLSID_TaskScheduler,
            ptr::null_mut(),
            CLSCTX_INPROC_SERVER,
            IID_ITaskService,
            // have to figure out how this works
            &mut task_service as *mut *mut ITaskService as *mut *mut c_void,
        );

        if FAILED(hr) {
            println!("Failed to create an instance of ITaskService: {:X}", hr);
            CoUninitialize();
            return;
        }
        let task_service = task_service.as_mut().unwrap();

        // connect to the task service
        let variant: VARIANT = Default::default();
        hr = task_service.Connect(variant, variant, variant, variant);

        if FAILED(hr) {
            println!("ITaskService::Connect failed: {:X}", hr);
            task_service.Release();
            CoUninitialize();
            return;
        }

        // get the pointer to the root task folder. This folder will hold the
        // new task that is registered
        let mut root_task_folder: *mut ITaskFolder = core::ptr::null_mut();
        hr = task_service.GetFolder(
            to_win_str("\\").as_mut_ptr(),
            &mut root_task_folder as *mut *mut ITaskFolder,
        );
        if FAILED(hr) {
            println!("Cannot get root folder pointer: {:X}", hr);
            task_service.Release();
            CoUninitialize();
            return;
        }
        let root_task_folder = root_task_folder.as_mut().unwrap();

        // if the same task exists, remove it
        root_task_folder.DeleteTask(task_name.as_mut_ptr(), 0);

        // Create the task definition object to create the task
        let mut task: *mut ITaskDefinition = core::ptr::null_mut();
        hr = task_service.NewTask(0, &mut task as *mut *mut ITaskDefinition);
        if FAILED(hr) {
            println!(
                "Failed to CoCreate an instance of the TaskService class: {:X}",
                hr
            );
            root_task_folder.Release();
            CoUninitialize();
            return;
        }
        let task = task.as_mut().unwrap();
        // COM clean up. Pointer is no longer used
        task_service.Release();

        // Get the registration info for setting the identification
        let mut registration_info: *mut IRegistrationInfo = core::ptr::null_mut();
        hr = task.get_RegistrationInfo(&mut registration_info as *mut *mut IRegistrationInfo);

        if FAILED(hr) {
            println!("Cannot get identification pointer: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }
        let mut registration_info = registration_info.as_mut().unwrap();
        hr = registration_info.put_Author(to_win_str("author name").as_mut_ptr());
        if FAILED(hr) {
            println!("Cannot put identification info: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }

        // Create the principal for the task - these credentials are
        // overwritten with the credentials passed to RegisterTaskDefinitio1n
        let mut principal: *mut IPrincipal = ptr::null_mut();
        hr = task.get_Principal(&mut principal as *mut *mut IPrincipal);
        if FAILED(hr) {
            println!("Cannot get principal pointer: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }
        let mut principal = principal.as_mut().unwrap();

        // set up principal logon type to interactive logon
        hr = principal.put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
        if FAILED(hr) {
            println!("Cannot put principal info: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }

        // create the settings for the task
        let mut settings: *mut ITaskSettings = ptr::null_mut();
        hr = task.get_Settings(&mut settings as *mut *mut ITaskSettings);
        if FAILED(hr) {
            println!("Cannot get settings pointer: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }

        let settings = settings.as_mut().unwrap();

        // set settings values for the task
        hr = settings.put_StartWhenAvailable(VARIANT_TRUE);
        settings.Release();
        if FAILED(hr) {
            println!("Cannot put setting information: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }

        // set the idle settings for the task
        let mut idle_settings: *mut IIdleSettings = ptr::null_mut();
        hr = settings.get_IdleSettings(&mut idle_settings as *mut *mut IIdleSettings);
        if FAILED(hr) {
            println!("Cannot get idle setting information: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }
        let idle_settings = idle_settings.as_mut().unwrap();

        idle_settings.put_WaitTimeout(to_win_str("PT5M").as_mut_ptr());
        idle_settings.Release();
        if FAILED(hr) {
            println!("Cannot put idle setting information: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }

        // get the trigger collection to insert the time trigger
        let mut trigger_collection: *mut ITriggerCollection = ptr::null_mut();
        hr = task.get_Triggers(&mut trigger_collection as *mut *mut ITriggerCollection);
        if FAILED(hr) {
            println!("Cannot get trigger collection: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }
        let trigger_collection = trigger_collection.as_mut().unwrap();

        // add the time trigger to the task
        let mut trigger: *mut ITrigger = ptr::null_mut();
        hr = trigger_collection.Create(TASK_TRIGGER_TIME, &mut trigger as *mut *mut ITrigger);
        if FAILED(hr) {
            println!("Cannot create trigger: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }
        let trigger = trigger.as_mut().unwrap();

        let mut time_trigger: *mut ITimeTrigger = ptr::null_mut();
        hr = trigger.QueryInterface(
            Box::into_raw(Box::new(ITimeTrigger::uuidof())),
            &mut time_trigger as *mut *mut ITimeTrigger as *mut *mut c_void,
        );
        trigger.Release();
        if FAILED(hr) {
            println!("QueryInterface call failed for ITimeTrigger: {:X}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }
        let time_trigger = time_trigger.as_mut().unwrap();

        hr = time_trigger.put_Id(to_win_str("Trigger1").as_mut_ptr());
        if FAILED(hr) {
            println!("Cannot put trigger ID: {:X}", hr);
        }

        hr = time_trigger.put_EndBoundary(to_win_str("2015-05-02T08:00:00").as_mut_ptr());
        if FAILED(hr) {
            println!("Cannot put end boundary on trigger: {:X}", hr);
        }

        //  Set the task to start at a certain time. The time
        //  format should be YYYY-MM-DDTHH:MM:SS(+-)(timezone).
        //  For example, the start boundary below
        //  is January 1st 2005 at 12:05
        hr = time_trigger.put_StartBoundary(to_win_str("2005-01-01T12:05:00").as_mut_ptr());
        time_trigger.Release();
        if FAILED(hr) {
            println!("Cannot add start boundary to trigger: {:x}", hr);
            root_task_folder.Release();
            task.Release();
            CoUninitialize();
            return;
        }
    }
}
