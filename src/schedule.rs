use windows::Win32::Foundation::BSTR;
use windows::Win32::System::TaskScheduler::{
    IActionCollection, IRegistrationInfo, ITaskDefinition, ITaskFolder, ITaskService,
    ITaskSettings, ITrigger, ITriggerCollection, TASK_LOGON_INTERACTIVE_TOKEN,
};

#[derive(Debug, PartialEq)]
/// A schedule is created by a ScheduleBuilder. Once created, the Schedule can be registered with
/// the Windows Task Scheduler.
pub struct Schedule {
    pub(crate) actions: IActionCollection,
    pub(crate) force_start_boundary: bool,
    pub(crate) registration_info: IRegistrationInfo,
    pub(crate) settings: ITaskSettings,
    pub(crate) task_definition: ITaskDefinition,
    pub(crate) task_service: ITaskService,
    pub(crate) trigger: Option<ITrigger>,
    pub(crate) triggers: ITriggerCollection,
    //repetition: IRepetitionPattern,
}

/// Task Creation constants
/// see <https://docs.microsoft.com/en-us/windows/win32/api/taskschd/ne-taskschd-task_creation>
#[derive(Debug, PartialEq)]
pub enum TaskCreationFlags {
    /// The Task Scheduler service registers the task as a new task.
    Create = 2,
    /// The Task Scheduler service either registers the task as a new task or as an updated version 
    /// if the task already exists. Equivalent to `TaskCreationFlags.Create | TaskCreationFlags.Update`
    CreateOrUpdate = 6,
    /// The Task Scheduler service registers the disabled task. A disabled task cannot run until it is enabled.
    Disable = 8,
    /// The Task Scheduler service is prevented from adding the allow access-control entry (ACE) 
    /// for the context principal. When the `register()` function is called with this flag to
    /// update a task, the Task Scheduler service does not add the ACE for the new context principal
    /// and does not remove the ACE from the old context principal.
    DontAddPrincipalAce = 10,
    /// The Task Scheduler service creates the task, but ignores the registration triggers in the task.
    /// By ignoring the registration triggers, the task will not execute when it is registered
    /// unless a time-based trigger causes it to execute on registration.
    IgnoreRegistrationTriggers = 20,
    /// The Task Scheduler service registers the task as an updated version of an existing task.
    /// When a task with a registration trigger is updated, the task will execute after the update occurs.
    Update = 4,
    /// The Task Scheduler service checks the syntax of the XML that describes the task but does not
    /// register the task. This constant cannot be combined with the TASK_CREATE, TASK_UPDATE, or
    ValidateOnly = 1,
}

impl Schedule {
    /// Registers the schedule
    pub fn register(self, task_name: &str, flags: i32) -> Result<(), Box<dyn std::error::Error>> {
        unsafe {
            let folder: ITaskFolder = self.task_service.GetFolder("\\")?;
            folder.RegisterTaskDefinition(
                BSTR::from(task_name),
                self.task_definition,
                flags,
                // TODO allow user to specify creds
                None,
                None,
                TASK_LOGON_INTERACTIVE_TOKEN,
                None,
            )?;
        }

        Ok(())
    }
}
