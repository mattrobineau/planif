use windows::Win32::Foundation::BSTR;
use windows::Win32::System::TaskScheduler::{
    IActionCollection, IRegistrationInfo, ITaskDefinition, ITaskFolder, ITaskService,
    ITaskSettings, ITrigger, ITriggerCollection, TASK_LOGON_INTERACTIVE_TOKEN,
};

#[derive(Debug, PartialEq)]
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
    Create = 2,
    CreateOrUpdate = 6,
    Disable = 8,
    DontAddPrincipalAce = 10,
    IgnoreRegistrationTriggers = 20,
    Update = 4,
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
