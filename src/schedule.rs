use windows::core::BSTR;
use windows::Win32::System::TaskScheduler::{
    IActionCollection, IRegistrationInfo, ITaskDefinition, ITaskFolder, ITaskService,
    ITaskSettings, ITrigger, ITriggerCollection, TASK_LOGON_INTERACTIVE_TOKEN,
};
use windows::Win32::System::Com::VARIANT;

use crate::com::ComRuntime;

/// Marker type for base [`Schedule<Unregistered>`]
pub struct Unregistered {}

/// Marker type for registered [`Schedule<Registered>`]
pub struct Registered {}

#[derive(Debug, PartialEq)]
/// A schedule is created by a [schedule builder](crate::schedule_builder). Once created, the
/// Schedule can be registered with the Windows Task Scheduler.
pub struct Schedule<Kind = Unregistered> {
    pub(crate) kind: std::marker::PhantomData<Kind>,
    pub(crate) task_folder: ITaskFolder,
    pub(crate) actions: IActionCollection,
    pub(crate) force_start_boundary: bool,
    pub(crate) registration_info: IRegistrationInfo,
    pub(crate) settings: ITaskSettings,
    pub(crate) task_definition: ITaskDefinition,
    // pub(crate) task_service: ITaskService,
    pub(crate) trigger: Option<ITrigger>,
    pub(crate) triggers: ITriggerCollection,
    pub(crate) com_runtime: ComRuntime,
    //repetition: IRepetitionPattern,
}

impl Schedule<Unregistered> {
    /// Registers the schedule. Flags can be set by using the [TaskCreationFlags](crate::enums::TaskCreationFlags) enum.
    pub fn register(self, task_name: &str, flags: i32) -> Result<(), Box<dyn std::error::Error>> {
        unsafe {
            self.task_folder.RegisterTaskDefinition(
                &BSTR::from(task_name),
                &self.task_definition,
                flags,
                // TODO allow user to specify creds
                VARIANT::default(),
                VARIANT::default(),
                TASK_LOGON_INTERACTIVE_TOKEN,
                VARIANT::default(),
            )?;
        }
        
        Ok(())
    }
}

/// TaskScheduler represents the actions you can take for using the Windows Task Scheduler.
/// For example: Creating new schedules, fetching the COM, etc.
pub struct TaskScheduler {
    com: ComRuntime
}

impl TaskScheduler {
    /// Create's a new TaskScheduler
    /// # Example
    /// ```
    /// use planif::schedule::TaskScheduler;
    ///
    /// let ts = TaskScheduler::new().unwrap();
    /// ```
    ///
    pub fn new() -> Result<Self, Box<dyn std::error::Error>> {
        Ok(Self {
            com: ComRuntime::new()?
        })
    }

    /// Returns the TaskScheduler's com runtime
    /// # Example
    /// ```
    /// use planif::schedule::TaskScheduler;
    /// 
    /// let ts = TaskScheduler::new().unwrap();
    /// let com = ts.get_com();
    /// ```
    pub fn get_com(&self) -> ComRuntime {
        self.com.clone()
    }
}

impl Schedule<Registered> {
    /// tesst
    pub fn test() -> bool {
        true
    }
    /// more test
    pub fn path(&self) -> String {
        unsafe {
            self.task_folder.Path().unwrap().to_string()
        }
    }
}
