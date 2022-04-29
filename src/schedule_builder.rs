use crate::error::{InitializationError, InvalidOperationError};
use crate::schedule::Schedule;
use windows::core::Interface;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::VARIANT;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_MULTITHREADED,
};
use windows::Win32::System::TaskScheduler::{
    IAction, IActionCollection, IBootTrigger, IDailyTrigger, IExecAction, IIdleTrigger,
    ILogonTrigger, IRegistrationInfo, IRepetitionPattern, ITaskDefinition, ITaskFolder,
    ITaskService, ITaskSettings, ITimeTrigger, ITrigger, ITriggerCollection, IWeeklyTrigger,
    TaskScheduler, TASK_ACTION_EXEC, TASK_LOGON_INTERACTIVE_TOKEN, TASK_TRIGGER_BOOT,
    TASK_TRIGGER_DAILY, TASK_TRIGGER_IDLE, TASK_TRIGGER_LOGON, TASK_TRIGGER_TIME,
    TASK_TRIGGER_WEEKLY,
};

/* triggers */
pub struct Base {}
pub struct Boot {}
pub struct Daily {}
pub struct Logon {}
pub struct Time {}
pub struct Weekly {}

#[derive(Debug)]
pub struct ScheduleBuilder<Frequency = Base> {
    pub(crate) frequency: std::marker::PhantomData<Frequency>,
    pub(crate) schedule: Schedule,
}

impl ScheduleBuilder<Base> {
    /// Create a new base builder.
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new();
    /// ```
    pub fn new() -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            // On error of unsafe, CoUnintialize!
            CoInitializeEx(std::ptr::null_mut(), COINIT_MULTITHREADED)?;

            let task_service: ITaskService = CoCreateInstance(&TaskScheduler, None, CLSCTX_ALL)?;
            task_service.Connect(
                VARIANT::default(),
                VARIANT::default(),
                VARIANT::default(),
                VARIANT::default(),
            )?;

            let task_definition: ITaskDefinition = task_service.NewTask(0)?;
            let triggers: ITriggerCollection = task_definition.Triggers()?;
            let registration_info: IRegistrationInfo = task_definition.RegistrationInfo()?;
            let actions: IActionCollection = task_definition.Actions()?;
            let settings: ITaskSettings = task_definition.Settings()?;

            Ok(Self {
                frequency: std::marker::PhantomData::<Base>,
                schedule: Schedule {
                    actions,
                    force_start_boundary: false,
                    registration_info,
                    settings,
                    task_service,
                    task_definition,
                    trigger: None,
                    triggers,
                },
            })
        }
    }

    /// Creates a builder for a boot trigger.
    ///
    /// # Example
    ///
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_boot();
    /// ```
    pub fn create_boot(self) -> ScheduleBuilder<Boot> {
        ScheduleBuilder::<Boot> {
            frequency: std::marker::PhantomData::<Boot>,
            schedule: self.schedule,
        }
    }

    /// Creates a builder for a daily trigger.
    ///
    /// # Example
    ///
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_daily();
    /// ```
    pub fn create_daily(mut self) -> ScheduleBuilder<Daily> {
        self.schedule.force_start_boundary = true;
        ScheduleBuilder::<Daily> {
            frequency: std::marker::PhantomData::<Daily>,
            schedule: self.schedule,
        }
    }

    pub fn create_time(mut self) -> ScheduleBuilder<Time> {
        self.schedule.force_start_boundary = true;
        ScheduleBuilder::<Time> {
            frequency: std::marker::PhantomData::<Time>,
            schedule: self.schedule,
        }
    }

    pub fn create_weekly(self) -> ScheduleBuilder<Weekly> {
        ScheduleBuilder::<Weekly> {
            frequency: std::marker::PhantomData::<Weekly>,
            schedule: self.schedule,
        }
    }
}

impl<Frequency> ScheduleBuilder<Frequency> {
    pub fn action(self, action: Action) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let i_action: IAction = self.schedule.actions.Create(TASK_ACTION_EXEC)?;
            let i_exec_action: IExecAction = i_action.cast()?;

            i_exec_action.SetPath(action.path)?;
            i_exec_action.SetId(action.id)?;
            i_exec_action.SetWorkingDirectory(action.working_dir)?;
            i_exec_action.SetArguments(action.args)?;
        }
        Ok(self)
    }

    /// Sets the author for this trigger.
    /// _optional_
    ///
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builer().new()
    ///     .create_daily()
    ///     .author("Alice")
    ///     .build();
    /// ```
    pub fn author(self, author: &str) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            self.schedule
                .registration_info
                .SetAuthor(BSTR::from(author))?;
        }
        Ok(self)
    }

    pub fn build(self) -> Result<Schedule, Box<dyn std::error::Error>> {
        // TODO validate folder & task & action are Some
        if self.schedule.trigger.is_none() {
            return Err(Box::new(InvalidOperationError {
                message: "Folder or trigger not set, cannot create scheduled task".to_string(),
            }));
        }
        Ok(self.schedule)
    }

    /// Sets the description for this trigger.
    /// _optional_
    ///
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_daily()
    ///     .description("This is my trigger")
    ///     .build();
    /// ```
    pub fn description(self, description: &str) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            self.schedule
                .registration_info
                .SetDescription(BSTR::from(description))?;
        }
        Ok(self)
    }

    pub fn uninitialize(self) {
        unsafe {
            CoUninitialize();
        }
    }

    pub fn execution_time_limit(
        self,
        time_limit: &str,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                trigger.SetExecutionTimeLimit(time_limit)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger"
                        .to_string(),
            }))
        }
    }

    /// Specifies the date and time when the trigger is activated. This call is required on
    /// Calendar triggers and Time Triggers.
    /// If `start_boundary` is not called on a required trigger, the start boundary will be set to
    /// `now` when the trigger is registered.
    /// `start_boundary`'s `start` parameter takes a rfc3339 formatted string (ie: 2007-01-01T08:00:00).
    ///
    /// ## References
    /// https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-startboundary-triggerbasetype-element
    /// https://datatracker.ietf.org/doc/html/rfc3339
    pub fn start_boundary(mut self, start: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                trigger.SetStartBoundary(start)?;
            }
            self.schedule.force_start_boundary = false;
            Ok(self)
        } else {
            self.uninitialize();
            Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger"
                        .to_string(),
            }))
        }
    }

    /// Specifies the date and time when the trigger is deactivated. The trigger cannot start the task after it is deactivated.
    /// `end_boundary`'s `end` parameter takes an rfc3339 formatted string (ie: 2007-01-01T08:00:00)
    ///
    /// ## References
    /// https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-endboundary-triggerbasetype-element
    /// https://datatracker.ietf.org/doc/html/rfc3339
    pub fn end_boundary(self, end: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                trigger.SetEndBoundary(end)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger"
                        .to_string(),
            }))
        }
    }

    //pub fn set_repetition(&self, IRepetitionPattern) {}
}

impl ScheduleBuilder<Boot> {
    pub fn trigger(mut self, id: &str, enabled: i16) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_BOOT)?;
            let i_boot_trigger: IBootTrigger = trigger.cast::<IBootTrigger>()?;
            i_boot_trigger.SetId(id)?;
            i_boot_trigger.SetEnabled(enabled)?;
            // Default start boundary to now()
            self.schedule.trigger = Some(i_boot_trigger.into());
        }

        Ok(self)
    }

    pub fn delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                let i_boot_trigger: IBootTrigger = trigger.cast::<IBootTrigger>()?;
                i_boot_trigger.SetDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }))
        }
    }
}

impl ScheduleBuilder<Daily> {
    pub fn trigger(mut self, id: &str, enabled: i16) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_DAILY)?;
            let i_daily_trigger: IDailyTrigger = trigger.cast::<IDailyTrigger>()?;
            i_daily_trigger.SetId(id)?;
            i_daily_trigger.SetEnabled(enabled)?;
            self.schedule.trigger = Some(i_daily_trigger.into());
        }
        Ok(self)
    }

    pub fn days_interval(self, days: i16) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_daily_trigger: IDailyTrigger = i_trigger.cast::<IDailyTrigger>()?;
                i_daily_trigger.SetDaysInterval(days)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            return Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }));
        }
    }

    /// Specifies the delay time that is randomly added to the start time of the trigger.
    /// The format for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2 day, 5 second delay).
    pub fn random_delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_daily_trigger: IDailyTrigger = i_trigger.cast::<IDailyTrigger>()?;
                i_daily_trigger.SetRandomDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            return Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }));
        }
    }
}

impl ScheduleBuilder<Logon> {
    pub fn trigger(mut self, id: &str, enabled: i16) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_LOGON)?;
            let i_logon_trigger: ILogonTrigger = trigger.cast::<ILogonTrigger>()?;
            i_logon_trigger.SetId(id)?;
            i_logon_trigger.SetEnabled(enabled)?;

            self.schedule.trigger = Some(i_logon_trigger.into());
        }
        Ok(self)
    }

    pub fn delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                let i_logon_trigger: ILogonTrigger = trigger.cast::<ILogonTrigger>()?;
                i_logon_trigger.SetDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }))
        }
    }

    pub fn user_id(self, id: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                let i_logon_trigger: ILogonTrigger = trigger.cast::<ILogonTrigger>()?;
                i_logon_trigger.SetUserId(id)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }))
        }
    }
}

impl ScheduleBuilder<Time> {
    /// Creates a time trigger
    /// It is important to note that a time trigger is different from other time-based triggers in that
    /// it is fired when the trigger is activated by its start boundary. Other time-based triggers are
    /// activated by their start boundary, but they do not start performing their actions
    /// until a scheduled date is reached.
    pub fn trigger(mut self, id: &str, enabled: i16) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_TIME)?;
            let i_time_trigger: ITimeTrigger = trigger.cast::<ITimeTrigger>()?;
            i_time_trigger.SetId(id)?;
            i_time_trigger.SetEnabled(enabled)?;

            self.schedule.trigger = Some(i_time_trigger.into());
        }
        Ok(self)
    }

    /// Specifies the delay time that is randomly added to the start time of the trigger.
    /// The format for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2 day, 5 second delay).
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-randomdelay-timetriggertype-element
    pub fn random_delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_time_trigger: ITimeTrigger = i_trigger.cast::<ITimeTrigger>()?;
                i_time_trigger.SetRandomDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            return Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }));
        }
    }
}

impl ScheduleBuilder<Weekly> {
    pub fn trigger(mut self, id: &str, enabled: i16) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_WEEKLY)?;
            let i_weekly_trigger: IWeeklyTrigger = trigger.cast::<IWeeklyTrigger>()?;
            i_weekly_trigger.SetId(id)?;
            i_weekly_trigger.SetEnabled(enabled)?;

            self.schedule.trigger = Some(i_weekly_trigger.into());
        }
        Ok(self)
    }

    /// Specifies the day(s) of the week to run the command.
    /// days is a bitwise mask that indicated the days of the week on which the task runs.
    /// |Weekday|Hex Value|Decimal Value|
    /// |---|---|
    /// |Sunday|0x01|1|
    /// |Monday|0x02|2|
    /// |Tuesday|0x04|4|
    /// |Wednesday|0x08|8|
    /// |Thursday|0x10|16|
    /// |Friday|0x20|32|
    /// |Saturday|0x40|64|
    pub fn days_of_week(self, days: i16) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_weekly_trigger: IWeeklyTrigger = i_trigger.cast::<IWeeklyTrigger>()?;
                i_weekly_trigger.SetDaysOfWeek(days)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            return Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }));
        }
    }

    /// Sets the interval between the weeks in the schedule.
    /// For example settings to 1 runs every week; setting to 2 runs every other week.
    pub fn weeks_interval(self, weeks: i16) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_weekly_trigger: IWeeklyTrigger = i_trigger.cast::<IWeeklyTrigger>()?;
                i_weekly_trigger.SetWeeksInterval(weeks)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            return Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }));
        }
    }

    /// Specifies the delay time that is randomly added to the start time of the trigger.
    /// The format for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2 day, 5 second delay).
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-randomdelay-timetriggertype-element
    pub fn random_delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_weekly_trigger: IWeeklyTrigger = i_trigger.cast::<IWeeklyTrigger>()?;
                i_weekly_trigger.SetRandomDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            return Err(Box::new(InvalidOperationError {
                message:
                    "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
                        .to_string(),
            }));
        }
    }
}

/* actions */

#[derive(Debug, Clone)]
pub struct Action {
    id: BSTR,
    path: BSTR,
    working_dir: BSTR,
    args: BSTR,
}

impl Action {
    pub fn new(id: &str, path: &str, working_dir: &str, args: &str) -> Self {
        Self {
            id: id.into(),
            path: path.into(),
            working_dir: working_dir.into(),
            args: args.into(),
        }
    }
}
