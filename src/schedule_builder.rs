use crate::{
    enums::{DayOfMonth, DayOfWeek, Month, WeekOfMonth},
    error::{InitializationError, InvalidOperationError},
    schedule::Schedule,
    settings::{PrincipalSettings, RunLevel},
};
use windows::core::Interface;
use windows::Win32::Foundation::BSTR;
use windows::Win32::System::Com::VARIANT;
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, COINIT_MULTITHREADED,
};
use windows::Win32::System::TaskScheduler::{
    IAction, IActionCollection, IBootTrigger, IDailyTrigger, IEventTrigger, IEventTrigger,
    IExecAction, IIdleTrigger, ILogonTrigger, IMonthlyDOWTrigger, IMonthlyTrigger, IPrincipal,
    IRegistrationInfo, IRegistrationTrigger, IRepetitionPattern, ITaskDefinition, ITaskFolder,
    ITaskService, ITaskSettings, ITimeTrigger, ITrigger, ITriggerCollection, IWeeklyTrigger,
    TaskScheduler, TASK_ACTION_EXEC, TASK_LOGON_INTERACTIVE_TOKEN, TASK_LOGON_TYPE,
    TASK_RUNLEVEL_TYPE, TASK_TRIGGER_BOOT, TASK_TRIGGER_DAILY, TASK_TRIGGER_EVENT,
    TASK_TRIGGER_IDLE, TASK_TRIGGER_LOGON, TASK_TRIGGER_MONTHLY, TASK_TRIGGER_MONTHLYDOW,
    TASK_TRIGGER_REGISTRATION, TASK_TRIGGER_TIME, TASK_TRIGGER_WEEKLY,
};

/* triggers */
pub struct Base {}
pub struct Boot {}
pub struct Daily {}
pub struct Event {}
pub struct Idle {}
pub struct Logon {}
pub struct Monthly {}
pub struct MonthlyDOW {}
pub struct Registration {}
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

    /// Creates a builder for an idle trigger.
    ///
    /// # Example
    ///
    /// ``` let schedule: Schedule = Schedule::builder().new()
    ///         .create_idle();
    /// ```
    pub fn create_idle(self) -> ScheduleBuilder<Idle> {
        ScheduleBuilder::<Idle> {
            frequency: std::marker::PhantomData::<Idle>,
            schedule: self.schedule,
        }
    }

    /// Creates a builder for a logon trigger.
    ///
    /// # Example
    ///
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_logon();
    /// ```
    pub fn create_logon(self) -> ScheduleBuilder<Logon> {
        ScheduleBuilder::<Logon> {
            frequency: std::marker::PhantomData::<Logon>,
            schedule: self.schedule,
        }
    }

    /// Creates a builder for a monthly trigger.
    ///
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly();
    /// ```
    pub fn create_monthly(self) -> ScheduleBuilder<Monthly> {
        ScheduleBuilder::<Monthly> {
            frequency: std::marker::PhantomData::<Monthly>,
            schedule: self.schedule,
        }
    }

    /// Creates a builder for a monthly day of week trigger.
    ///
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly_dow();
    /// ```
    pub fn create_monthly_dow(self) -> ScheduleBuilder<MonthlyDOW> {
        ScheduleBuilder::<MonthlyDOW> {
            frequency: std::marker::PhantomData::<MonthlyDOW>,
            schedule: self.schedule,
        }
    }

    /// Creates a builder for a time trigger.
    ///
    /// # Example
    ///
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_time();
    /// ```
    pub fn create_time(mut self) -> ScheduleBuilder<Time> {
        self.schedule.force_start_boundary = true;
        ScheduleBuilder::<Time> {
            frequency: std::marker::PhantomData::<Time>,
            schedule: self.schedule,
        }
    }

    /// Creates a builder for a weekly trigger.
    ///
    /// # Example
    ///
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_weekly();
    /// ```
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

    /// Returns the schedule
    pub fn build(self) -> Result<Schedule, Box<dyn std::error::Error>> {
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

    /// Closes the COM library on the current thread, unloads all DLLs loaded
    /// by the thread, frees any other resources that the thread maintains, and
    /// forces all RPC connections on the thread to close.
    pub fn uninitialize(self) {
        unsafe {
            CoUninitialize();
        }
    }

    /// The amount of time that is allowed to complete the task.
    /// The format for this string is PnYnMnDTnHnMnS, where nY is the number of years,
    /// nM is the number of months, nD is the number of days, 'T' is the date/time separator,
    /// nH is the number of hours, nM is the number of minutes, and nS is the number of
    /// seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies one month,
    /// four days, two hours, and five minutes). A value of PT0S will enable the task to run indefinitely.
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

    /// Sets the repetition duration for a task.
    /// If you specify a repetition duration for a task, you must also specify the repetition interval.
    /// If you register a task that contains a trigger with a repetition interval equal to one minute
    /// and a repetition duration equal to four minutes, the task will be launched five times. The five
    /// repetitions can be defined by the following pattern.
    ///
    /// A task starts at the beginning of the first minute.
    /// - The next task starts at the end of the first minute.
    /// - The next task starts at the end of the second minute.
    /// - The next task starts at the end of the third minute.
    /// - The next task starts at the end of the fourth minute.
    ///
    /// # Parameters
    /// ## duration
    /// How long the pattern is repeated. The format for this string is PnYnMnDTnHnMnS, where nY is
    /// the number of years, nM is the number of months, nD is the number of days, 'T' is the date/time
    /// separator, nH is the number of hours, nM is the number of minutes, and nS is the number of seconds
    /// (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies one month, four days, two hours,
    /// and five minutes). The minimum time allowed is one minute.
    ///
    /// ## interval
    /// The amount of time between each restart of the task. The format for this string is
    /// P<days>DT<hours>H<minutes>M<seconds>S (for example, "PT5M" is 5 minutes, "PT1H" is 1 hour, and "PT20M"
    /// is 20 minutes). The maximum time allowed is 31 days, and the minimum time allowed is 1 minute.
    ///
    /// # stop_at_duration_end
    /// A Boolean value that indicates if a running instance of the task is stopped at the end of the repetition
    /// pattern duration.
    ///
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/repetitionpattern
    pub fn repetition(
        self,
        duration: &str,
        interval: &str,
        stop_at_duration_end: bool,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                let repetition: IRepetitionPattern = trigger.Repetition()?;
                repetition.SetDuration(duration)?;
                repetition.SetInterval(interval)?;
                repetition.SetStopAtDurationEnd(stop_at_duration_end as i16)?;
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

    /// Sets the task's principal
    /// When specifying an account, remember to properly use the double backslash in code to specify the
    /// domain and user name. For example, use DOMAIN\\UserName to specify a value for the UserId property.
    ///
    /// # reference
    /// https://docs.microsoft.com/en-us/windows/win32/taskschd/principal
    pub fn principal(
        self,
        settings: PrincipalSettings,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let principal: IPrincipal = self.schedule.task_definition.Principal()?;
            principal.SetDisplayName(settings.display_name)?;

            if settings.group_id.is_some() && settings.user_id.is_some() {
                return Err(Box::new(InvalidOperationError {
                    message: "Invalid operation: group_id and user_id are mutually exclusive and cannot both be set."
                        .to_string(),
                }));
            } else {
                if let Some(gid) = settings.group_id {
                    principal.SetGroupId(gid)?;
                } else if let Some(uid) = settings.user_id {
                    principal.SetUserId(uid)?;
                }
            }

            principal.SetId(settings.id)?;
            principal.SetLogonType(TASK_LOGON_TYPE(settings.logon_type as i32))?;
            principal.SetRunLevel(TASK_RUNLEVEL_TYPE(settings.run_level as i32))?;
            Ok(self)
        }
    }
}

impl ScheduleBuilder<Boot> {
    /// Create a task that is started when the operating system is booted,
    /// and boot trigger tasks are set to start when the Task Scheduler service starts.
    /// Only a member of the Administrators group can create a task with a boot trigger.
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/boottrigger
    pub fn trigger(mut self, id: &str, enabled: bool) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_BOOT)?;
            let i_boot_trigger: IBootTrigger = trigger.cast::<IBootTrigger>()?;
            i_boot_trigger.SetId(id)?;
            i_boot_trigger.SetEnabled(enabled.into())?;
            // Default start boundary to now()
            self.schedule.trigger = Some(i_boot_trigger.into());
        }

        Ok(self)
    }

    /// Specifies a value that indicates the amount of time between when the user logs on and when the task is started.
    /// The format for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2 day, 5 second delay).
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/logontrigger-delay
    pub fn delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                let i_boot_trigger: IBootTrigger = trigger.cast::<IBootTrigger>()?;
                i_boot_trigger.SetDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }
}

impl ScheduleBuilder<Daily> {
    /// Creates a daily trigger
    /// The time of day that the task is started is set by the start_boundary method.
    /// If `start_boundary()` is not set, it will default to `now` when the `schedule` is `registered()`
    ///An interval of 1 produces a daily schedule. An interval of 2 produces an every other day schedule and so on.
    pub fn trigger(mut self, id: &str, enabled: bool) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_DAILY)?;
            let i_daily_trigger: IDailyTrigger = trigger.cast::<IDailyTrigger>()?;
            i_daily_trigger.SetId(id)?;
            i_daily_trigger.SetEnabled(enabled.into())?;
            self.schedule.trigger = Some(i_daily_trigger.into());
        }
        Ok(self)
    }

    /// Sets the interval for days.
    /// ie: An interval of 1 produces a daily schedule. An interval of 2 produces an every-other day schedule. Etc.
    pub fn days_interval(self, days: i16) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_daily_trigger: IDailyTrigger = i_trigger.cast::<IDailyTrigger>()?;
                i_daily_trigger.SetDaysInterval(days)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Specifies the delay time that is randomly added to the start time of the trigger.
    /// The format for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2 day, 5 second delay).
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-randomdelay-timetriggertype-element
    pub fn random_delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_daily_trigger: IDailyTrigger = i_trigger.cast::<IDailyTrigger>()?;
                i_daily_trigger.SetRandomDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }
}

impl ScheduleBuilder<Idle> {
    /// Create an idle trigger
    pub fn trigger(mut self, id: &str, enabled: bool) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_IDLE)?;
            let i_idle_trigger: IIdleTrigger = trigger.cast::<IIdleTrigger>()?;
            i_idle_trigger.SetId(id)?;
            i_idle_trigger.SetEnabled(enabled.into())?;
            self.schedule.trigger = Some(i_idle_trigger.into());
        }
        Ok(self)
    }
}

impl ScheduleBuilder<Logon> {
    /// Create a logon trigger.
    pub fn trigger(mut self, id: &str, enabled: bool) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_LOGON)?;
            let i_logon_trigger: ILogonTrigger = trigger.cast::<ILogonTrigger>()?;
            i_logon_trigger.SetId(id)?;
            i_logon_trigger.SetEnabled(enabled.into())?;

            self.schedule.trigger = Some(i_logon_trigger.into());
        }
        Ok(self)
    }

    /// Specifies a value that indicates the amount of time between when the user logs on and when the task is started.
    /// The format for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2 day, 5 second delay).
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/logontrigger-delay
    pub fn delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                let i_logon_trigger: ILogonTrigger = trigger.cast::<ILogonTrigger>()?;
                i_logon_trigger.SetDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// The identifier of the user. For example, "MyDomain\MyName" or for a local account, "Administrator".
    /// _required_
    /// This property can be in one of the following formats:
    ///  - User name or SID: The task is started when the user logs on to the computer.
    ///  - NULL: The task is started when any user logs on to the computer.
    ///  see https://docs.microsoft.com/en-us/windows/win32/taskschd/logontrigger-userid
    pub fn user_id(self, id: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                let i_logon_trigger: ILogonTrigger = trigger.cast::<ILogonTrigger>()?;
                i_logon_trigger.SetUserId(id)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }
}

impl ScheduleBuilder<Monthly> {
    /// Set the days of the month during which the task runs.
    /// # Example
    /// ```
    /// use planif::enums::DayOfMonth;
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly()
    ///     .days_of_months(vec![DayOfMonth::Day(1), DayOfMonth::Day(15), DayOfMonth::Day(31)]);
    /// ```
    pub fn days_of_months(
        mut self,
        days: Vec<DayOfMonth>,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            let result = days.iter().any(|&x| match &x {
                DayOfMonth::Day(int) => int < &1 || int > &31,
                DayOfMonth::Last => false,
            });

            let bitwise = days.into_iter().fold(0, |acc, item| {
                let day: i32 = item.into();
                acc + (1 << day - 1)
            });

            unsafe {
                let i_monthly_trigger: IMonthlyTrigger = i_trigger.cast::<IMonthlyTrigger>()?;
                i_monthly_trigger.SetDaysOfMonth(bitwise);
            }

            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Set the months of the year during which the task runs.
    /// # Example
    /// ```
    /// use planif::enums::Month;
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly()
    ///     .months_of_year(vec![Month::January, Month::June, Month::December]);
    /// ```
    pub fn months_of_year(
        mut self,
        months: Vec<Month>,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            let bitwise: i16 = months.into_iter().fold(0, |acc, item| acc + item as i16);

            unsafe {
                let i_monthly_trigger: IMonthlyTrigger = i_trigger.cast::<IMonthlyTrigger>()?;
                i_monthly_trigger.SetMonthsOfYear(bitwise);
            }

            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Specifies the delay time that is randomly added to the start time of the trigger.
    /// The format for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2 day, 5 second delay).
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-randomdelay-timetriggertype-element
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly()
    ///     .random_delay("P2DT5S");
    /// ```
    pub fn random_delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_monthly_trigger: IMonthlyTrigger = i_trigger.cast::<IMonthlyTrigger>()?;
                i_monthly_trigger.SetRandomDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Sets the task to be run on the last day of the month, regardless of the actual date of
    /// that day.
    ///
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly()
    ///     .run_on_last_day(true);
    /// ```
    pub fn run_on_last_day(self, is_run: bool) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_monthly_trigger: IMonthlyTrigger = i_trigger.cast::<IMonthlyTrigger>()?;
                i_monthly_trigger.SetRunOnLastDayOfMonth(is_run as i16);
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Creates a trigger based on a monthly schedule, for example, the task starts on specific
    /// days of specific months
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly()
    ///     .trigger("MyTrigger", true);
    /// ```
    pub fn trigger(mut self, id: &str, enabled: bool) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_MONTHLY)?;
            let i_monthly_trigger: IMonthlyTrigger = trigger.cast::<IMonthlyTrigger>()?;
            i_monthly_trigger.SetId(id)?;
            i_monthly_trigger.SetEnabled(enabled.into())?;
            self.schedule.trigger = Some(i_monthly_trigger.into());
        }
        Ok(self)
    }
}

impl ScheduleBuilder<MonthlyDOW> {
    /// Sets the days of the week during which the task runs.
    /// # Example
    /// ```
    /// use planif::enums::DayOfWeek;
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly_dow()
    ///     .trigger("MonthlyDOWTrigger", true)
    ///     .days_of_week(vec![DayOfWeek::Sunday, DayOfWeek::Thursday]);
    /// ```
    pub fn days_of_week(self, days: Vec<DayOfWeek>) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            let bitwise: i16 = days.into_iter().fold(0, |acc, item| acc + item as i16);
            unsafe {
                let i_monthly_dow_trigger: IMonthlyDOWTrigger =
                    trigger.cast::<IMonthlyDOWTrigger>()?;
                i_monthly_dow_trigger.SetDaysOfWeek(bitwise)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Set the months of the year during which the task runs.
    /// # Example
    /// ```
    /// use planif::enums::Month;
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly_dow()
    ///     .trigger("MonthlyDOWTrigger", true)
    ///     .months_of_year(vec![Month::January, Month::June, Month::December]);
    /// ```
    pub fn months_of_year(
        mut self,
        months: Vec<Month>,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            let bitwise: i16 = months.into_iter().fold(0, |acc, item| acc + item as i16);

            unsafe {
                let i_monthly_dow_trigger: IMonthlyDOWTrigger =
                    i_trigger.cast::<IMonthlyDOWTrigger>()?;
                i_monthly_dow_trigger.SetMonthsOfYear(bitwise);
            }

            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Specifies the delay time that is randomly added to the start time of the trigger.
    /// The format for this string is P<days>DT<hours>H<minutes>M<seconds>S (for example, P2DT5S is a 2 day, 5 second delay).
    /// see https://docs.microsoft.com/en-us/windows/win32/taskschd/taskschedulerschema-randomdelay-timetriggertype-element
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly_dow()
    ///     .trigger("MonthlyDOWTrigger", true)
    ///     .random_delay("P2DT5S");
    /// ```
    pub fn random_delay(self, delay: &str) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            unsafe {
                let i_monthly_dow_trigger: IMonthlyDOWTrigger =
                    i_trigger.cast::<IMonthlyDOWTrigger>()?;
                i_monthly_dow_trigger.SetRandomDelay(delay)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Sets the task to be run on the last day of the month, regardless of the actual date of
    /// that day.
    ///
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly_dow()
    ///     .trigger("MonthlyDOWTrigger", true)
    ///     .run_on_last_week(true);
    /// ```
    pub fn run_on_last_week(self, is_run: bool) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            unsafe {
                let i_monthly_dow_trigger: IMonthlyDOWTrigger =
                    trigger.cast::<IMonthlyDOWTrigger>()?;
                i_monthly_dow_trigger.SetRunOnLastWeekOfMonth(is_run as i16);
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Sets the weeks of the month during which the task runs.
    ///
    /// # Example
    /// ```
    /// use planif::enums::WeekOfMonth;
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly_dow()
    ///     .trigger("MonthlyDOWTrigger", true)
    ///     .weeks_of_month(vec![WeekOfMonth::Third]);
    /// ```
    pub fn weeks_of_month(
        mut self,
        weeks: Vec<WeekOfMonth>,
    ) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(trigger) = &self.schedule.trigger {
            let bitwise: i16 = weeks.into_iter().fold(0, |acc, item| acc + item as i16);
            unsafe {
                let i_monthly_dow_trigger: IMonthlyDOWTrigger =
                    trigger.cast::<IMonthlyDOWTrigger>()?;
                i_monthly_dow_trigger.SetWeeksOfMonth(bitwise);
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
        }
    }

    /// Creates a trigger that starts a task on a monthly day-of-week schedule. For example, the
    /// task starts on every first Thursday, May through October.
    /// # Example
    /// ```
    /// let schedule: Schedule = Schedule::builder().new()
    ///     .create_monthly_dow()
    ///     .trigger("MonthlyDOWTrigger", true);
    /// ```
    pub fn trigger(mut self, id: &str, enabled: bool) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_MONTHLYDOW)?;
            let i_monthly_dow_trigger: IMonthlyDOWTrigger = trigger.cast::<IMonthlyDOWTrigger>()?;
            i_monthly_dow_trigger.SetId(id)?;
            i_monthly_dow_trigger.SetEnabled(enabled.into())?;
        }
        Ok(self)
    }
}

impl ScheduleBuilder<Time> {
    /// Creates a time trigger
    /// It is important to note that a time trigger is different from other time-based triggers in that
    /// it is fired when the trigger is activated by its start boundary. Other time-based triggers are
    /// activated by their start boundary, but they do not start performing their actions
    /// until a scheduled date is reached.
    pub fn trigger(mut self, id: &str, enabled: bool) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_TIME)?;
            let i_time_trigger: ITimeTrigger = trigger.cast::<ITimeTrigger>()?;
            i_time_trigger.SetId(id)?;
            i_time_trigger.SetEnabled(enabled.into())?;

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
            Err(trigger_uninitialised_error())
        }
    }
}

impl ScheduleBuilder<Weekly> {
    pub fn trigger(mut self, id: &str, enabled: bool) -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            let trigger = self.schedule.triggers.Create(TASK_TRIGGER_WEEKLY)?;
            let i_weekly_trigger: IWeeklyTrigger = trigger.cast::<IWeeklyTrigger>()?;
            i_weekly_trigger.SetId(id)?;
            i_weekly_trigger.SetEnabled(enabled.into())?;

            self.schedule.trigger = Some(i_weekly_trigger.into());
        }
        Ok(self)
    }

    /// Sets the days of the week during which the task runs.
    pub fn days_of_week(self, days: Vec<DayOfWeek>) -> Result<Self, Box<dyn std::error::Error>> {
        if let Some(i_trigger) = &self.schedule.trigger {
            let bitwise: i16 = days.into_iter().fold(0, |acc, item| acc + item as i16);
            unsafe {
                let i_weekly_trigger: IWeeklyTrigger = i_trigger.cast::<IWeeklyTrigger>()?;
                i_weekly_trigger.SetDaysOfWeek(bitwise)?;
            }
            Ok(self)
        } else {
            self.uninitialize();
            Err(trigger_uninitialised_error())
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
            Err(trigger_uninitialised_error())
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
            Err(trigger_uninitialised_error())
        }
    }
}

fn trigger_uninitialised_error() -> Box<dyn std::error::Error> {
    return Box::new(InvalidOperationError {
        message: "Trigger has not been created yet. Consider calling ScheduleBuilder.Trigger()"
            .to_string(),
    });
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
