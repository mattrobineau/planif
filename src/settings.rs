use std::fmt;

/// Values for task compatibility  
/// Task compatibility, which is set through the Compatibility property, should only be set to `Compatibility.V1`
/// if a task needs to be accessed or modified from a Windows XP, Windows Server 2003, or Windows 2000 computer.
/// Otherwise, it is recommended that Task Scheduler 2.0 compatibility be used because the task will have more features.
/// Tasks compatible with the AT command can only have one time trigger.
/// Tasks compatible with Task Scheduler 1.0 can only have a time trigger, a logon trigger, or a boot trigger, and the
/// task can only have an executable action.
///
/// See <https://docs.microsoft.com/en-us/windows/win32/taskschd/tasksettings-compatibility>
pub enum Compatibility {
    /// The task is compatible with the AT command.
    AT = 0,
    /// The task is compatible with Task Scheduler 1.0.
    V1,
    /// The task is compatible with Task Scheduler 2.0.
    V2,
}

pub(crate) use windows::Win32::System::TaskScheduler::TASK_COMPATIBILITY;
impl From<Compatibility> for TASK_COMPATIBILITY {
    fn from(item: Compatibility) -> Self {
        TASK_COMPATIBILITY(item as i32)
    }
}

/// Defines idle settings on a task
/// # Example
/// ```
/// use planif::settings::Settings;
/// // All values are set to `None` when using `new()`
/// let mut settings = Settings::new();
/// settings.idle_settings = Some(IdleSettings::new());
/// ```
#[allow(deprecated)]
pub struct IdleSettings {
    #[deprecated]
    /// This field is deprecated.
    /// Gets or sets a String that indicates the amount of time that the computer must be in an idle state before
    /// the task is run.
    ///
    /// A value that indicates the amount of time that the computer must be in an idle state before the task
    /// is run). The format for this string is PnYnMnDTnHnMnS, where nY is the number of years, nM is the number
    /// of months, nD is the number of days, 'T' is the date/time separator, nH is the number of hours,
    /// nM is the number of minutes, and nS is the number of seconds (for example, PT5M specifies 5 minutes
    /// and P1M4DT2H5M specifies one month, four days, two hours, and five minutes). The minimum value is
    /// one minute.
    pub idle_duration: Option<String>,
    /// Gets or sets a Boolean value that indicates whether the task is restarted when the computer cycles into an idle
    /// condition more than once.
    pub restart_on_idle: Option<bool>,
    /// Gets or sets a Boolean value that indicates that the Task Scheduler will terminate the task if the
    /// idle condition ends before the task is completed.
    pub stop_on_idle_end: Option<bool>,
    #[deprecated]
    /// This field is deprecated.
    /// Get or sets a String that indicates the amount of time that the Task Scheduler will wait for an idle
    /// condition to occur.
    ///
    /// The format for this String is PnYnMnDTnHnMnS, where nY is the number of years, nM is the number of months,
    /// nD is the number of days, 'T' is the date/time separator, nH is the number of hours, nM is the number
    /// of minutes, and nS is the number of seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies
    /// one month, four days, two hours, and five minutes). The minimum time allowed is 1 minute.
    pub wait_timeout: Option<String>,
}

#[allow(deprecated)]
impl IdleSettings {
    /// Creates a new IdleSettings struct with all values set to None.
    pub fn new() -> IdleSettings {
        IdleSettings {
            idle_duration: None,
            restart_on_idle: None,
            stop_on_idle_end: None,
            wait_timeout: None,
        }
    }
}

/// Values for the instance policy.
pub enum InstancesPolicy {
    /// Starts a new instance while an existing instance of the task is running.
    Parallel = 0,
    /// Starts a new instance of the task after all other instances of the task are complete.
    Queue,
    /// Does not start a new instance if an existing instance of the task is running.
    IgnoreNew,
    /// Stops an existing instance of the task before it starts new instance.
    StopExisting,
}

pub(crate) use windows::Win32::System::TaskScheduler::TASK_INSTANCES_POLICY;
impl From<InstancesPolicy> for TASK_INSTANCES_POLICY {
    fn from(item: InstancesPolicy) -> Self {
        TASK_INSTANCES_POLICY(item as i32)
    }
}

/// Values for the security logon method.
pub enum LogonType {
    /// The logon method is not specified. Used for non-NT credentials.
    None = 0,
    /// Use a password for logging on the user. The password must be supplied at registration time.
    Password,
    /// Use an existing interactive token to run a task. The user must log on using a service for user (S4U) logon.
    /// When an S4U logon is used, no password is stored by the system and there is no access to either the network
    /// or encrypted files.
    S4U,
    /// User must already be logged on. The task will be run only in an existing interactive session.
    InteractiveToken,
    /// Group activation. The [user_id](PrincipalSettings::user_id) field specifies the group.
    Group,
    /// Indicates that a Local System, Local Service, or Network Service account is being used as a security context
    /// to run the task.
    ServiceAccount,
    /// First use the interactive token. If the user is not logged on (no interactive token is available), then the password
    /// is used. The password must be specified when a task is registered. This flag is not recommended for new tasks because
    /// it is less reliable than [LogonType::Password](LogonType::Password).
    InteractiveTokenOrPassword,
}

/// Use to set a network profile identifier and name.
pub struct NetworkSettings {
    /// GUID value that identifies a network profile.
    pub id: String,
    /// The name of a network profile. The name is used for display purposes.
    pub name: String,
}

/// Use to set the settings for the principal
/// # Reference
/// <https://docs.microsoft.com/en-us/windows/win32/taskschd/principal>
pub struct PrincipalSettings {
    /// Gets or sets the name of the principal that is displayed in the Task Scheduler UI.
    pub display_name: String,
    /// Gets or sets the identifier of the user group that is required to run the tasks that are associated with the principal.
    /// Do not set this property if a user identifier is specified in the [user_id](PrincipalSettings::user_id) property.
    pub group_id: Option<String>,
    /// Gets or sets the identifier of the principal.
    pub id: String,
    /// Gets or sets the security logon method that is required to run the tasks that are associated with the principal.
    /// This property is valid only when a user identifier is specified by the [user_id](PrincipalSettings::user_id) property.
    pub logon_type: LogonType,
    /// Gets or sets the identifier that is used to specify the privilege level that is required to run the tasks
    /// that are associated with the principal.
    pub run_level: RunLevel,
    /// Gets or sets the user identifier that is required to run the tasks that are associated with the principal.
    /// Do not set this property if a group identifier is specified in the [group_id](PrincipalSettings::group_id) property.
    pub user_id: Option<String>,
}

/// Values for the identifier that is used to specify the privilege level that is required to run the tasks
/// that are associated with the principal.
pub enum RunLevel {
    /// Tasks will be run with the highest privileges.
    Highest = 1,
    /// Tasks will be run with the least privileges (LUA).
    LUA = 0,
}

/// Defines all available settings on a task.
///
/// # Example
/// ```
/// use planif::settings::Settings;
/// // All values are set to `None` when using `new()`
/// let mut settings = Settings::new();
/// settings.allow_demand_start = Some(true);
/// ```
///
/// # References
/// - <https://docs.microsoft.com/en-us/windows/win32/taskschd/tasksettings>
/// - <https://docs.microsoft.com/en-us/windows/win32/taskschd/tasksettings-priority>
/// - <https://docs.microsoft.com/en-us/windows/win32/procthread/scheduling-priorities>
/// - <https://docs.microsoft.com/en-us/windows/win32/taskschd/networksettings>
/// - <https://docs.microsoft.com/en-us/windows/win32/taskschd/idlesettings>
pub struct Settings {
    /// Gets or sets a Boolean value that indicates that the task can be started by using either the Run command
    /// or the Context menu.
    pub allow_demand_start: Option<bool>,
    /// Gets or sets a Boolean value that indicates that the task may be terminated by using TerminateProcess.
    pub allow_hard_terminate: Option<bool>,
    /// Gets or sets an integer value that indicates which version of Task Scheduler a task is compatible with.
    pub compatibility: Option<Compatibility>,
    /// Gets or sets the amount of time that the Task Scheduler will wait before deleting the task after it expires.
    ///
    /// A string that gets or sets the amount of time that the Task Scheduler will wait before deleting the task after
    /// it expires. The format for this string is PnYnMnDTnHnMnS, where nY is the number of years, nM is the number of
    /// months, nD is the number of days, 'T' is the date/time separator, nH is the number of hours, nM is the number
    /// of minutes, and nS is the number of seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies one
    /// month, four days, two hours, and five minutes).
    pub delete_expired_task_after: Option<String>,
    /// Gets or sets a Boolean value that indicates that the task will not be started if the computer is running on
    /// battery power.
    pub disallow_start_if_on_batteries: Option<bool>,
    /// Gets or sets a Boolean value that indicates that the task is enabled. The task can be performed only when this
    /// setting is True.
    pub enabled: Option<bool>,
    /// Gets or sets the amount of time allowed to complete the task.
    pub execution_time_limit: Option<String>,
    /// Gets or sets a Boolean value that indicates that the task will not be visible in the UI. However, administrators
    /// can override this setting through the use of a "master switch" that makes all tasks visible in the UI.
    pub hidden: Option<bool>,
    /// Gets or sets the information that specifies how the Task Scheduler performs tasks when the computer is in an idle state.
    pub idle_settings: Option<IdleSettings>,
    /// Gets or sets the policy that defines how the Task Scheduler deals with multiple instances of the task.
    pub multiple_instances_policy: Option<InstancesPolicy>,
    /// The network settings object that contains a network profile identifier and name.
    /// If the `run_only_if_network_available` property is true and
    /// a network profile is specified in the [network_settings](Settings::network_settings)
    /// field, then the task will run only if the specified network profile is available.
    pub network_settings: Option<NetworkSettings>,
    /// Gets or sets the priority level of the task.
    pub priority: Option<i32>,
    /// Gets or sets the number of times that the Task Scheduler will attempt to restart the task.
    pub restart_count: Option<i32>,
    /// Gets or sets a value that specifies how long the Task Scheduler will attempt to restart the task.
    pub restart_interval: Option<String>,
    /// Gets or sets a Boolean value that indicates that the Task Scheduler will run the task only if the
    /// computer is in an idle state.
    pub run_only_if_idle: Option<bool>,
    /// Gets or sets a Boolean value that indicates that the Task Scheduler will run the task only when a
    /// network is available.
    pub run_only_if_network_available: Option<bool>,
    /// Gets or sets a Boolean value that indicates that the Task Scheduler can start the task at any time
    /// after its scheduled time has passed.
    pub start_when_available: Option<bool>,
    /// Gets or sets a Boolean value that indicates that the task will be stopped if the computer begins to
    /// run on battery power.
    pub stop_if_going_on_batteries: Option<bool>,
    /// Gets or sets a Boolean value that indicates that the Task Scheduler will wake the computer when it is
    /// time to run the task.
    pub wake_to_run: Option<bool>,
    /// Gets or sets an XML-formatted definition of the task settings.
    pub xml_text: Option<String>,
}

impl Settings {
    /// Creates a new Settings struct with all values set to None.
    pub fn new() -> Settings {
        Settings {
            allow_demand_start: None,
            allow_hard_terminate: None,
            compatibility: None,
            delete_expired_task_after: None,
            disallow_start_if_on_batteries: None,
            enabled: None,
            execution_time_limit: None,
            hidden: None,
            idle_settings: None,
            multiple_instances_policy: None,
            network_settings: None,
            priority: None,
            restart_count: None,
            restart_interval: None,
            run_only_if_idle: None,
            run_only_if_network_available: None,
            start_when_available: None,
            stop_if_going_on_batteries: None,
            wake_to_run: None,
            xml_text: None,
        }
    }
}

/// Represents a duration of time.
#[derive(Debug, Clone, Copy)]
#[allow(missing_docs)]
pub struct Duration {
    pub days: Option<usize>,
    pub hours: Option<usize>,
    pub minutes: Option<usize>,
    pub months: Option<usize>,
    pub seconds: Option<usize>,
    pub years: Option<usize>,
}

impl Duration {
    /// Creates a new instance of Duration with all values set to None.
    pub fn new() -> Duration {
        Duration {
            days: None,
            hours: None,
            minutes: None,
            months: None,
            seconds: None,
            years: None,
        }
    }
}

macro_rules! format_duration {
    ($str:expr, $num:expr, $code:literal) => {
        if let Some(num) = $num {
            if num > 0 {
                $str = format!("{}{}{}", $str, num, $code);
            }
        }
    };
}

impl fmt::Display for Duration {
    /// Formats a duration to a string similar to the duration of the ISO 8601 spec.
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let mut s = String::new();

        // The order of the string is very important. This is a requirement of the Task Scheduler
        // API
        format_duration!(s, self.years, "Y");
        format_duration!(s, self.months, "M");
        format_duration!(s, self.days, "D");

        if self.hours.is_some() || self.minutes.is_some() || self.hours.is_some() {
            s = format!("{}T", s);
            format_duration!(s, self.hours, "H");
            format_duration!(s, self.minutes, "M");
            format_duration!(s, self.seconds, "S");
        }

        write!(f, "P{}", s)
    }
}

impl Default for Duration {
    fn default() -> Self {
        Duration::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn macro_test() {
        let mut s = String::new();
        format_duration!(s, Some(2), "Y");
        format_duration!(s, Some(1), "D");
        assert_eq!(s, "2Y1D");
    }

    #[test]
    fn duration_all_set() {
        let d = Duration {
            years: Some(1),
            months: Some(2),
            days: Some(3),
            hours: Some(4),
            minutes: Some(5),
            seconds: Some(6),
        };

        assert_eq!("P1Y2M3DT4H5M6S", d.to_string());
    }

    #[test]
    fn duration_time_only() {
        let mut d = Duration::new();
        d.hours = Some(1);
        d.minutes = Some(2);
        d.seconds = Some(3);

        assert_eq!("PT1H2M3S", d.to_string());
    }
    
    #[test]
    fn duration_zero_year_removed() {
        let mut d = Duration::new();
        d.years = Some(0);

        assert_eq!("", d.to_string());
    }
}
