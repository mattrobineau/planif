/// Values for the security logon method.
///
/// # None
/// The logon method is not specified. Used for non-NT credentials.
///
/// # Password
/// Use a password for logging on the user. The password must be supplied at registration time.
///
/// # S4U
/// Use an existing interactive token to run a task. The user must log on using a service for user (S4U) logon.
/// When an S4U logon is used, no password is stored by the system and there is no access to either the network
/// or encrypted files.
///
/// # InteractiveToken
/// User must already be logged on. The task will be run only in an existing interactive session.
///
/// # Group
/// Group activation. The user_id field specifies the group.
///
/// # ServiceAccount
/// Indicates that a Local System, Local Service, or Network Service account is being used as a security context
/// to run the task.
///
/// # InteractiveTokenOrPassword
/// First use the interactive token. If the user is not logged on (no interactive token is available), then the password
/// is used. The password must be specified when a task is registered. This flag is not recommended for new tasks because
/// it is less reliable than LogonType::Password.
///
pub enum LogonType {
    None = 0,
    Password,
    S4U,
    InteractiveToken,
    Group,
    ServiceAccount,
    InteractiveTokenOrPassword,
}

/// Use to set the settings for the principal
/// # Properties
///
/// ## display_name
/// Gets or sets the name of the principal that is displayed in the Task Scheduler UI.
///
/// ## group_id
/// Gets or sets the identifier of the user group that is required to run the tasks that are associated with the principal.
/// Do not set this property if a user identifier is specified in the user_id property.
///
/// ## id
/// Gets or sets the identifier of the principal.
///
/// ## logon_type
/// Gets or sets the security logon method that is required to run the tasks that are associated with the principal.
/// This property is valid only when a user identifier is specified by the UserId property.
///
/// ## run_level
/// Gets or sets the identifier that is used to specify the privilege level that is required to run the tasks
/// that are associated with the principal.
///
/// ## user_id
/// Gets or sets the user identifier that is required to run the tasks that are associated with the principal.
/// Do not set this property if a group identifier is specified in the group_id property.
///
/// # Reference
/// https://docs.microsoft.com/en-us/windows/win32/taskschd/principal
pub struct PrincipalSettings {
    pub display_name: String,
    pub group_id: Option<String>,
    pub id: String,
    pub logon_type: LogonType,
    pub run_level: RunLevel,
    pub user_id: Option<String>,
}

/// Values for the identifier that is used to specify the privilege level that is required to run the tasks
/// that are associated with the principal.
/// # Highest
/// Tasks will be run with the highest privileges.
///
/// # LUA
/// Tasks will be run with the least privileges (LUA).
pub enum RunLevel {
    Highest = 1,
    LUA = 0,
}

/// Defines all available settings on a task.
///
/// # Example
/// ```
/// // All values are set to `None` when using `new()`
/// let settings = settings::new();
/// settings.allow_demand_start = Some(true);
/// ```
///
/// # Description
/// ## allow_demand_start
/// Gets or sets a Boolean value that indicates that the task can be started by using either the Run command
/// or the Context menu.
///
/// ## allow_hard_terminate
/// Gets or sets a Boolean value that indicates that the task may be terminated by using TerminateProcess.
///
/// ## compatibility
/// Gets or sets an integer value that indicates which version of Task Scheduler a task is compatible with.
///
/// ## delete_expired_task_after
/// Gets or sets the amount of time that the Task Scheduler will wait before deleting the task after it expires.
///
/// A string that gets or sets the amount of time that the Task Scheduler will wait before deleting the task after
/// it expires. The format for this string is PnYnMnDTnHnMnS, where nY is the number of years, nM is the number of
/// months, nD is the number of days, 'T' is the date/time separator, nH is the number of hours, nM is the number
/// of minutes, and nS is the number of seconds (for example, PT5M specifies 5 minutes and P1M4DT2H5M specifies one
/// month, four days, two hours, and five minutes).
///
/// ## disallow_start_if_on_batteries
/// Gets or sets a Boolean value that indicates that the task will not be started if the computer is running on
/// battery power.
///
/// ## enabled
/// Gets or sets a Boolean value that indicates that the task is enabled. The task can be performed only when this
/// setting is True.
///
/// ## execution_time_limit
/// Gets or sets the amount of time allowed to complete the task.
///
/// ## hidden
/// Gets or sets a Boolean value that indicates that the task will not be visible in the UI. However, administrators
/// can override this setting through the use of a "master switch" that makes all tasks visible in the UI.
///
/// ## restart_on_idle
/// Gets or sets a Boolean value that indicates whether the task is restarted when the computer cycles into an idle
/// condition more than once.
///
/// ## multiple_instances_policy
/// Gets or sets the policy that defines how the Task Scheduler deals with multiple instances of the task.
///
/// ## network_id
/// Gets or sets a GUID value that identifies a network profile.
///
/// ## network_name
/// Gets or sets the name of a network profile. The name is used for display purposes.
///
/// ## priority
/// Gets or sets the priority level of the task.
///
/// ## restart_count
/// Gets or sets the number of times that the Task Scheduler will attempt to restart the task.
///
/// ## restart_interval
/// Gets or sets a value that specifies how long the Task Scheduler will attempt to restart the task.
///
/// ## run_only_if_idle
/// Gets or sets a Boolean value that indicates that the Task Scheduler will run the task only if the
/// computer is in an idle state.
///
/// ## run_only_if_network_available
/// Gets or sets a Boolean value that indicates that the Task Scheduler will run the task only when a
/// network is available.
///
/// ## start_when_available
/// Gets or sets a Boolean value that indicates that the Task Scheduler can start the task at any time
/// after its scheduled time has passed.
///
/// ## stop_if_going_on_batteries
/// Gets or sets a Boolean value that indicates that the task will be stopped if the computer begins to
/// run on battery power.
///
/// ## stop_on_idle_end
/// Gets or sets a Boolean value that indicates that the Task Scheduler will terminate the task if the
/// idle condition ends before the task is completed.
///
/// ## wake_to_run
/// Gets or sets a Boolean value that indicates that the Task Scheduler will wake the computer when it is
/// time to run the task.
///
/// ## xml_text
/// Gets or sets an XML-formatted definition of the task settings.
///
/// # References
/// https://docs.microsoft.com/en-us/windows/win32/taskschd/tasksettings
/// https://docs.microsoft.com/en-us/windows/win32/taskschd/tasksettings-priority
/// https://docs.microsoft.com/en-us/windows/win32/procthread/scheduling-priorities
/// https://docs.microsoft.com/en-us/windows/win32/taskschd/networksettings
/// https://docs.microsoft.com/en-us/windows/win32/taskschd/idlesettings

pub struct Settings {
    pub allow_demand_start: Option<bool>,
    pub allow_hard_terminate: Option<bool>,
    pub compatibility: Option<Compatibility>,
    pub delete_expired_task_after: Option<String>,
    pub disallow_start_if_on_batteries: Option<bool>,
    pub enabled: Option<bool>,
    pub execution_time_limit: Option<String>,
    pub hidden: Option<bool>,
    pub restart_on_idle: Option<bool>,
    pub multiple_instances_policy: Option<InstancesPolicy>,
    pub network_id: Option<String>,
    pub network_name: Option<String>,
    pub priority: Option<i32>,
    pub restart_count: Option<i32>,
    pub restart_interval: Option<String>,
    pub run_only_if_idle: Option<bool>,
    pub run_only_if_network_available: Option<bool>,
    pub start_when_available: Option<bool>,
    pub stop_if_going_on_batteries: Option<bool>,
    pub stop_on_idle_end: Option<bool>,
    pub wake_to_run: Option<bool>,
    pub xml_text: Option<String>,
}

impl Settings {
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
            restart_on_idle: None,
            stop_on_idle_end: None,
            multiple_instances_policy: None,
            network_id: None,
            network_name: None,
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

/// Values for task compatibility
/// Task compatibility, which is set through the Compatibility property, should only be set to TASK_COMPATIBILITY_V1
/// if a task needs to be accessed or modified from a Windows XP, Windows Server 2003, or Windows 2000 computer.
/// Otherwise, it is recommended that Task Scheduler 2.0 compatibility be used because the task will have more features.
/// Tasks compatible with the AT command can only have one time trigger.
/// Tasks compatible with Task Scheduler 1.0 can only have a time trigger, a logon trigger, or a boot trigger, and the
/// task can only have an executable action.
/// see https://docs.microsoft.com/en-us/windows/win32/taskschd/tasksettings-compatibility
pub enum Compatibility {
    AT = 0,
    V1,
    V2,
}

/// Values for the instance policy.
pub enum InstancesPolicy {
    Parallel = 0,
    Queue,
    IgnoreNew,
    StopExisting,
}
