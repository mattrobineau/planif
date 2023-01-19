#[derive(Debug, Clone, Copy)]
/// Represents the day of the month.
pub enum DayOfMonth {
    /// Day of the month, between 1 and 31 inclusive.
    Day(i32),
    /// Represents the last day of the month.
    Last,
}

impl From<DayOfMonth> for i32 {
    /// Shifts the DayOfMonth::Day(x) to the decimal value
    ///
    /// See Remarks table: <https://docs.microsoft.com/en-us/windows/win32/taskschd/monthlytrigger-daysofmonth>
    #[allow(overflowing_literals)]
    fn from(day: DayOfMonth) -> Self {
        match day {
            DayOfMonth::Day(i) => 1 << (i - 1),
            DayOfMonth::Last => 0x80000000,
        }
    }
}

#[derive(Debug, Clone, Copy)]
/// Day of the week.
pub enum DayOfWeek {
    /// Sunday (0x01)
    Sunday = 1 << 0,
    /// Monday (0x02)
    Monday = 1 << 1,
    /// Tuesday (0x04)
    Tuesday = 1 << 2,
    /// Wednesday (0x08)
    Wednesday = 1 << 3,
    /// Thursday (0x10)
    Thursday = 1 << 4,
    /// Friday (0x20)
    Friday = 1 << 5,
    /// Saturday (0x40)
    Saturday = 1 << 6,
}

#[derive(Debug, Clone, Copy)]
/// Month of the year.
pub enum Month {
    /// January (0x01)
    January = 1 << 0,
    /// Frebruary (0x02)
    February = 1 << 1,
    /// March (0x04)
    March = 1 << 2,
    /// April (0x08)
    April = 1 << 3,
    /// May (0x10)
    May = 1 << 4,
    /// June (0x20)
    June = 1 << 5,
    /// July (0x40)
    July = 1 << 6,
    /// August (0x80)
    August = 1 << 7,
    /// September (0x100)
    September = 1 << 8,
    /// October (0x200)
    October = 1 << 9,
    /// November (0x400)
    November = 1 << 10,
    /// December (0x800)
    December = 1 << 11,
}

impl From<Month> for i16 {
    fn from(month: Month) -> Self {
        month as i16
    }
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
    /// for the context principal. When the [register method](crate::schedule::Schedule::register) function is called with this flag to
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
    /// register the task. This constant cannot be combined with the [Create](TaskCreationFlags::Create),
    /// [Update](TaskCreationFlags::Update), or [CreateOrUpdate](TaskCreationFlags::CreateOrUpdate) values.
    ValidateOnly = 1,
}

#[derive(Debug, Clone, Copy)]
/// The week of the month
pub enum WeekOfMonth {
    /// First (0x01)
    First = 1 << 0,
    /// Second (0x02)
    Second = 1 << 1,
    /// Third (0x04)
    Third = 1 << 2,
    /// Fourth (0x08)
    Fourth = 1 << 3,
}
