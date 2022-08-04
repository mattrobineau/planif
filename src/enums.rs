#[derive(Debug, Clone, Copy)]
pub enum DayOfMonth {
    Day(i32),
    Last,
}

impl From<DayOfMonth> for i32 {
    /// Shifts the DayOfMonth::Day(x) to the decimal value
    ///
    /// See Remarks table: <https://docs.microsoft.com/en-us/windows/win32/taskschd/monthlytrigger-daysofmonth>
    #[allow(overflowing_literals)]
    fn from(day: DayOfMonth) -> Self {
        match day {
            DayOfMonth::Day(i) => 1 << i - 1,
            DayOfMonth::Last => 0x80000000,
        }
    }
}

#[derive(Debug, Clone, Copy)]
pub enum DayOfWeek {
    Sunday = 1 << 0,
    Monday = 1 << 1,
    Tuesday = 1 << 2,
    Wednesday = 1 << 3,
    Thursday = 1 << 4,
    Friday = 1 << 5,
    Saturday = 1 << 6,
}

#[derive(Debug, Clone, Copy)]
pub enum Month {
    January = 1 << 0,
    February = 1 << 1,
    March = 1 << 2,
    April = 1 << 3,
    May = 1 << 4,
    June = 1 << 5,
    July = 1 << 6,
    August = 1 << 7,
    September = 1 << 8,
    October = 1 << 9,
    November = 1 << 10,
    December = 1 << 11,
}

impl From<Month> for i16 {
    fn from(month: Month) -> Self {
        month as i16
    }
}

#[derive(Debug, Clone, Copy)]
pub enum WeekOfMonth {
    First = 1 << 0,
    Second = 1 << 1,
    Third = 1 << 2,
    Fourth = 1 << 3,
}
