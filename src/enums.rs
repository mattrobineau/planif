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
            DayOfMonth::Day(i) => 1 << i - 1,
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
