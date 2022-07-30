pub enum DayOfMonth {
    Day(i32),
    Last,
}

impl From<DayOfMonth> for i32 {
    fn from(day: DayOfMonth) -> Self {
        match day {
            DayOfMonth::Day(i) => 1 << i,
            DayOfMonth::Last => 0x80000000,
        }
    }
}

pub enum DayOfWeek {
    Sunday = 1 << 1,
    Monday = 1 << 2,
    Tuesday = 1 << 3,
    Wednesday = 1 << 4,
    Thursday = 1 << 5,
    Friday = 1 << 6,
    Saturday = 1 << 7,
}

pub enum Month {
    January = 1 << 1,
    February = 1 << 2,
    March = 1 << 3,
    April = 1 << 4,
    May = 1 << 5,
    June = 1 << 6,
    July = 1 << 7,
    August = 1 << 8,
    September = 1 << 9,
    October = 1 << 10,
    November = 1 << 11,
    December = 1 << 12,
}

impl From<Month> for i16 {
    fn from(month: Month) -> Self {
        month as i16
    }
}

pub enum WeekOfMonth {
    First = 1 << 1,
    Second = 1 << 2,
    Third = 1 << 3,
    Fourth = 1 << 4,
}
