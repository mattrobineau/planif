use chrono::prelude::*;
use planif::enums::{DayOfWeek, Month, TaskCreationFlags, WeekOfMonth};
use planif::schedule_builder::{Action, ComRuntime, MonthlyDOW, ScheduleBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let com = ComRuntime::new()?;
    let builder: ScheduleBuilder<MonthlyDOW> = ScheduleBuilder::new(&com)?
        .create_monthly_dow()
        .author("Matt")?
        .description("Test Trigger")?
        .trigger("MyTrigger", true)?
        .action(Action::new("test", "notepad.exe", "", ""))?
        .months_of_year(vec![Month::January])?
        .days_of_week(vec![DayOfWeek::Monday])?
        .weeks_of_month(vec![WeekOfMonth::First])?
        .start_boundary(&Local::now().to_rfc3339())?;

    builder.build()?.register(
        "MonthlyDOWTaskName",
        TaskCreationFlags::CreateOrUpdate as i32,
    )?;

    Ok(())
}
