use chrono::prelude::*;
use planif::enums::{DayOfMonth, Month, TaskCreationFlags};
use planif::schedule_builder::{Action, Monthly, ScheduleBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let builder: ScheduleBuilder<Monthly> = ScheduleBuilder::new()?
        .create_monthly()
        .author("Matt")?
        .description("Test Trigger")?
        .trigger("MyTrigger", true)?
        .action(Action::new("test", "notepad.exe", "", ""))?
        .days_of_month(vec![DayOfMonth::Day(1), DayOfMonth::Day(15)])?
        .months_of_year(vec![Month::December])?
        .start_boundary(&Local::now().to_rfc3339())?;

    builder
        .build()?
        .register("MonthlyTaskName", TaskCreationFlags::CreateOrUpdate as i32)?;

    Ok(())
}
