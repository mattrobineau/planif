use chrono::prelude::*;
use planif::schedule::TaskCreationFlags;
use planif::schedule_builder::{Action, ScheduleBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let sb = ScheduleBuilder::new().unwrap();
    sb.create_weekly()
        .author("Matt")?
        .description("Test Weekly Trigger")?
        .trigger("test_weekly_trigger", 1)?
        .action(Action::new("test", "notepad.exe", "", ""))?
        .start_boundary(&Local::now().to_rfc3339())?
        .days_of_week(16)?
        .weeks_interval(3)?
        .build()?
        .register("WeeklyTaskName", TaskCreationFlags::CreateOrUpdate as i32)?;
    Ok(())
}
