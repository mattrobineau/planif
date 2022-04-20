use chrono::prelude::*;
use planif::schedule::{Action, ScheduleBuilder, TaskCreationFlags};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let sb = ScheduleBuilder::new().unwrap();
    sb.create_daily()
        .author("Matt")?
        .description("Test Trigger")?
        .trigger("test_trigger", 1)?
        .days_interval(1)?
        .action(Action::new("test", "notepad.exe", "", ""))?
        .start_boundary(&Local::now().to_rfc3339())?
        .build()?
        .register("TaskName", TaskCreationFlags::CreateOrUpdate as i32)?;
    Ok(())
}
