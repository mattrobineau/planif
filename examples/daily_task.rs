use chrono::prelude::*;
use planif::enums::TaskCreationFlags;
use planif::schedule_builder::{Action, ScheduleBuilder};
use planif::schedule::TaskScheduler;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let ts = TaskScheduler::new()?;
    let com = ts.get_com();
    let sb = ScheduleBuilder::new(&com).unwrap();

    sb.create_daily()
        .author("Matt")?
        .description("Test Trigger")?
        .trigger("test_trigger", true)?
        .days_interval(1)?
        .action(Action::new("test", "notepad.exe", "", ""))?
        .start_boundary(&Local::now().to_rfc3339())?
        .build()?
        .register("TaskName", TaskCreationFlags::CreateOrUpdate as i32)?;

    Ok(())
}
