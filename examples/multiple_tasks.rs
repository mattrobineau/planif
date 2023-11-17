use chrono::prelude::*;
use planif::enums::TaskCreationFlags;
use planif::schedule::TaskScheduler;
use planif::schedule_builder::{Action, ScheduleBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let ts = TaskScheduler::new()?;
    let com = ts.get_com();

    if true {
        let sb = ScheduleBuilder::new(&com).unwrap();
        sb.create_daily()
            .author("Matt")?
            .description("Test Trigger 1")?
            .trigger("test_trigger_1", true)?
            .days_interval(1)?
            .action(Action::new("test", "notepad.exe", "", ""))?
            .start_boundary(&Local::now().to_rfc3339())?
            .build()?
            .register("TaskName1", TaskCreationFlags::CreateOrUpdate as i32)?;
    }

    if true {
        let sb2 = ScheduleBuilder::new(&com).unwrap();
        sb2.create_daily()
            .author("Matt")?
            .description("Test Trigger 2")?
            .trigger("test_trigger_2", true)?
            .days_interval(1)?
            .action(Action::new("test", "notepad.exe", "", ""))?
            .start_boundary(&Local::now().to_rfc3339())?
            .build()?
            .register("TaskName2", TaskCreationFlags::CreateOrUpdate as i32)?;
    }
    Ok(())
}
