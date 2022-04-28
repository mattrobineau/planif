use chrono::prelude::*;
use planif::schedule::TaskCreationFlags;
use planif::schedule_builder::{Action, ScheduleBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let sb = ScheduleBuilder::new().unwrap();
    sb.create_time()
        .author("Matt")?
        .description("Test Time Trigger")?
        .trigger("test_time_trigger", 1)?
        .action(Action::new("test_time_action", "notepad.exe", "", ""))?
        .start_boundary("2022-04-28T02:14:08.660633427+00:00")?
        // RandomDelay of 2 seconds
        .random_delay("PT2S")?
        .build()?
        .register("TimeTaskName", TaskCreationFlags::CreateOrUpdate as i32)?;
    Ok(())
}
