// use chrono::prelude::*;
// use planif::enums::TaskCreationFlags;
// use planif::schedule_builder::{Action, ComRuntime, ScheduleBuilder};

// fn main() -> Result<(), Box<dyn std::error::Error>> {
//     let com = ComRuntime::new()?;

//     let sb = ScheduleBuilder::new(&com).unwrap();
//     sb.create_daily()
//         .author("Matt")?
//         .description("Test Trigger")?
//         .trigger("test_trigger", true)?
//         .days_interval(1)?
//         .action(Action::new("test", "notepad.exe", "", ""))?
//         .start_boundary(&Local::now().to_rfc3339())?
//         .build()?
//         .register("TaskName", TaskCreationFlags::CreateOrUpdate as i32)?;

//     Ok(())
// }
