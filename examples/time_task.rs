use planif::enums::TaskCreationFlags;
use planif::schedule_builder::{Action, ScheduleBuilder};
use planif::schedule::TaskScheduler;

use planif::settings::Duration;
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let ts = TaskScheduler::new()?;
    let com = ts.get_com();

    let sb = ScheduleBuilder::new(&com).unwrap();
    sb.create_time()
        .author("Matt")?
        .description("Test Time Trigger")?
        .trigger("test_time_trigger", true)?
        .action(Action::new("test_time_action", "notepad.exe", "", ""))?
        .start_boundary("2022-04-28T02:14:08.660633427+00:00")?
        // RandomDelay of 2 seconds
        .random_delay(Duration {
            seconds: Some(5),
            days: Some(2),
            ..Default::default()
        })?
        .build()?
        .register("TimeTaskName", TaskCreationFlags::CreateOrUpdate as i32)?;
    Ok(())
}
