use planif::enums::TaskCreationFlags;
use planif::schedule_builder::{Action, ScheduleBuilder};
use planif::schedule::TaskScheduler;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let ts = TaskScheduler::new()?;
    let com = ts.get_com();
    let sb = ScheduleBuilder::new(&com).unwrap();

    sb.create_boot()
        .author("Matt")?
        .description("Test Time Trigger")?
        .trigger("test_time_trigger", true)?
        .action(Action::new("test_time_action", "notepad.exe", "", ""))?
        .start_boundary("2022-04-28T02:14:08.660633427+00:00")?
        .build()?
        .register("TimeTaskName", TaskCreationFlags::CreateOrUpdate as i32)?;
    Ok(())
}
