use chrono::{Duration, Local};
use planif::enums::TaskCreationFlags;
use planif::schedule_builder::{Action, ScheduleBuilder};
use planif::schedule::TaskScheduler;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let ts = TaskScheduler::new()?;
    let com = ts.get_com();
    let sb = ScheduleBuilder::new(&com).unwrap();

    sb.create_time()
        .author("Matt")?
        .description("Test Time Trigger in folder")?
        .in_folder("Test folder")?
        .trigger("test_time_folder_trigger", true)?
        .action(Action::new("test_time_action", "notepad.exe", "", ""))?
        .start_boundary(
            &Local::now()
                .checked_add_signed(Duration::seconds(5))
                .unwrap()
                .to_rfc3339(),
        )?
        .build()?
        .register("Time Folder Task", TaskCreationFlags::CreateOrUpdate as i32)?;

    Ok(())
}
