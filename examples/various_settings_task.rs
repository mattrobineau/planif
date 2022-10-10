use chrono::prelude::*;
use planif::schedule::TaskCreationFlags;
use planif::schedule_builder::{Action, ScheduleBuilder};
use planif::settings::Settings;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let sb = ScheduleBuilder::new().unwrap();
    let mut settings = Settings::new();
    settings.run_only_if_idle = Some(true);

    sb.create_daily()
        .author("Matt")?
        .description("Test Trigger")?
        .settings(settings)?
        .trigger("test_trigger", true)?
        .days_interval(1)?
        .action(Action::new("test", "notepad.exe", "", ""))?
        .start_boundary(&Local::now().to_rfc3339())?
        .build()?
        .register("TaskName", TaskCreationFlags::CreateOrUpdate as i32)?;
    Ok(())
}
