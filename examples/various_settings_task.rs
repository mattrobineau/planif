use chrono::prelude::*;
use planif::enums::TaskCreationFlags;
use planif::schedule_builder::{Action, ComRuntime, ScheduleBuilder};
use planif::settings::Settings;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let com = ComRuntime::new()?;
    let sb = ScheduleBuilder::new(&com).unwrap();
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
