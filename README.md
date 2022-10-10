`planif` is a builder pattern wrapper around the windows task scheduler API ([windows-rs](https://github.com/microsoft/windows-rs)).

## Functionality

The `planif` crate provides an ergonomic builder over top of the Win32 Task Scheduler API.

The builder supports the following trigger types:
- [x] Boot 
- [x] Daily 
- [X] Event 
- [X] Idle 
- [x] Logon 
- [X] MonthlyDOW 
- [x] Monthly 
- [X] Registration 
- [x] Time 
- [x] Weekly 

## Usage
Add this to your `Cargo.toml` file:
```
[dependencies]
planif = "0.2"
```

## Example

```rust
use chrono::prelude::*;
use planif::schedule::TaskCreationFlags;
use planif::schedule_builder::{Action, ScheduleBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let sb = ScheduleBuilder::new().unwrap();
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
```

For more examples, refer to the `planif/examples` folder. The folder contains code for creating each of the triggers.


## Trigger settings
All settings are available for the tasks.

The documentation contains all relevant information from the
[Microsoft Task Scheduler documentation](https://learn.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-reference).

## Changelog
See the [changelog file](CHANGELOG.md).
