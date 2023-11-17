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
```

For more examples, refer to the `planif/examples` folder. The folder contains code for creating each of the triggers.


## Trigger settings
All settings are available for the tasks.

The documentation contains all relevant information from the
[Microsoft Task Scheduler documentation](https://learn.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-reference).

## Running examples
After cloning the repository, examples can be run using either 

`cargo run --example <name>` to run a specific example or,
`cargo run --examples` to run all examples.

## Changelog
See the [changelog file](CHANGELOG.md).

## Upgrade Guide (0.* to 1.0)

### ComRuntime

#### Creating
The `ComRuntime` is now handled by the `TaskScheduler` and should be created using: 
```rust
    let ts = TaskScheduler::new()?;
    let com = ts.get_com();
    let sb = ScheduleBuilder::new(&com).unwrap();
    // ... snip
```

#### Uninilizing
`ScheduleBuilder`'s no longer need to be manually uninitialized. In pre-1.0, `ScheduleBuilding::uninitalize()` would call `CoUninitialize` which would effectively close the COM. This could be problamatic if you had multiple builders or schedules built since having n+1 COMs would simply reuse the initial COM.

The changes brought in 1.0 will now keep an `Rc<Com>` privately stored in the `TaskScheduler`. Now when all the references are dropped, the COM uninitialization will be handled "automagically".  
