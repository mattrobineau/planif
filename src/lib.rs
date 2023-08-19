#![deny(missing_docs)]

//! A builder pattern wrapper around the [windows-rs](https://github.com/microsoft/windows-rs) task scheduler API .
//!
//! Provides an ergonomic builder for creating the following task types:
//! - Boot
//! - Daily
//! - Event
//! - Idle
//! - Logon
//! - Monthly
//! - MonthlyDOW
//! - Registration
//! - Time
//! - Weekly
//!
//! ## Example
//!
//! ```rust
//! use chrono::prelude::*;
//! use planif::schedule::TaskCreationFlags;
//! use planif::schedule_builder::{Action, ScheduleBuilder};
//!
//! fn main() -> Result<(), Box<dyn std::error::Error>> {
//!     let sb = ScheduleBuilder::new().unwrap();
//!     sb.create_daily()
//!         .author("planif")?
//!         .description("Daily Trigger")?
//!         .trigger("daily_trigger", true)?
//!         .days_interval(1)?
//!         .action(Action::new("test", "notepad.exe", "", ""))?
//!         .start_boundary(&Local::now().to_rfc3339())?
//!         .build()?
//!         .register("TaskName", TaskCreationFlags::CreateOrUpdate as i32)?;
//!     Ok(())
//! }
//! ```
//!
//! For more examples, refer to the `planif/examples` folder. The folder contains code for creating each of the triggers.

/// Enums used throughout the crate.
pub mod enums;
/// Errors used throughout the crate.
pub mod error;
/// Register scheduled tasks.
pub mod schedule;
/// Build different [Schedules](schedule::Schedule) for the Windows Task Scheduler.
pub mod schedule_builder;
/// Various settings available while building [Schedules](schedule::Schedule).
pub mod settings;
/// The base struct used to manage all the others.
pub mod task_scheduler;
/// Com
pub mod com;

#[cfg(test)]
mod tests {
    use chrono::{Duration, Local};

    use crate::{task_scheduler::{TaskScheduler}, enums::TaskCreationFlags, schedule_builder::Action};

    #[test]
    fn it_works() -> Result<(), Box<dyn std::error::Error>> {
        let ts = TaskScheduler::new();
        ts
            .create_schedule()
            .create_time()
            .author("Test dummy")?
            .description("Test Time Trigger")?
            .in_folder("\\Test folder")?
            .trigger("test_time_trigger", true)?
            .action(Action::new("test_time_action", "notepad.exe", "", ""))?
            .start_boundary(
                &Local::now()
                    .checked_add_signed(Duration::seconds(10))
                    .unwrap()
                    .to_rfc3339(),
            )?
            .build()?
            .register("TimeTaskName", TaskCreationFlags::CreateOrUpdate as i32)?;

        println!("{}", ts.get_schedule("\\Test folder\\TimeTaskName").path());

        // ts.get_schedules_from_folder("Mozilla")
        //     .unwrap()
        //     .into_iter()
        //     .for_each(|x| println!("{}", x.path()));

        Ok(())
    }
}
