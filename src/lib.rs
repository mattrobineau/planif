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

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() -> Result<(), Box<dyn std::error::Error>> {
        use crate::schedule_builder::{Action, ScheduleBuilder};
        use crate::enums::TaskCreationFlags;
        use chrono::prelude::*;

        ScheduleBuilder::new()?
            .create_daily()
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
}
