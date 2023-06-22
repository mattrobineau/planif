# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
* Add `Duration` struct to ease usage of delays and other methods using string duration formats
* Add Default impl for `Duration`
* Add tests for `Duration` to_string() impl
* Add example for multiple builders reusing a com runtime
* Change namespaces for `TaskCreationFlags`
* Bump version to 1.0
    * Bump windows dependency to 0.43.0 and fix related breaking changes
* Add tests for `Duration` to_string() impl
* Add ability to specify a folder for scheduled tasks
* Update examples for in_folder()

### Breaking changes
* Change `delay` methods to take durations instead of `str`
* Change `schedule_builder`s to take in a ComRuntime when creating them.
    * ComRuntime allows users to create multiple builders while reusing the com and stops users from accidentally
      dropping the com

## [0.2.1]
### Updated
* Update documentation, add deny missing_docs

## [0.2.0] - 2022-10-09

### Added
* Add builder function for a task's settings
* Add `NetworkSettings` (removed from `Settings`)
* Add `IdleSettings` (also removed from `Settings`)

## [0.1.1] - 2022-08-04

### Added
* Examples for monthly and monthly dow (day of week)
* Added Error when not setting a start_boundary for calendar events

### Fixed
* Assign value queries to trigger
* Fix documentation examples
* Fix DaysOfMonth values
* Fix Month values
* Fix monthly and monthly dow triggers

### Known issues
- Setting the DaysOfMonth to `DaysOfMonth::Last` causes an i32 overflow. The windows-rs call to 
[SetDaysOfMonth](https://microsoft.github.io/windows-docs-rs/doc/windows/Win32/System/TaskScheduler/struct.IMonthlyTrigger.html#method.SetDaysOfMonth)
takes an i32 but also expects the `Last` value to be `0x80000000`. For the time being, the library ignores
the overflow error with `#[allow(overflowing_literals)]`.

## [0.1.0] - 2022-07-30

This release completes the available triggers in the Windows Task Scheduler.

### Added
* Add monthly dow trigger
* Add monthly trigger
* Add Changelog
* Add event trigger
* Add registration trigger

### Changed
* Improve documentation
* Change triggers to accept bools instead of i16
* Update cargo toml
* Update readme

### Fixed
* Fix casting issues
* Removed mutability to parameters that do not need it

## [0.0.1] - 2022-04-10
### Added
* Add ability to set settings on task
* Add ability to set principal
* Add repetition (untested)
* Add create logon task
* Add weekly task example
* Add time trigger example
* Create LICENSE
* Create builder/schedule for daily triggers
