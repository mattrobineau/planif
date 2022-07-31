# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
* Assign value queries to trigger

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
