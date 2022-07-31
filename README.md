`planif` is a builder pattern wrapper around the windows task scheduler API ([windows-rs](https://github.com/microsoft/windows-rs)).

## How to
Please refer to the `examples` folder. The folder contains code for creating each of the triggers.

## TODO
### Create Triggers
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

### Trigger settings
- [ ] IdleSettings 

Other settings may also be missing.

### Other (maybe)
- [ ] Delete triggers
- [ ] List triggers

## A builder library for the Windows Task Scheduler.

Inspired by the great work of [j-hc](https://github.com/j-hc/windows-taskscheduler-api-rust) which allowed my to figure out how to grok the windows-rs library.
