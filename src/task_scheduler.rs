use std::mem::ManuallyDrop;
use std::rc::Rc;

use windows::core::BSTR;
use windows::Win32::System::Com::*;
use windows::Win32::System::TaskScheduler::*;

use crate::schedule::Registered;
use crate::schedule::Schedule;
use crate::schedule_builder::Base;
use crate::schedule_builder::ScheduleBuilder;

// new plan
// TaskScheduler struct which represents...the service?
// can Get() a Schedule - an individual task
// the Schedule can set the author and all the other things
// can Save() a Schedule, persists the changes to the task scheduler and consumes it

// TaskScheduler can call Create() which returns a Schedule
// can set on all that
// can register a schedule builder

// remove schedulebuilder
// make build just "finish" everything, validation
// can change the marker type
// and register takes a Final schedule

// when you get a schedule its mostly valid
// it

/// Represents a COM runtime required for building [`TaskScheduler`s](task_scheduler::TaskScheduler)
#[derive(Debug, PartialEq)]
pub struct ComRuntime(Rc<Com>);

impl ComRuntime {
    /// Creates a COM runtime for use with one or more
    /// [ScheduleBuilder]
    pub fn new() -> Result<Self, Box<dyn std::error::Error>> {
        Ok(ComRuntime(Rc::new(Com::initialize()?)))
    }
}

impl Clone for ComRuntime {
    fn clone(&self) -> Self {
        ComRuntime(Rc::clone(&self.0))
    }
}

// impl Drop for ComRuntime {
//     fn drop(&mut self) {
//         if Rc::strong_count(&self.0) == 1 {
//             println!("go time");
//             unsafe {
//                 CoUninitialize();
//             }
//         } else {
//             println!("not time yet");
//         }
//     }
// }

#[derive(Debug, PartialEq)]
struct Com;

impl Com {
    /// Allows us to just initialize once over the course of the program
    fn initialize() -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)?;
        }

        Ok(Com)
    }
}

impl Drop for Com {
    fn drop(&mut self) {
        println!("dropping com");
        unsafe {
            CoUninitialize();
        }
    }
}

/// The base struct, used for making Schedules and getting them
pub struct TaskScheduler {
    task_service: ITaskService,
    com_runtime: ComRuntime,
}

impl TaskScheduler {
    /// makes a new thing
    pub fn new() -> Self {
        let com = ComRuntime::new().unwrap();

        unsafe {
            let task_service: ITaskService =
                CoCreateInstance(&TaskScheduler, None, CLSCTX_ALL).unwrap();

            task_service
                .Connect(
                    VARIANT::default(),
                    VARIANT::default(),
                    VARIANT::default(),
                    VARIANT::default(),
                )
                .unwrap();

            TaskScheduler {
                com_runtime: com,
                task_service,
            }
        }
    }

    /// Create a new `ScheduleBuilder` that can be used to make a schedule.
    pub fn create_schedule(&self) -> ScheduleBuilder<Base> {
        ScheduleBuilder::new(&self.com_runtime, &self.task_service).unwrap()
    }

    /// Gets the task at the specified path, returns something on error
    pub fn get_schedule(&self, path: &str) -> Schedule<Registered> {
        let folder;
        let name;

        // if the name has no \, it is assumed to be in the root folder
        match path.rsplit_once('\\') {
            Some(x) if x.0.is_empty() => (folder, name) = ("\\", x.1),
            Some(x) => (folder, name) = x,
            None => {
                folder = "\\";
                name = path;
            }
        };

        unsafe {
            // TODO: match on "task not existing"
            let com_runtime = self.com_runtime.clone();
            let task_folder = self.task_service.GetFolder(&BSTR::from(folder)).unwrap();
            let registered_task = task_folder.GetTask(&BSTR::from(name)).unwrap();
            let task_definition = registered_task.Definition().unwrap();
            let registration_info = task_definition.RegistrationInfo().unwrap();
            let actions = task_definition.Actions().unwrap();
            let settings = task_definition.Settings().unwrap();
            let triggers = task_definition.Triggers().unwrap();

            Schedule::<Registered> {
                kind: std::marker::PhantomData::<Registered>,
                task_folder,
                task_definition,
                registration_info,
                actions,
                force_start_boundary: false,
                settings,
                trigger: None,
                triggers,
                com_runtime,
            }
        }
    }

    /// Gets all the schedules in a folder
    pub fn get_schedules_from_folder(&self, folder: &str) -> Option<Vec<Schedule<Registered>>> {
        unsafe {
            // TODO: match on "task not existing"
            let task_folder = self.task_service.GetFolder(&BSTR::from(folder)).unwrap();

            let task_collection = task_folder.GetTasks(TASK_ENUM_HIDDEN.0).unwrap();
            let task_num = task_collection.Count().unwrap();

            if task_num == 0 {
                return None;
            }

            let mut schedules = Vec::with_capacity(task_num as usize);

            for i in 0..task_num {
                let index = VARIANT {
                    Anonymous: VARIANT_0 {
                        Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                            vt: VT_I4,
                            wReserved1: 0,
                            wReserved2: 0,
                            wReserved3: 0,
                            Anonymous: VARIANT_0_0_0 { lVal: i + 1 },
                        }),
                    },
                };

                let com_runtime = self.com_runtime.clone();
                let registered_task = task_collection.get_Item(index).unwrap();

                let task_definition = registered_task.Definition().unwrap();
                let registration_info = task_definition.RegistrationInfo().unwrap();
                let actions = task_definition.Actions().unwrap();
                let settings = task_definition.Settings().unwrap();
                let triggers = task_definition.Triggers().unwrap();

                schedules.push(Schedule::<Registered> {
                    kind: std::marker::PhantomData::<Registered>,
                    task_folder: task_folder.clone(),
                    task_definition,
                    registration_info,
                    actions,
                    force_start_boundary: false,
                    settings,
                    trigger: None,
                    triggers,
                    com_runtime,
                });
            }

            Some(schedules)
        }
    }
}

impl Default for TaskScheduler {
    fn default() -> Self {
        Self::new()
    }
}
