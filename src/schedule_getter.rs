use std::mem::ManuallyDrop;

use windows::core::BSTR;
use windows::Win32::System::Com::*;
use windows::Win32::System::TaskScheduler::*;

struct ScheduleGetter {
    task_service: ITaskService,
}

impl ScheduleGetter {
    pub fn new() -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            // On error of unsafe, CoUnintialize!
            CoInitializeEx(None, COINIT_MULTITHREADED)?;

            let task_service: ITaskService = CoCreateInstance(&TaskScheduler, None, CLSCTX_ALL)?;
            task_service.Connect(
                VARIANT::default(),
                VARIANT::default(),
                VARIANT::default(),
                VARIANT::default(),
            )?;

            Ok(Self { task_service })
        }
    }

    pub fn get_folder_tasks(
        &self,
        folder_name: &str,
    ) -> Result<Option<Vec<IRegisteredTask>>, windows::core::Error> {
        unsafe {
            // Get the task folder.
            let task_folder = self.task_service.GetFolder(&BSTR::from(folder_name))?;

            let task_collection = task_folder.GetTasks(TASK_ENUM_HIDDEN.0)?;
            let task_num = task_collection.Count()?;

            if task_num == 0 {
                return Ok(None);
            }

            let mut tasks = Vec::with_capacity(task_num as usize);

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

                tasks.push(task_collection.get_Item(index)?);
            }
            Ok(Some(tasks))
        }
    }
}
