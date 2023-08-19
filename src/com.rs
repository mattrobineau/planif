use std::rc::Rc;
use windows::Win32::System::Com::{
    CoInitializeEx, CoUninitialize, COINIT_MULTITHREADED
};

/// Represents a COM runtime required for building schedules tasks
#[derive(Clone)]
pub struct ComRuntime(Rc<Com>);

impl ComRuntime {
    /// Creates a COM runtime for use with one or more
    /// [ScheduleBuilder]
    pub fn new() -> Result<Self, Box<dyn std::error::Error>> {
        Ok(ComRuntime(Rc::new(Com::initialize()?)))
    }
}

struct Com;

impl Com {
    fn initialize() -> Result<Self, Box<dyn std::error::Error>> {
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)?;
        }
        Ok(Com)
    }
}

impl Drop for Com {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

