use std::{error::Error, fmt};


#[derive(Debug)]
pub struct InvalidOperationError {
    pub message: String,
}

impl Error for InvalidOperationError {}

impl fmt::Display for InvalidOperationError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Invalid Operation Error: {}", self.message)
    }
}

#[derive(Debug)]
pub struct InitializationError {
    pub message: String,
}

impl Error for InitializationError {}

impl fmt::Display for InitializationError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Initialization Error: {}", self.message)
    }
}
