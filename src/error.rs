use std::{error::Error, fmt};

#[derive(Debug)]
/// The `InvalidOperationError` identifies errors where operations (function calls) are made with
/// invalid or incomplete data.
pub struct InvalidOperationError {
    /// A message with additional information on the error.
    pub message: String,
}

impl Error for InvalidOperationError {}

impl fmt::Display for InvalidOperationError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Invalid Operation Error: {}", self.message)
    }
}

#[derive(Debug)]
/// The `RequiredPropertyError` identifies errors when attempting to build a [Schedule](crate::schedule::Schedule) where
/// required builder functions have not been called.
pub struct RequiredPropertyError {
    /// A message with additional information on the error.
    pub message: String,
}

impl Error for RequiredPropertyError {}

impl fmt::Display for RequiredPropertyError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "Required Property Error: {}", self.message)
    }
}
