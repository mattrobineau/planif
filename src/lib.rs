//#![deny(missing_docs)]
pub mod enums;
pub mod error;
pub mod schedule;
pub mod schedule_builder;
pub mod settings;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        let result = 2 + 2;
        assert_eq!(result, 4);
    }
}
