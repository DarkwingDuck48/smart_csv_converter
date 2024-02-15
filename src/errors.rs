use std::fmt;
use std::fmt::Formatter;

#[derive(Debug, Clone)]
pub struct ExcelLetterConvertError(pub String);

impl fmt::Display for ExcelLetterConvertError {
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        write!(f, "Cant parse letter to column")
    }
}
