use crate::errors::ExcelLetterConvertError;

#[derive(Debug, PartialEq,Default, Clone, Copy)]
pub struct CellAddress {
    column: u32,
    row: u32,
}

impl CellAddress {

    pub fn to_tuple(self) -> (u32, u32) {
        (self.row, self.column)
    }
    /// Creates absolute cell address from Excel address, so it can be used for construct calamine cell objects
    /// A1 => CellAddress(0, 0)
    /// AA1 => CellAddress(26, 0)
    /// Note that columns and rows indexes starts from 0.
    pub fn from_excel(column_letter: &str, row: u32) -> Self {
        let chars_to_parse: Vec<char> = column_letter.chars().collect();
        let mut column_result;
        match chars_to_parse.len() {
            1 => column_result = 0u32,
            _ => column_result = ((chars_to_parse.len() - 1) * 25 + 1) as u32
        }

        column_result += Self::convert_letter_to_column(chars_to_parse.last().unwrap()).unwrap();
        Self {
            column: column_result,
            row: row - 1,
        }
    }
    fn convert_letter_to_column(letter: &char) -> Result<u32, ExcelLetterConvertError> {
        match letter.to_uppercase().last().unwrap() {
            'A' => Ok(0u32),
            'B' => Ok(1u32),
            'C' => Ok(2u32),
            'D' => Ok(3u32),
            'E' => Ok(4u32),
            'F' => Ok(5u32),
            'G' => Ok(6u32),
            'H' => Ok(7u32),
            'I' => Ok(8u32),
            'J' => Ok(9u32),
            'K' => Ok(10u32),
            'L' => Ok(11u32),
            'M' => Ok(12u32),
            'N' => Ok(13u32),
            'O' => Ok(14u32),
            'P' => Ok(15u32),
            'Q' => Ok(16u32),
            'R' => Ok(17u32),
            'S' => Ok(18u32),
            'T' => Ok(19u32),
            'U' => Ok(20u32),
            'V' => Ok(21u32),
            'W' => Ok(22u32),
            'X' => Ok(23u32),
            'Y' => Ok(24u32),
            'Z' => Ok(25u32),
            _ => Err(ExcelLetterConvertError(format!(
                "Cant parse letter {} to Excel column",
                letter
            ))),
        }
    }
}
