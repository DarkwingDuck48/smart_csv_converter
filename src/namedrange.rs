use std::{fs::File, io::BufReader};

use crate::celladress::CellAddress;
use calamine::{Xlsx, Range, Data, Reader};



#[derive(Debug, Default, Clone)]
pub struct NamedRange {
    pub name: String,
    pub sheet_name: String,
    pub range: (CellAddress, CellAddress),
}

impl NamedRange {
    pub fn new(name: String, sheet_name: String, range: (CellAddress, CellAddress)) -> Self {
        Self { name, sheet_name, range }
    }
    pub fn construct_range(self, wb: &mut Xlsx<BufReader<File>>) -> Range<Data> {
        let sheet = wb.worksheet_range(&self.sheet_name).expect("No sheet in workbook");
        sheet.range(self.range.0.to_tuple(), self.range.1.to_tuple())

    }
}