use crate::celladdress::CellAddress;
use std::{fs::File, io::BufReader};
use regex::Regex;
use calamine::{Data, Range, Reader, Xls, Xlsx};



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

fn parse_defined_name(name: &str, range_address: &String) -> NamedRange {
    let regex_range = Regex::new(r"(?<sheetname>\w*)?!?(?<startcell>\$?\w+\$?\d+):?(?<endcell>\$?\w+\$?\d+)?").unwrap();
    let parts = regex_range.captures(range_address).unwrap();
    let defined_sheet_name = parts.name("sheetname").map_or("", |m| m.as_str());
    let start_cell_address = CellAddress::convert_excel_range_to_numbers(parts.name("startcell").map_or("", |m| m.as_str()));
    let end_cell_address = CellAddress::convert_excel_range_to_numbers(parts.name("endcell").map_or(parts.name("startcell").unwrap().as_str(), |m| m.as_str()));
    NamedRange::new(name.to_string(), defined_sheet_name.to_string(), (start_cell_address, end_cell_address))
}

pub fn get_named_ranges(wb: &mut Xls<BufReader<File>>) -> Vec<NamedRange> {
    todo!()
}