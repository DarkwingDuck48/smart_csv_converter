use calamine::{open_workbook, Reader, Xlsx};
use crate::config::Config;
use log::{debug, info, error};
use crate::namedrange::{NamedRange, get_named_ranges};


fn check_sheetname_in_config(sheet_name: &String, configSheets: Vec<String>) {}

// Default excel file parser
pub fn parse(config: Config) -> () {
    info!("Start parse excel file {}", config.filesettings.source.path.display());
    let mut workbook: Xlsx<_> = open_workbook(config.filesettings.source.path).expect("Cannot open file");

    // Get sheets for parsing
    let mut worksheets: Vec<String>;
    if config.filesettings.source.sheets.is_none() {
        worksheets = workbook.sheet_names();
    } else {
        worksheets = Vec::new();
        for sheet in config.filesettings.source.sheets.unwrap() {
            if !workbook.sheet_names().contains(&sheet) {
                info!("Not found {} in source file", sheet)
            } else {
                worksheets.push(sheet);
            }
        }
    }
    if worksheets.is_empty() {
        panic!("No list for parsing!!!");
    }
    info!("Sheets to parse {:?}", worksheets);

    for sheet_name in worksheets {}
}


