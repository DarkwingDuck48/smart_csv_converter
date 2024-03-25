use calamine::{open_workbook, Reader, Xlsx};
use crate::config::{Config, GlobalSheet, LocalSheet, Sheets};
use log::{debug, info, error};
use crate::namedrange::{NamedRange, get_named_ranges};

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

    info!("Check localsheet section in config");
    let mut only_global = false;

    if config.sheetsettings.local.is_none() {
        only_global = true
    }

    for sheet_name in worksheets {
        let mut sheet_schema;
        if only_global {
            sheet_schema = config.sheetsettings.global.clone()
        }
    }
}


