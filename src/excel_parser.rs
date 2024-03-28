use std::fs::OpenOptions;
use std::path::PathBuf;
use calamine::{open_workbook, Reader, Xlsx};
use csv::Writer;
use crate::config::{Config, GlobalSheet, LocalSheet, Sheets};
use log::{debug, info, error};
use crate::namedrange::{NamedRange, get_named_ranges};


#[derive(Debug, Clone)]
struct SheetParsingSchema {
    pub sheet_name: String,
    pub path: PathBuf,
    pub separator: char,
    pub columns: Option<Vec<String>>,
    pub named_ranges: Option<Vec<String>>,
}


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
        only_global = true;
        info!("Only Global sheets section was found");
    };
    // Create default schema for parsing
    let global_sheet_config = config.sheetsettings.global;
    let global_sheet_schema = SheetParsingSchema {
        sheet_name: "".to_string(),
        path: config.filesettings.target.path,
        separator: config.filesettings.target.separator,
        columns: global_sheet_config.columns,
        named_ranges: config.filesettings.source.named_ranges,
    };

    // let output_file = OpenOptions::new()
    //     .write(true)
    //     .create(true)
    //     .truncate(true)
    //     .open(&global_sheet_schema.path)
    //     .unwrap();
    // let mut wtr = Writer::from_writer(output_file);

    for sheet_name in worksheets {
        // If we have special settings for some sheets => find them before parse
        let mut sheet_schema: SheetParsingSchema = global_sheet_schema.clone();
        if !only_global {
            // Working with local sheet config
            for sheet_cfg in config.sheetsettings.local.clone().unwrap() {
                if sheet_cfg.sheet_name.eq(&sheet_name) {
                    info!("Find local schema parsing for {:?}", sheet_name);
                    sheet_schema = SheetParsingSchema {
                        sheet_name: sheet_cfg.sheet_name,
                        path: sheet_cfg.path.unwrap_or(global_sheet_schema.path.clone()),
                        separator: sheet_cfg.separator.unwrap_or(global_sheet_schema.separator),
                        columns: sheet_cfg.columns,
                        named_ranges: sheet_cfg.named_ranges,
                    };
                    break;
                }
            }
        }
        println!("Sheet name : {:?}", sheet_name);
        println!("Sheet Schema: {:#?}", sheet_schema);

        if global_sheet_schema.path != sheet_schema.path {
            info!("Create additional file for export - {:?}", sheet_schema.path);
            error!("Not implemented functional - all data will be extracted to {:?}", global_sheet_schema.path);
        }
    }
}