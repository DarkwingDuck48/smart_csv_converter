/// TODO:
/// 1. Implement Config module -> parse from toml file
/// 2. Parse excel fle
pub mod cli;
pub mod config;

use std::string::String;
use std::fs::{File, OpenOptions};
use std::path::PathBuf;
use calamine::{open_workbook, Data, Reader, Xlsx, Range, Cell};
use clap::Parser;
use csv::Writer;
use log::{LevelFilter};
use log::{info, warn, debug, error};
use crate::config::Config;

#[derive(Debug)]
struct AbsoluteCellAddress {
    column: u32,
    row: u32,
}

impl AbsoluteCellAddress {}


fn convert_range_to_numbers(rng: String) -> AbsoluteCellAddress {
    /// Columns and rows starts from 0
    let parts: Vec<_> = rng.split("$").collect();
    let column_letter = parts[1];
    let row_number = parts[2];
    println!("Column {:?}; Row {row_number} ", column_letter);
    AbsoluteCellAddress {
        column: 0,
        row: 0,
    }
}

fn parse_defined_name(name: &String, range_address: &String) {
    debug!("Work with {name} and address is {range_address}");
    let parts: Vec<_> = range_address.split("!").collect();
    println!("SheetName: {:?}", parts[0]);
    println!("Range Address: {}", parts[1]);
    let rng_parts: Vec<_> = parts[1].split(":").collect();
    let start_rng = rng_parts[0].to_string();
    let end_rng = rng_parts[1].to_string();
    println!("Start cell address {start_rng}");
    println!("End cell address {end_rng}");
    convert_range_to_numbers(start_rng);
}

fn parse_sheet(sheet: Range<Data>, wtr: &mut Writer<File>) {
    // Write data from sheet to target csv file
    for _row in sheet.rows() {
        debug!("row={:?}", _row);
        let mut rows_values: Vec<String> = Vec::new();
        for ele in _row {
            let field;
            match ele {
                Data::DateTime(ele) => field = ele.as_datetime().unwrap().date().to_string(),
                Data::Int(ele) => field = ele.to_string(),
                Data::Float(ele) => field = ele.to_string(),
                Data::String(ele) => field = ele.to_string(),
                Data::Error(ele) => field = ele.to_string(),
                Data::Empty => field = String::from(""),
                _ => field = String::from("No Type")
            };
            rows_values.push(field);
        }
        wtr.write_record(rows_values).unwrap();
    }
}


fn main() {
    let cli = cli::Cli::parse();

    // Create logger for process
    let log_path;
    match cli.log_file {
        Some(log_file) => log_path = log_file,
        None => log_path = PathBuf::from(".\\test.log"),
    }

    simple_logging::log_to_file(log_path, LevelFilter::Debug).expect("Cant find log file");
    info!("Debug mode: {:?}", cli.debug);

    // Create target_path for file
    let target_path;
    match cli.target_file {
        Some(target_file) => target_path = target_file,
        None => target_path = cli.source_file.with_extension("csv"),
    }

    let config = Config {
        source_file: cli.source_file,
        target_file: target_path,
        sheets_list: cli.parsed_sheets,
        named_range: vec![],
        tables_list: vec![],
        debug: cli.debug,
    };

    if config.debug {
        info!("Source Path is {:?}", config.source_file.display());
        info!("Target Path is {:?}", config.target_file.display());
    }

    let mut workbook: Xlsx<_> = open_workbook(config.source_file).expect("Cannot open file");
    let tables_list;
    info!("Options sheets : {:?}", config.sheets_list);

    // Get worksheets for parsing and compare it with book
    let mut worksheets: Vec<String>;

    if config.sheets_list.is_empty() {
        worksheets = workbook.sheet_names();
    } else {
        worksheets = Vec::new();
        for sheet in config.sheets_list {
            if !workbook.sheet_names().contains(&sheet) {
                info!("Not found {} in source file", sheet)
            } else {
                worksheets.push(sheet);
            }
        }
    }
    if worksheets.is_empty() {
        panic!("No list for parsing!!!!");
    }
    info!("Sheets to parse {:?}", worksheets);

    if cli.tables {
        workbook.load_tables().unwrap();
        tables_list = workbook.table_names();
        info!("{:?}", tables_list)
    }

    for def_name in workbook.defined_names() {
        parse_defined_name(&def_name.0, &def_name.1);
    }
    let csv_file = OpenOptions::new()
        .write(true)
        .create(true)
        .truncate(true)
//        .append(true)
        .open(&config.target_file)
        .unwrap();
    let mut wtr = Writer::from_writer(csv_file);
    for sheet_name in worksheets {
        let sheet = workbook.worksheet_range(&sheet_name).unwrap();
        debug!("Working with {sheet_name}");
        parse_sheet(sheet, &mut wtr);
    }
    wtr.flush().unwrap();
}
