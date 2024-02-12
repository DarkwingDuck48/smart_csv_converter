pub mod cli;
pub mod config;
pub mod errors;
pub mod celladdress;
pub mod namedrange;
mod test;


use crate::config::Config;
use crate::namedrange::NamedRange;


use calamine::{open_workbook, Data, Range, Reader, Xlsx};
use clap::Parser;
use csv::Writer;
use log::LevelFilter;
use log::{debug, error, info, warn};
use std::fs::{File, OpenOptions};
use std::path::PathBuf;
use std::string::String;





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
                _ => field = String::from("No Type"),
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

    let source_path: PathBuf;
    match cli.source_file {
        source_file => source_path = source_file,
        _ => source_path = PathBuf::from("D:\\Work\\GIT_work\\rust_excel_to_csv_conterter\\excel_to_csv\\example_excel\\test1.xlsx"),
    }
    // Create target_path for file
    let target_path;
    match cli.target_file {
        Some(target_file) => target_path = target_file,
        None => target_path = source_path.with_extension("csv"),
    }

    let config = Config {
        source_file: source_path,
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

    let mut named_range: Vec<NamedRange> = vec![];

    for def_name in workbook.defined_names() {
        println!("{:?}", def_name);
        let tmp_range = namedrange::parse_defined_name(
            &def_name.0,
            &def_name.1,
        );
        named_range.push(tmp_range);
    }
    println!("{:?}", named_range.len());
    for nrange in named_range {
        println!("Named range name: {:?}", nrange.name);
        println!("Named range sheet_name: {:?}", nrange.sheet_name);
        let named_sheet_range = nrange.construct_range(&mut workbook);
        println!("{:?}", named_sheet_range);
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
        parse_sheet(sheet, &mut wtr);
    }
    wtr.flush().unwrap();
}
