/// TODO:
/// 1. Implement Config module -> parse from toml file
/// 2. Parse excel fle
pub mod cli;
pub mod config;

use std::string::String;
use crate::config::Config;
use calamine::{open_workbook, Data, Reader, Xlsx};
use clap::Parser;
use csv::Writer;
use std::fs::OpenOptions;

fn main() {
    let cli = cli::Cli::parse();
    println!("Debug mode: {:?}", cli.debug);

    let target_path;

    // Create target_path for file
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
        println!("Source Path is {:?}", config.source_file.display());
        println!("Target Path is {:?}", config.target_file.display());
    }

    let mut workbook: Xlsx<_> = open_workbook(config.source_file).expect("Cannot open file");
    let tables_list;
    println!("Options sheets : {:?}", config.sheets_list);

    // Get worksheets for parsing and compare it with book
    let mut worksheets: Vec<String>;

    if config.sheets_list.is_empty() {
        worksheets = workbook.sheet_names();
    } else {
        worksheets = Vec::new();
        for sheet in config.sheets_list {
            if !workbook.sheet_names().contains(&sheet) {
                println!("Not found {} in source file", sheet)
            } else {
                worksheets.push(sheet);
            }
        }
    }
    if worksheets.is_empty() {
        panic!("No list for parsing!!!!");
    }
    println!("Sheets to parse {:?}", worksheets);

    if cli.tables {
        // Not worked properly code - if no tables in source file caused error
        workbook.load_tables().unwrap();
        tables_list = workbook.table_names();
        println!("{:?}", tables_list)
    }

    let defined_names = workbook.defined_names();
    println!("Defined Names: {:?}", defined_names);
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
        println!("Work with {:?}", sheet_name);
        // Write data from sheet to target csv file
        for _row in sheet.rows() {
            println!("row={:?}", _row);
            let mut rows_values: Vec<String> = Vec::new();
            for ele in _row {
                let field;
                match ele {
                    Data::DateTime(ele) => field = ele.to_string(),
                    Data::Float(ele) => field = ele.to_string(),
                    Data::String(ele) => field = ele.to_string(),
                    Data::Empty => field = String::from(""),
                    _ => field = String::from("No Type")
                };
                rows_values.push(field);
            }
            wtr.write_record(rows_values).expect("Can't write row");
        }
    }
    wtr.flush().unwrap();
}
