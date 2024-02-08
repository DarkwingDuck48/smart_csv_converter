/// TODO:
/// 1. Implement Config module -> parse from toml file
/// 2. Parse excel fle
/// 3. Issue with smart tables in source files: when no tables in file raises error
pub mod cli;
pub mod config;

use crate::config::Config;
use calamine::{open_workbook, DataType, Reader, Xlsx};
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
    }

    let defined_names = workbook.defined_names();
    println!("Defined Names: {:?}", defined_names);
    let csv_file = OpenOptions::new()
        .write(true)
        .create(true)
        .append(true)
        .open(&config.target_file)
        .unwrap();
    let mut wtr = Writer::from_writer(csv_file);
    for sheet_name in worksheets {
        // Read whole worksheet data and provide some statistics
        if let Ok(range) = workbook.worksheet_range(&sheet_name) {
            if cli.debug {
                let total_cells = range.get_size().0 * range.get_size().1;
                let non_empty_cells: usize = range.used_cells().count();
                println!(
                    "Found {} cells in {}, including {} non empty cells",
                    total_cells, sheet_name, non_empty_cells
                );
            }
        }

        let sheet = workbook.worksheet_range(&sheet_name).unwrap();
        wtr.write_record(vec![sheet_name]).expect("Cant write row");
        // Write data from sheet to target csv file
        // for _row in sheet.rows() {
        //     println!("row={:?}", _row);
        //     let mut ele_row = vec![];
        //     for ele in _row {
        //         match ele {
        //             DataType::Int(ele) => ele_row.push(ele.clone().to_string()),
        //             DataType::Empty => ele_row.push(" "),
        //
        //         }
        //
        //     }
        //     //wtr.write_record(_row).expect("Cant write row");
        // }
    }
}
