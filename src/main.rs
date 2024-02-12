pub mod cli;
pub mod config;
pub mod errors;
mod test;


use crate::config::Config;
use crate::errors::ExcelLetterConvertError;
use regex::Regex;
use calamine::{open_workbook, Data, Range, Reader, Xlsx};
use clap::Parser;
use csv::Writer;
use log::{LevelFilter};
use log::{debug, error, info, warn};
use std::fs::{File, OpenOptions};
use std::io::BufReader;
use std::path::PathBuf;
use std::string::String;
use std::borrow::BorrowMut;
use std::string;


#[derive(Debug, Default)]
struct NamedRange {
    name: String,
    sheet_name: String,
    range: (CellAddress, CellAddress),
}

impl NamedRange {
    fn new(name: String, sheet_name: String, range: (CellAddress, CellAddress)) -> Self {
        Self { name, sheet_name, range }
    }
    fn init_default() -> Self {
        Self {
            name: String::from(""),
            sheet_name: String::from(""),
            range: (
                CellAddress::new(0, 0),
                CellAddress::new(0, 0)
            ),
        }
    }
}


#[derive(Debug, PartialEq, Default)]
struct CellAddress {
    column: u32,
    row: u32,
}

impl CellAddress {
    fn new(column: u32, row: u32) -> Self {
        Self { column, row }
    }
    fn to_tuple(self) -> (u32, u32) {
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
        return match letter.to_uppercase().last().unwrap() {
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
        };
    }
}

fn convert_excel_range_to_numbers(rng: &str) -> CellAddress {
    // Columns and rows starts from 0
    let parts: Vec<_> = rng.split("$").collect();
    let column_letter = parts[1];
    let row_number: u32 = parts[2].to_string().parse().unwrap();
    CellAddress::from_excel(column_letter, row_number)
}

fn parse_defined_name(name: &String, range_address: &String) -> NamedRange {
    let regex_range = Regex::new(r"(?<sheetname>\w*)?!?(?<startcell>\$?\w+\$?\d+):?(?<endcell>\$?\w+\$?\d+)?").unwrap();
    let parts = regex_range.captures(range_address).unwrap();
    let defined_sheet_name = parts.name("sheetname").map_or("", |m| m.as_str());
    let start_cell_address = convert_excel_range_to_numbers(parts.name("startcell").map_or("", |m| m.as_str()));
    let end_cell_address = convert_excel_range_to_numbers(parts.name("endcell").map_or(parts.name("startcell").unwrap().as_str(), |m| m.as_str()));
    NamedRange::new(name.to_string(), defined_sheet_name.to_string(), (start_cell_address, end_cell_address))
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

    let mut named_range: NamedRange = Default::default();

    for def_name in workbook.defined_names() {
        println!("{:?}", def_name);
        named_range = parse_defined_name(
            &def_name.0,
            &def_name.1,
        );
    }
    let sheet_ref = workbook.worksheet_range(&named_range.sheet_name).expect("No sheet in workbook");
    let named_sheet_range = sheet_ref.range(named_range.range.0.to_tuple(), named_range.range.1.to_tuple());
    println!("Named range name: {:?}", named_range.name);
    println!("Named range sheet_name: {:?}", named_range.sheet_name);
    println!("{:?}", named_sheet_range);

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
