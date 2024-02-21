pub mod cli;
pub mod config;
pub mod errors;
pub mod celladdress;
pub mod namedrange;
mod test;
use path::PathBuf;
use crate::{
    config::{SourceFile, TargetFile, GlobalSheets, LocalSheet},
    namedrange::NamedRange,
    namedrange::get_named_ranges,
};
use calamine::{open_workbook, Data, Range, Reader, Xlsx};
use clap::Parser;
use std::{fs, path};
use toml::Table;
use std::process::exit;
use serde::Serialize;
use crate::config::{FileSettings, Sheets};

/// Priority to CLI arguments, next from config
/// If in CLI have --config parameter : all options read from toml config
/// Else - read options from CLI arguments.
fn main() {
    let config_path = PathBuf::from("D:\\Work\\GIT_work\\rust_excel_to_csv_conterter\\excel_to_csv\\test_config.toml");
    let content = fs::read_to_string(&config_path).expect("Could not read file");
    let data: Table = content.parse().unwrap();
    let file_settings: FileSettings = data["file"].clone().try_into().unwrap();
    let sheet_settings: Sheets = match data["sheets"].clone().try_into() {
        Ok(d) => d,
        Err(e) => {
            eprintln!("TOML parse error: {:?}", e.message());
            exit(1);
        }
    };
    //TODO: Exception handling when TomlError
    println!("{:#?}", sheet_settings);
    // println!("{:#?}", file_settings.source);
    // println!("{:#?}", file_settings.target);
}
