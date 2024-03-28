use std::fs;
/// Config struct for project
/// Keep in one place all settings for parser
/// And use as Deserializer for TOML
use std::path::PathBuf;
use std::process::exit;
use log::error;
use serde::Deserialize;
use toml::Table;
use crate::cli::CliArgs;


#[derive(Deserialize, Debug, Clone)]
pub struct Config {
    pub source: SourceFile,
    pub target: TargetFile,
}

impl Config {
    /// Way to create smart Excel parser configuration
    pub fn from_toml(config_path: PathBuf) -> Config {
        let content = fs::read_to_string(&config_path).expect("Could not read file");
        let data: Table = content.parse().unwrap();
        let config: Config = data["file"].clone().try_into().unwrap();
        if config.source.named_ranges.is_none() && config.source.tables.is_none() && config.source.sheets.is_none() {
            error!("No defined source parameters to parse. Sections Named Ranges, Tables and Sheets is empty!");
            panic!("No defined source parameters to parse. Sections Named Ranges, Tables and Sheets is empty!");
        }
        config
    }

    /// Way to create only global sheet configuration
    pub fn from_cli(cli: CliArgs) -> Config {
        if cli.parsed_sheets.is_empty() && cli.named_ranges.is_empty() {
            error!("No defined source parameters to parse. Sections Named Ranges, Tables and Sheets is empty!");
            panic!("No defined source parameters to parse. Sections Named Ranges, Tables and Sheets is empty!");
        }

        let source_file: SourceFile = SourceFile {
            path: cli.source_file,
            sheets: Option::Some(cli.parsed_sheets),
            named_ranges: Some(cli.named_ranges),
            tables: Option::None,
        };
        let target_file: TargetFile = TargetFile {
            path: cli.target_file,
            separator: ',',
            columns: Option::None,
        };
        let config = Config { source: source_file, target: target_file };
        config
    }
}

/// Usage logic for struct:
/// User must define at least one of sheets or named_ranges or tables
/// We do not want to parse full workbook and all sheets in workbook
/// If all of this params is empty - we have panic exception
pub struct SourceFile {
    pub path: PathBuf,
    pub sheets: Option<Vec<String>>,
    pub named_ranges: Option<Vec<String>>,
    pub tables: Option<Vec<String>>,
}

#[derive(Deserialize, Debug, Clone)]
pub struct TargetFile {
    pub path: PathBuf,
    pub separator: char,
    pub columns: Option<Vec<String>>,
}