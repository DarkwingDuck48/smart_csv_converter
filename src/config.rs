use std::fs;
/// Config struct for project
/// Keep in one place all settings for parser
/// And use as Deserializer for TOML
use std::path::PathBuf;
use std::process::exit;
use log::info;
use serde::Deserialize;
use toml::Table;
use crate::cli::{Cli, CliArgs};


#[derive(Deserialize, Debug, Clone)]
pub struct Config {
    pub filesettings: FileSettings,
    pub sheetsettings: Sheets,
}

impl Config {
    /// Way to create smart Excel parser configuration
    pub fn from_toml(config_path: PathBuf) -> Config {
        let content = fs::read_to_string(&config_path).expect("Could not read file");
        let data: Table = content.parse().unwrap();
        let file_settings: FileSettings = data["file"].clone().try_into().unwrap();
        let sheet_settings: Sheets = match data["sheets"].clone().try_into() {
            Ok(d) => d,
            Err(e) => {
                eprintln!("TOML parse error: {:?}", e.message());
                exit(0);
            }
        };
        println!("{:#?}", file_settings);
        println!("{:#?}", sheet_settings);
        let config: Config = Config { sheetsettings: sheet_settings, filesettings: file_settings };
        config
    }

    /// Way to create only global sheet configuration
    pub fn from_cli(cli: CliArgs) -> Config {
        let source_file: SourceFile = SourceFile {
            path: cli.source_file,
            sheets: Option::Some(cli.parsed_sheets),
        };
        let target_file: TargetFile = TargetFile {
            path: cli.target_file,
            separator: Option::Some(','),
            columns: Option::None,
        };
        let global_sheet: GlobalSheets = GlobalSheets {
            columns: Option::None,
            named_ranges: Some(cli.named_ranges),
            tables: Option::None,
            checks: Option::None,
        };

        let config = Config {
            filesettings: FileSettings { source: source_file, target: target_file },
            sheetsettings: Sheets { global: global_sheet, local: Option::None },
        };
        config
    }
}

#[derive(Deserialize, Debug, Clone)]
pub struct FileSettings {
    pub source: SourceFile,
    pub target: TargetFile,
}

#[derive(Deserialize, Debug, Clone)]
pub struct SourceFile {
    pub path: PathBuf,
    pub sheets: Option<Vec<String>>,
}

#[derive(Deserialize, Debug, Clone)]
pub struct TargetFile {
    pub path: PathBuf,
    pub separator: Option<char>,
    pub columns: Option<Vec<String>>,
}

#[derive(Deserialize, Debug, Clone)]
pub struct Sheets {
    pub global: GlobalSheets,
    pub local: Option<Vec<LocalSheet>>,
}
#[derive(Deserialize, Debug, Clone)]
pub struct GlobalSheets {
    pub columns: Option<Vec<String>>,
    pub named_ranges: Option<Vec<String>>,
    pub tables: Option<Vec<String>>,
    pub checks: Option<Vec<String>>,
}

#[derive(Deserialize, Debug, Clone)]
pub struct LocalSheet {
    pub sheet_name: String,
    pub path: Option<PathBuf>,
    pub separator: Option<char>,
    pub columns: Option<Vec<String>>,
    pub named_ranges: Option<Vec<String>>,
    pub checks: Option<Vec<String>>,
}

impl PartialEq<String> for LocalSheet {
    fn eq(&self, other: &String) -> bool {
        other.eq(&self.sheet_name)
    }
}

impl PartialEq<LocalSheet> for String {
    fn eq(&self, other: &LocalSheet) -> bool {
        self.eq(&other.sheet_name)
    }
}