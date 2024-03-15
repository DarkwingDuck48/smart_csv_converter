use std::fs;
/// Config struct for project
/// Keep in one place all settings for parser
/// And use as Deserializer for TOML
use std::path::PathBuf;
use std::process::exit;
use serde::Deserialize;
use toml::Table;

pub fn create_config_from_toml(config_path: PathBuf) -> Config {
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
    println!("{:#?}", file_settings);
    let config: Config = Config { sheetsettings: sheet_settings, filesettings: file_settings };
    config
}

#[derive(Deserialize, Debug, Clone)]
pub struct Config {
    pub filesettings: FileSettings,
    pub sheetsettings: Sheets,
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
    pub global: Option<GlobalSheets>,
    pub local: Option<Vec<LocalSheet>>,
}
#[derive(Deserialize, Debug, Clone)]
pub struct GlobalSheets {
    columns: Option<Vec<String>>,
    named_ranges: Option<Vec<String>>,
    conditions: Option<Vec<String>>
}

#[derive(Deserialize, Debug, Clone)]
pub struct LocalSheet {
    pub path: Option<PathBuf>,
    pub separator: Option<char>,
    pub columns: Option<Vec<String>>,
    pub named_ranges: Option<Vec<String>>,
    pub conditions: Option<Vec<String>>
}