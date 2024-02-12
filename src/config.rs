use std::path::PathBuf;
use toml::Table;
pub struct Config {
    pub source_file: PathBuf,
    pub target_file: PathBuf,
    pub sheets_list: Vec<String>,
    pub named_range: Vec<String>,
    pub tables_list: Vec<String>,
    pub debug: bool,
}
//
// impl Config {
//     pub fn new(config_path: PathBuf) -> Self {
//         /// Implement toml config parcer
//
//     }
// }
