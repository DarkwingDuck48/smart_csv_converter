use std::path::PathBuf;
use log::debug;
use crate::config::{Config, create_config};

pub fn parse_with_config(config_path: PathBuf, debug_flag: bool) {
    // let config_path = PathBuf::from("D:\\Work\\GIT_work\\rust_excel_to_csv_conterter\\excel_to_csv\\test_config.toml");
    let config: Config = create_config(config_path);
    debug!("Config result: {:?}", config);

}