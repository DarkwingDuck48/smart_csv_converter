pub mod cli;
pub mod config;
pub mod errors;
pub mod celladdress;
pub mod namedrange;
#[cfg(test)]
mod test;
mod parse_with_config;
mod parse_with_manual;
mod excel_parser;

use path::PathBuf;
use clap::Parser;
use std::path;
use log::{info, LevelFilter};
use crate::config::{Config, create_config_from_toml};

/// Priority to CLI arguments, next from config
/// If in CLI have --config parameter : all options read from toml config
/// Else - read options from CLI arguments.
fn main() {

    // Parse CLI arguments with clap
    let cli = cli::Cli::parse();
    // Debug mode
    let debug_mode = cli.debug;

    // Create log file
    let log_path_default = PathBuf::from(".\\default.log");
    let log_level:LevelFilter;
    match debug_mode {
        true => log_level = LevelFilter::Debug,
        false => log_level = LevelFilter::Info,
    }
    match cli.log_file {
        Some(log_file) => {
            simple_logging::log_to_file(log_file, log_level).expect("Cant find log file");
        },
        None => {
            simple_logging::log_to_file(log_path_default, log_level).expect("Cant find log file");
        }
    };
    info!("Current log_level set to {log_level}");
    match cli.config {
        Some(config_path) => {
            info!("Passed config {:?} - start working with config!", config_path);
            let config: Config = create_config_from_toml(config_path);
        },
        None => {
            info!("No config file passed in arguments -> work with command args:");
            let source_file: PathBuf = cli.source_file.expect("Not provided source file path");
            let target_file: PathBuf = cli.target_file.expect("Not provided target file path");
            let parsed_sheets: Vec<String> = cli.parsed_sheets;
            let named_ranges: Vec<String> = cli.ranges;

            info!("Source file path: {:?}", source_file);
            info!("Target file path: {:?}", target_file);
            info!("Parsed sheets list {:?}", parsed_sheets);
            info!("Named range list {:?}", named_ranges);

        }
    }
}
