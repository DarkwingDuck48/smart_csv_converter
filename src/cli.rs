use clap::Parser;
use std::path::PathBuf;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
pub struct Cli {
    /// Path to log file
    pub log_file: Option<PathBuf>,

    /// Debug Mode
    #[arg(short, long, action, default_value = "false")]
    pub debug: bool,

    /// Path to source file (manual import)
    #[arg(long, group = "manual")]
    pub source_file: Option<PathBuf>,

    /// Path to target file (manual import)
    #[arg(long, group = "manual")]
    pub target_file: Option<PathBuf>,

    /// Parsed sheets in file (required source and target sheets)
    #[arg(short = 's', long = "sheets", value_delimiter = ' ', requires = "manual", num_args = 1..)]
    pub parsed_sheets: Vec<String>,

    // Ranges with defined names to parse on sheets
    #[arg(short = 'r', long = "ranges", value_delimiter = ' ', requires = "manual", num_args = 1..)]
    pub ranges: Vec<String>,

    /// Path to config TOML file (only config file needed)
    #[arg(long = "config")]
    pub config: Option<PathBuf>,
}
