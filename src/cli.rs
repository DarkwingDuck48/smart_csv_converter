use clap::Parser;
use std::path::PathBuf;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
pub struct Cli {

    /// Debug Mode
    #[arg(short, long, action, default_value = "false")]
    pub debug: bool,

    /// Path to log file
    #[arg(short = 'l', long = "log")]
    pub log_file: Option<PathBuf>,

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
    pub named_ranges: Vec<String>,

    /// Path to config TOML file (only config file needed)
    #[arg(long = "config")]
    pub config: Option<PathBuf>,
}

pub struct CliArgs {
    //Debug mode
    pub debug: bool,
    //Source file path
    pub source_file: PathBuf,
    //Target file path
    pub target_file: PathBuf,
    //Parsed sheets list
    pub parsed_sheets: Vec<String>,
    // Parsed named ranges
    pub named_ranges: Vec<String>,

}