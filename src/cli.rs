use clap::ArgGroup;
use clap::Parser;
use std::path::PathBuf;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
#[clap(group(
ArgGroup::new("settings")
.required(true)
.args(& ["parsed_sheets", "named_ranges"])
))]
#[clap(group(ArgGroup::new("manual").multiple(true).args(& ["source_file", "target_file"])))]
pub struct Cli {
    /// Path to source file (manual import)
    #[arg(long, requires = "target_file")]
    pub source_file: Option<PathBuf>,

    /// Path to target file (manual import)
    #[arg(long, requires = "source_file")]
    pub target_file: Option<PathBuf>,

    // Ranges with defined names to parse on sheets
    #[arg(short = 'r', long = "ranges", value_delimiter = ' ', num_args = 1..)]
    pub named_ranges: Vec<String>,

    /// Parsed sheets in file (required source and target sheets)
    #[arg(short = 's', long = "sheets", value_delimiter = ' ', num_args = 1..)]
    pub parsed_sheets: Vec<String>,

    /// Debug Mode
    #[arg(short, long, action, default_value = "false")]
    pub debug: bool,

    /// Path to log file
    #[arg(short = 'l', long = "log")]
    pub log_file: Option<PathBuf>,

    /// Path to config TOML file (only config file needed)
    #[arg(long = "config", conflicts_with = "manual")]
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