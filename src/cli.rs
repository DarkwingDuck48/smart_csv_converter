use clap::Parser;
use std::path::PathBuf;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
pub struct Cli {
    /// Path to source file
    pub source_file: PathBuf,

    /// Path to target file
    pub target_file: Option<PathBuf>,

    /// Parsed sheets in file
    #[arg(short='s',long="sheets", value_delimiter = ' ', num_args = 1..)]
    pub parsed_sheets: Vec<String>,

    /// Analyze tables in file
    #[arg(short, long, action)]
    pub tables: bool,

    /// Debug Mode
    #[arg(short, long, action)]
    pub debug: bool,
}
