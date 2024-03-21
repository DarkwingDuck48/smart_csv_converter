use calamine::{open_workbook, Xlsx};
use crate::config::Config;


// Default excel file parser
pub fn parse(config: Config) -> () {
    let mut workbook: Xlsx<_> = open_workbook(config.filesettings.source.path).expect("Cannot open file");
}


