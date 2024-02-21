/// Module with TOML schema struct
use std::path::PathBuf;
use serde::Deserialize;

#[derive(Deserialize, Debug, Clone)]
pub struct FileSettings {
    pub source: Option<SourceFile>,
    pub target: Option<TargetFile>,
}

#[derive(Deserialize, Debug, Clone)]
pub struct SourceFile {
    pub path: Option<PathBuf>,
    pub sheets: Option<Vec<String>>,
}

#[derive(Deserialize, Debug, Clone)]
pub struct TargetFile {
    pub path: Option<PathBuf>,
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
    conditions: Option<Vec<String>>,
}

#[derive(Deserialize, Debug, Clone)]
pub struct LocalSheet {
    pub path: Option<PathBuf>,
    pub separator: Option<char>,
    pub columns: Option<Vec<String>>,
    pub named_ranges: Option<Vec<String>>,
    pub conditions: Option<Vec<String>>,
}