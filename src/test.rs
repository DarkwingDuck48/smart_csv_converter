use log::debug;
use crate::CellAddress;
use regex::Regex;
use serde::de::Unexpected::Str;

#[test]
fn test_cell_address_one_letter() {
    assert_eq!(
        CellAddress::from_excel("A", 1u32),
        CellAddress {
            column: 0u32,
            row: 0u32,
        }
    )
}

#[test]
fn test_cell_address_two_letter() {
    assert_eq!(
        CellAddress::from_excel("AA", 1u32),
        CellAddress {
            column: 26u32,
            row: 0u32,
        }
    )
}
