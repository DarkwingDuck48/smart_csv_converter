use crate::celladdress::CellAddress;
use crate::config::{GlobalSheet, LocalSheet, Sheets};

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

#[test]
fn test_convert_range_to_cell_full_dollar_sign_range() {
    assert_eq!(
        CellAddress::convert_excel_cell_address_to_numbers("$AA$2"),
        CellAddress {
            column: 26u32,
            row: 1u32,
        }
    )
}

#[test]
fn test_convert_range_to_cell_one_dollar_sign_row() {
    assert_eq!(
        CellAddress::convert_excel_cell_address_to_numbers("AA$2"),
        CellAddress {
            column: 26u32,
            row: 1u32,
        }
    )
}

#[test]
fn test_convert_range_to_cell_one_dollar_sign_column() {
    assert_eq!(
        CellAddress::convert_excel_cell_address_to_numbers("$AA2"),
        CellAddress {
            column: 26u32,
            row: 1u32,
        }
    )
}


#[test]
fn test_convert_range_to_cell_no_dollar_sign() {
    assert_eq!(
        CellAddress::convert_excel_cell_address_to_numbers("AA2"),
        CellAddress {
            column: 26u32,
            row: 1u32,
        }
    )
}

#[test]
fn test_localsheet_name_with_string() {
    assert_eq!(
        String::from("Test1"),
        LocalSheets {
            sheet_name: String::from("Test1"),
            path: None,
            separator: None,
            columns: None,
            named_ranges: None,
            checks: None,
        }
    )
}

#[test]
fn test_string_in_localsheet_vec() {
    let sheets = Sheets {
        global: GlobalSheet{
            columns: None,
            named_ranges: None,
            tables:None,
            checks: None
        },
        local: Some(vec![
            LocalSheet {
                sheet_name:String::from("Test1"), 
                path: None, 
                separator: None, 
                columns: None, 
                named_ranges: None, 
                checks: None },
            LocalSheet {
                sheet_name:String::from("Test2"), 
                path: None, 
                separator: None, 
                columns: None, 
                named_ranges: None, 
                checks: None },
            LocalSheet {
                sheet_name:String::from("Test3"), 
                path: None, 
                separator: None, 
                columns: None, 
                named_ranges: None, 
                checks: None }
        ]),
    };
    let checked_sheet_name = String::from("Test4");
    let local_sheets_vec = sheets.local.unwrap();
    let found_sheet_name = local_sheets_vec.iter().find(|&lh| lh.sheet_name.eq(&checked_sheet_name));
    println!("Found sheet name {:?}", found_sheet_name);
    assert_eq!(checked_sheet_name, found_sheet_name.unwrap().sheet_name);

}
