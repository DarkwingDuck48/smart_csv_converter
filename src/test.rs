use crate::celladdress::CellAddress;

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

