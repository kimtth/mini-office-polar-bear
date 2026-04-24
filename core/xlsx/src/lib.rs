//! XLSX format support — import and export between `.xlsx` files and `CalcDocument`.

mod reader;
mod writer;

pub use reader::read_xlsx;
pub use writer::write_xlsx;

#[cfg(test)]
mod tests {
    use super::*;
    use pbo_common::model::calc::{
        CalcDocument, Cell, CellValue, MergedRange, Sheet, SheetHyperlink,
    };
    use std::collections::BTreeMap;

    #[test]
    fn default_calc_document_is_empty() {
        let doc = CalcDocument::default();
        assert!(doc.sheets.is_empty());
    }

    #[test]
    fn round_trip_simple_sheet() {
        let mut sheet = Sheet {
            name: "Data".into(),
            ..Default::default()
        };
        let mut row1 = BTreeMap::new();
        row1.insert(
            1,
            Cell {
                value: CellValue::Text("Hello".into()),
                ..Default::default()
            },
        );
        row1.insert(
            2,
            Cell {
                value: CellValue::Number(42.5),
                ..Default::default()
            },
        );
        sheet.cells.insert(1, row1);

        let mut row2 = BTreeMap::new();
        row2.insert(
            1,
            Cell {
                value: CellValue::Boolean(true),
                ..Default::default()
            },
        );
        sheet.cells.insert(2, row2);

        let doc = CalcDocument {
            sheets: vec![sheet],
            ..Default::default()
        };

        let bytes = write_xlsx(&doc).unwrap();
        let parsed = read_xlsx(&bytes).unwrap();

        assert_eq!(parsed.sheets.len(), 1);
        let s = &parsed.sheets[0];
        assert_eq!(s.name, "Data");

        let a1 = &s.cells[&1][&1];
        assert!(matches!(&a1.value, CellValue::Text(t) if t == "Hello"));

        let b1 = &s.cells[&1][&2];
        assert!(matches!(&b1.value, CellValue::Number(n) if (*n - 42.5).abs() < 1e-9));

        let a2 = &s.cells[&2][&1];
        assert!(matches!(a2.value, CellValue::Boolean(true)));
    }

    #[test]
    fn round_trip_merges_and_hyperlinks() {
        let sheet = Sheet {
            name: "Hub".into(),
            merges: vec![MergedRange {
                row_start: 1,
                col_start: 1,
                row_end: 1,
                col_end: 3,
            }],
            hyperlinks: vec![SheetHyperlink {
                cell_ref: "A1".into(),
                target: "https://example.com/".into(),
                display: Some("Example".into()),
            }],
            ..Default::default()
        };
        let doc = CalcDocument {
            sheets: vec![sheet],
            ..Default::default()
        };
        let bytes = write_xlsx(&doc).unwrap();
        let parsed = read_xlsx(&bytes).unwrap();
        let s = &parsed.sheets[0];
        assert_eq!(s.merges.len(), 1);
        assert_eq!(s.merges[0].col_end, 3);
        assert_eq!(s.hyperlinks.len(), 1);
        assert_eq!(s.hyperlinks[0].target, "https://example.com/");
    }
}