//! DOCX format support — import and export between `.docx` files and `WriterDocument`.

mod reader;
mod writer;

pub use reader::read_docx;
pub use writer::write_docx;

#[cfg(test)]
mod tests {
    use super::*;
    use pbo_common::model::writer::{
        Alignment, Block, Paragraph, ParagraphStyle, Run, RunElement, RunStyle, Table, TableCell,
        TableRow, WriterDocument,
    };

    #[test]
    fn default_writer_document_is_empty() {
        let doc = WriterDocument::default();
        assert!(doc.body.is_empty());
    }

    #[test]
    fn round_trip_simple_paragraph() {
        let mut doc = WriterDocument::default();
        doc.body.push(Block::Paragraph(Paragraph {
            runs: vec![RunElement::Run(Run {
                text: "Hello, world!".into(),
                style: Some(RunStyle {
                    bold: true,
                    ..Default::default()
                }),
            })],
            style: Some(ParagraphStyle {
                alignment: Some(Alignment::Center),
                heading_level: Some(1),
                ..Default::default()
            }),
            bookmarks: vec![],
        }));

        let bytes = write_docx(&doc).unwrap();
        let parsed = read_docx(&bytes).unwrap();

        assert_eq!(parsed.body.len(), 1);
        let Block::Paragraph(p) = &parsed.body[0] else {
            panic!("expected paragraph");
        };
        assert_eq!(p.runs.len(), 1);
        let RunElement::Run(r) = &p.runs[0] else {
            panic!("expected Run")
        };
        assert_eq!(r.text, "Hello, world!");
        assert!(r.style.as_ref().unwrap().bold);
        assert!(matches!(
            p.style.as_ref().unwrap().alignment,
            Some(Alignment::Center)
        ));
        assert_eq!(p.style.as_ref().unwrap().heading_level, Some(1));
    }

    #[test]
    fn round_trip_table_with_merged_cells() {
        let cell = |text: &str, span: u32| TableCell {
            content: vec![Block::Paragraph(Paragraph {
                runs: vec![RunElement::Run(Run {
                    text: text.into(),
                    style: None,
                })],
                ..Default::default()
            })],
            width: 2000,
            grid_span: span,
            ..Default::default()
        };
        let table = Table {
            grid_cols: vec![2000, 2000, 2000],
            rows: vec![
                TableRow {
                    cells: vec![cell("A", 1), cell("B", 1), cell("C", 1)],
                    header: true,
                    ..Default::default()
                },
                TableRow {
                    cells: vec![cell("Spans two", 2), cell("Third", 1)],
                    ..Default::default()
                },
            ],
            style: None,
        };
        let mut doc = WriterDocument::default();
        doc.body.push(Block::Table(table));

        let bytes = write_docx(&doc).unwrap();
        let parsed = read_docx(&bytes).unwrap();

        let Block::Table(t) = &parsed.body[0] else {
            panic!("expected table");
        };
        assert_eq!(t.grid_cols, vec![2000, 2000, 2000]);
        assert_eq!(t.rows.len(), 2);
        assert!(t.rows[0].header);
        assert_eq!(t.rows[0].cells.len(), 3);
        assert_eq!(t.rows[1].cells[0].grid_span, 2);
    }
}