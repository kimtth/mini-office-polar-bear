//! PPTX format support — import and export between `.pptx` files and `SlidesDocument`.

mod reader;
mod render;
mod writer;

pub use reader::read_pptx;
pub use render::{read_pptx_rendered, render_document};
pub use writer::write_pptx;

#[cfg(test)]
mod tests {
    use super::*;
    use pbo_common::model::slides::{
        Shape, ShapeType, Slide, SlideElement, SlidesDocument, TextBox,
    };

    #[test]
    fn default_slides_document_is_empty() {
        let doc = SlidesDocument::default();
        assert!(doc.slides.is_empty());
    }

    #[test]
    fn round_trip_simple_slide() {
        let slide = Slide {
            elements: vec![SlideElement::TextBox(TextBox {
                text: "Hello\nWorld".into(),
                x: 914400.0,
                y: 914400.0,
                width: 4572000.0,
                height: 1828800.0,
                font_size: Some(18.0),
                bold: true,
                ..Default::default()
            })],
            ..Default::default()
        };

        let doc = SlidesDocument {
            slides: vec![slide],
            slide_width: 9144000.0,
            slide_height: 6858000.0,
            ..Default::default()
        };

        let bytes = write_pptx(&doc).unwrap();
        let parsed = read_pptx(&bytes).unwrap();

        assert_eq!(parsed.slides.len(), 1);
        assert_eq!(parsed.slide_width as i64, 9144000);
        let SlideElement::TextBox(tb) = &parsed.slides[0].elements[0] else {
            panic!("expected text box");
        };
        assert_eq!(tb.text, "Hello\nWorld");
        assert_eq!(tb.x as i64, 914400);
        assert_eq!(tb.width as i64, 4572000);
    }

    #[test]
    fn round_trip_preset_shape() {
        let slide = Slide {
            elements: vec![SlideElement::Shape(Shape {
                shape_type: ShapeType::Ellipse,
                x: 1_000_000.0,
                y: 2_000_000.0,
                width: 3_000_000.0,
                height: 1_500_000.0,
                rotation: 0,
                fill: None,
                outline: None,
                text: None,
                custom_geometry: None,
                ..Default::default()
            })],
            ..Default::default()
        };
        let doc = SlidesDocument {
            slides: vec![slide],
            slide_width: 9144000.0,
            slide_height: 6858000.0,
            ..Default::default()
        };
        let bytes = write_pptx(&doc).unwrap();
        let parsed = read_pptx(&bytes).unwrap();
        let SlideElement::Shape(sh) = &parsed.slides[0].elements[0] else {
            panic!("expected shape");
        };
        assert_eq!(sh.shape_type, ShapeType::Ellipse);
        assert_eq!(sh.width as i64, 3_000_000);
    }

    #[test]
    fn round_trip_shape_style_and_background() {
        let slide = Slide {
            background: Some(pbo_common::model::slides::Fill::Solid("112233".into())),
            elements: vec![SlideElement::Shape(Shape {
                shape_type: ShapeType::RoundedRectangle,
                x: 100_000.0,
                y: 200_000.0,
                width: 1_000_000.0,
                height: 800_000.0,
                rotation: 5_400_000,
                fill: Some(pbo_common::model::slides::Fill::Solid("FFAA00".into())),
                outline: Some(pbo_common::model::slides::Outline {
                    width_emu: 12_700,
                    color: Some("003366".into()),
                    dash: Some("dash".into()),
                    head_end: None,
                    tail_end: None,
                    ..Default::default()
                }),
                text: Some(TextBox {
                    text: "Styled".into(),
                    width: 1_000_000.0,
                    height: 800_000.0,
                    color: Some("FFFFFF".into()),
                    ..Default::default()
                }),
                flip_h: true,
                ..Default::default()
            })],
            ..Default::default()
        };
        let doc = SlidesDocument {
            slides: vec![slide],
            slide_width: 9144000.0,
            slide_height: 6858000.0,
            ..Default::default()
        };

        let bytes = write_pptx(&doc).unwrap();
        let parsed = read_pptx(&bytes).unwrap();

        assert!(parsed.slides[0].background.is_some());
        let SlideElement::Shape(shape) = &parsed.slides[0].elements[0] else {
            panic!("expected styled shape");
        };
        assert!(shape.fill.is_some());
        assert_eq!(shape.rotation, 5_400_000);
        assert!(shape.flip_h);
        assert_eq!(
            shape.outline.as_ref().and_then(|o| o.color.as_deref()),
            Some("003366")
        );
        assert_eq!(
            shape.text.as_ref().map(|text| text.text.as_str()),
            Some("Styled")
        );
    }

    #[test]
    fn round_trip_table() {
        use pbo_common::model::slides::{SlideTable, SlideTableCell, SlideTableRow};

        let slide = Slide {
            elements: vec![SlideElement::Table(SlideTable {
                x: 100_000.0,
                y: 100_000.0,
                width: 2_000_000.0,
                height: 1_000_000.0,
                columns: vec![1_000_000.0, 1_000_000.0],
                rows: vec![SlideTableRow {
                    height: 500_000.0,
                    cells: vec![
                        SlideTableCell {
                            text: "A".into(),
                            col_span: 1,
                            row_span: 1,
                            fill: Some(pbo_common::model::slides::Fill::Solid("EEEEEE".into())),
                        },
                        SlideTableCell {
                            text: "B".into(),
                            col_span: 1,
                            row_span: 1,
                            fill: None,
                        },
                    ],
                }],
            })],
            ..Default::default()
        };
        let doc = SlidesDocument {
            slides: vec![slide],
            slide_width: 9144000.0,
            slide_height: 6858000.0,
            ..Default::default()
        };

        let bytes = write_pptx(&doc).unwrap();
        let parsed = read_pptx(&bytes).unwrap();

        let SlideElement::Table(table) = &parsed.slides[0].elements[0] else {
            panic!("expected table");
        };
        assert_eq!(table.rows[0].cells[0].text, "A");
        assert_eq!(table.columns.len(), 2);
    }

    #[test]
    fn render_document_materializes_layout_chrome() {
        use pbo_common::model::slides::{Fill, SlideLayout, SlideMaster};

        let doc = SlidesDocument {
            slides: vec![Slide {
                elements: vec![SlideElement::TextBox(TextBox {
                    text: "Title".into(),
                    placeholder: Some("title".into()),
                    ..Default::default()
                })],
                layout_idx: Some(0),
                ..Default::default()
            }],
            layouts: vec![SlideLayout {
                master_idx: Some(0),
                elements: vec![SlideElement::TextBox(TextBox {
                    placeholder: Some("title".into()),
                    x: 10.0,
                    y: 20.0,
                    width: 30.0,
                    height: 40.0,
                    ..Default::default()
                })],
                ..Default::default()
            }],
            masters: vec![SlideMaster {
                elements: vec![SlideElement::Shape(Shape {
                    fill: Some(Fill::Solid("000000".into())),
                    width: 100.0,
                    height: 100.0,
                    ..Default::default()
                })],
                ..Default::default()
            }],
            ..Default::default()
        };

        let rendered = render_document(&doc);

        assert_eq!(rendered.slides[0].elements.len(), 2);
        let SlideElement::TextBox(text) = &rendered.slides[0].elements[1] else {
            panic!("expected inherited placeholder geometry");
        };
        assert_eq!(text.width as i64, 30);
    }
}
