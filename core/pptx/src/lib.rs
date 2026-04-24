//! PPTX format support — import and export between `.pptx` files and `SlidesDocument`.

mod reader;
mod writer;

pub use reader::read_pptx;
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
}