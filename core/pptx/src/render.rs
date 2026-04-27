//! PPTX rendering helpers.
//!
//! These helpers build a display-oriented `SlidesDocument` from the canonical
//! import model. The editing model still stays in `pbo-common`; this module only
//! resolves slide master/layout inheritance so clients can render a slide without
//! re-implementing PresentationML placeholder lookup rules.

use pbo_common::model::slides::{
    Shape, Slide, SlideElement, SlidesDocument, TextBodyProperties, TextBox, TextRun,
};

/// Parse a `.pptx` file and return a render-ready document with inherited
/// layout/master elements materialized into each slide.
pub fn read_pptx_rendered(data: &[u8]) -> pbo_common::PboResult<SlidesDocument> {
    let doc = crate::read_pptx(data)?;
    Ok(render_document(&doc))
}

/// Return a render-ready copy of `doc`.
///
/// Each slide receives non-placeholder master/layout chrome first, followed by
/// slide-local elements. Slide-local placeholders with missing geometry inherit
/// geometry from matching layout/master placeholders.
pub fn render_document(doc: &SlidesDocument) -> SlidesDocument {
    let mut rendered = doc.clone();
    rendered.slides = doc
        .slides
        .iter()
        .map(|slide| render_slide(doc, slide))
        .collect();
    rendered.layouts.clear();
    rendered.masters.clear();
    rendered
}

fn render_slide(doc: &SlidesDocument, slide: &Slide) -> Slide {
    let layout = slide.layout_idx.and_then(|idx| doc.layouts.get(idx));
    let master = layout
        .and_then(|layout| layout.master_idx)
        .and_then(|idx| doc.masters.get(idx));

    let master_elements = master
        .map(|master| non_placeholder_elements(&master.elements))
        .unwrap_or_default();
    let layout_elements = layout
        .map(|layout| non_placeholder_elements(&layout.elements))
        .unwrap_or_default();
    let slide_elements = slide
        .elements
        .iter()
        .map(|element| inherit_placeholder(element, doc, layout, master))
        .collect::<Vec<_>>();

    let mut elements =
        Vec::with_capacity(master_elements.len() + layout_elements.len() + slide_elements.len());
    elements.extend(master_elements);
    elements.extend(layout_elements);
    elements.extend(slide_elements);

    Slide {
        elements,
        notes: slide.notes.clone(),
        layout_idx: None,
        background: slide.background.clone(),
    }
}

fn non_placeholder_elements(elements: &[SlideElement]) -> Vec<SlideElement> {
    elements
        .iter()
        .filter(|element| placeholder_key(element).is_none())
        .cloned()
        .collect()
}

fn inherit_placeholder(
    element: &SlideElement,
    doc: &SlidesDocument,
    layout: Option<&pbo_common::model::slides::SlideLayout>,
    master: Option<&pbo_common::model::slides::SlideMaster>,
) -> SlideElement {
    match element {
        SlideElement::TextBox(text_box) if text_box.placeholder.is_some() => {
            let Some(key) = placeholder_key(element) else {
                return element.clone();
            };
            let mut text_box = text_box.clone();
            if has_empty_box(&text_box)
                && let Some((x, y, width, height)) =
                    find_placeholder_geometry_chain(doc, layout, master, &key)
            {
                text_box.x = x;
                text_box.y = y;
                text_box.width = width;
                text_box.height = height;
            }
            if let Some(source) = find_placeholder_text_chain(doc, layout, master, &key) {
                merge_text_style(&mut text_box, source);
            }
            apply_placeholder_fallback_style(&mut text_box);
            SlideElement::TextBox(text_box)
        }
        SlideElement::Shape(shape)
            if shape
                .text
                .as_ref()
                .and_then(|text| text.placeholder.as_ref())
                .is_some() =>
        {
            let Some(key) = placeholder_key(element) else {
                return element.clone();
            };
            let mut shape = shape.clone();
            if has_empty_shape(&shape)
                && let Some((x, y, width, height)) =
                    find_placeholder_geometry_chain(doc, layout, master, &key)
            {
                shape.x = x;
                shape.y = y;
                shape.width = width;
                shape.height = height;
            }
            if let (Some(target), Some(source)) = (
                shape.text.as_mut(),
                find_placeholder_text_chain(doc, layout, master, &key),
            ) {
                merge_text_style(target, source);
            }
            if let Some(target) = shape.text.as_mut() {
                apply_placeholder_fallback_style(target);
            }
            SlideElement::Shape(shape)
        }
        _ => element.clone(),
    }
}

fn find_placeholder_geometry_chain(
    doc: &SlidesDocument,
    layout: Option<&pbo_common::model::slides::SlideLayout>,
    master: Option<&pbo_common::model::slides::SlideMaster>,
    key: &str,
) -> Option<(f64, f64, f64, f64)> {
    layout
        .and_then(|layout| find_placeholder_geometry(&layout.elements, key))
        .or_else(|| master.and_then(|master| find_placeholder_geometry(&master.elements, key)))
        .or_else(|| find_placeholder_geometry_in_slides(doc, key))
}

fn find_placeholder_text_chain<'a>(
    doc: &'a SlidesDocument,
    layout: Option<&'a pbo_common::model::slides::SlideLayout>,
    master: Option<&'a pbo_common::model::slides::SlideMaster>,
    key: &str,
) -> Option<&'a TextBox> {
    layout
        .and_then(|layout| find_placeholder_text(&layout.elements, key))
        .or_else(|| master.and_then(|master| find_placeholder_text(&master.elements, key)))
        .or_else(|| {
            doc.slides
                .iter()
                .find_map(|slide| find_placeholder_text(&slide.elements, key))
        })
}

fn merge_text_style(target: &mut TextBox, source: &TextBox) {
    if target.font_size.is_none() {
        target.font_size = source.font_size;
    }
    if target.color.is_none() {
        target.color = source.color.clone();
    }
    if target.align.is_none() {
        target.align = source.align;
    }
    merge_body_style(&mut target.body, &source.body);

    let source_run = source
        .paragraphs
        .iter()
        .find_map(|paragraph| paragraph.runs.first());
    for paragraph in &mut target.paragraphs {
        if paragraph.align.is_none() {
            paragraph.align = source
                .paragraphs
                .first()
                .and_then(|paragraph| paragraph.align);
        }
        if paragraph.margin_left.is_none() {
            paragraph.margin_left = source
                .paragraphs
                .first()
                .and_then(|paragraph| paragraph.margin_left);
        }
        if paragraph.indent.is_none() {
            paragraph.indent = source
                .paragraphs
                .first()
                .and_then(|paragraph| paragraph.indent);
        }
        for run in &mut paragraph.runs {
            merge_run_style(run, source_run, source);
        }
    }
}

fn apply_placeholder_fallback_style(text_box: &mut TextBox) {
    let (font_size, color, family) = match (
        text_box.placeholder.as_deref(),
        text_box.placeholder_idx.as_deref(),
    ) {
        // Common PowerPoint/Office title placeholder used by the Fabric deck's
        // layout. The corresponding layout placeholder can be omitted by older
        // simplified parsers, so keep a conservative render fallback.
        (Some("body"), Some("23")) => (Some(28.0), Some("001919"), Some("メイリオ")),
        (Some("title" | "ctrTitle"), _) => (Some(28.0), Some("001919"), Some("メイリオ")),
        (Some("subTitle"), _) => (Some(18.0), Some("0070C0"), Some("メイリオ")),
        _ => (None, None, None),
    };
    if text_box.font_size.is_none() {
        text_box.font_size = font_size;
    }
    if text_box.color.is_none() {
        text_box.color = color.map(str::to_owned);
    }
    for paragraph in &mut text_box.paragraphs {
        for run in &mut paragraph.runs {
            if run.font_size.is_none() {
                run.font_size = font_size;
            }
            if run.color.is_none() {
                run.color = color.map(str::to_owned);
            }
            if run.font_family.is_none() {
                run.font_family = family.map(str::to_owned);
            }
        }
    }
}

fn merge_body_style(target: &mut TextBodyProperties, source: &TextBodyProperties) {
    if target.inset_left.is_none() {
        target.inset_left = source.inset_left;
    }
    if target.inset_top.is_none() {
        target.inset_top = source.inset_top;
    }
    if target.inset_right.is_none() {
        target.inset_right = source.inset_right;
    }
    if target.inset_bottom.is_none() {
        target.inset_bottom = source.inset_bottom;
    }
    if target.anchor.is_none() {
        target.anchor = source.anchor.clone();
    }
    if target.rotation.is_none() {
        target.rotation = source.rotation;
    }
    if target.vertical.is_none() {
        target.vertical = source.vertical.clone();
    }
    if target.wrap.is_none() {
        target.wrap = source.wrap.clone();
    }
    if target.vertical_overflow.is_none() {
        target.vertical_overflow = source.vertical_overflow.clone();
    }
    if target.horizontal_overflow.is_none() {
        target.horizontal_overflow = source.horizontal_overflow.clone();
    }
}

fn merge_run_style(run: &mut TextRun, source_run: Option<&TextRun>, source_box: &TextBox) {
    if run.font_size.is_none() {
        run.font_size = source_run
            .and_then(|run| run.font_size)
            .or(source_box.font_size);
    }
    if run.color.is_none() {
        run.color = source_run
            .and_then(|run| run.color.clone())
            .or_else(|| source_box.color.clone());
    }
    if run.font_family.is_none() {
        run.font_family = source_run.and_then(|run| run.font_family.clone());
    }
    if !run.bold {
        run.bold = source_run.is_some_and(|run| run.bold) || source_box.bold;
    }
    if !run.italic {
        run.italic = source_run.is_some_and(|run| run.italic) || source_box.italic;
    }
    if !run.underline {
        run.underline = source_run.is_some_and(|run| run.underline);
    }
    if !run.strike {
        run.strike = source_run.is_some_and(|run| run.strike);
    }
    if run.baseline.is_none() {
        run.baseline = source_run.and_then(|run| run.baseline);
    }
}

fn has_empty_box(text_box: &TextBox) -> bool {
    text_box.width <= 0.0 || text_box.height <= 0.0
}

fn has_empty_shape(shape: &Shape) -> bool {
    shape.width <= 0.0 || shape.height <= 0.0
}

fn placeholder_key(element: &SlideElement) -> Option<String> {
    match element {
        SlideElement::TextBox(text_box) => text_box
            .placeholder
            .as_deref()
            .map(|kind| ph_key(kind, text_box.placeholder_idx.as_deref())),
        SlideElement::Shape(shape) => shape.text.as_ref().and_then(|text| {
            text.placeholder
                .as_deref()
                .map(|kind| ph_key(kind, text.placeholder_idx.as_deref()))
        }),
        _ => None,
    }
}

fn ph_key(kind: &str, idx: Option<&str>) -> String {
    let kind = if kind == "ctrTitle" { "title" } else { kind };
    match idx {
        Some(idx) if !idx.is_empty() => format!("{kind}|{idx}"),
        _ => kind.to_owned(),
    }
}

fn find_placeholder_text<'a>(elements: &'a [SlideElement], key: &str) -> Option<&'a TextBox> {
    elements.iter().find_map(|element| {
        if placeholder_key(element).as_deref() != Some(key) {
            return None;
        }
        match element {
            SlideElement::TextBox(text_box) => Some(text_box),
            SlideElement::Shape(shape) => shape.text.as_ref(),
            _ => None,
        }
    })
}

fn find_placeholder_geometry(elements: &[SlideElement], key: &str) -> Option<(f64, f64, f64, f64)> {
    elements.iter().find_map(|element| {
        if placeholder_key(element).as_deref() != Some(key) {
            return None;
        }
        match element {
            SlideElement::TextBox(text_box) if text_box.width > 0.0 && text_box.height > 0.0 => {
                Some((text_box.x, text_box.y, text_box.width, text_box.height))
            }
            SlideElement::Shape(shape) if shape.width > 0.0 && shape.height > 0.0 => {
                Some((shape.x, shape.y, shape.width, shape.height))
            }
            _ => None,
        }
    })
}

fn find_placeholder_geometry_in_slides(
    doc: &SlidesDocument,
    key: &str,
) -> Option<(f64, f64, f64, f64)> {
    doc.slides
        .iter()
        .find_map(|slide| find_placeholder_geometry(&slide.elements, key))
}
