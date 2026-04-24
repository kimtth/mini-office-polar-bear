// PPTX reader: parses a .pptx OPC package into a SlidesDocument.
// Covers slides with text boxes, preset shapes, pictures, connectors,
// groups; notes slides; core metadata.

use pbo_common::b64;
use pbo_common::model::slides::{
    Connector, ConnectorType, Fill, GradientStop, GroupShape, Outline, Picture,
    PresentationMetadata, Shape, ShapeType, Slide, SlideElement, SlideLayout, SlideMaster,
    SlidesDocument, SlidesImage, TextAlign, TextBox, TextParagraph, TextRun, Theme, ThemeColor,
};
use pbo_common::{PboError, PboResult};
use pbo_ooxml::xmlio::{attr_i64, attr_local, attr_qualified, local_name};
use pbo_ooxml::{OpcPackage, Relationships};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::HashMap;

const PRESENTATION_PART: &str = "ppt/presentation.xml";
const PRESENTATION_RELS: &str = "ppt/_rels/presentation.xml.rels";
const CORE_PART: &str = "docProps/core.xml";
const REL_SLIDE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_NOTES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_LAYOUT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const REL_MASTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const REL_THEME: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

pub fn read_pptx(data: &[u8]) -> PboResult<SlidesDocument> {
    let package = OpcPackage::from_bytes(data)?;
    let mut doc = SlidesDocument::default();

    if let Ok(xml) = package.get_part_str(CORE_PART) {
        doc.metadata = parse_core(xml)?;
    }

    let pres_xml = package.get_part_str(PRESENTATION_PART)?;
    let (w, h) = parse_slide_size(pres_xml)?;
    doc.slide_width = w;
    doc.slide_height = h;
    let slide_rids = parse_slide_id_list(pres_xml)?;

    let pres_rels = package
        .get_part_str(PRESENTATION_RELS)
        .ok()
        .map(Relationships::parse)
        .transpose()?
        .unwrap_or_default();
    let mut layout_indices = HashMap::new();
    let mut master_indices = HashMap::new();

    for rid in slide_rids {
        let Some(rel) = pres_rels.find_by_id(&rid) else {
            continue;
        };
        if rel.rel_type != REL_SLIDE {
            continue;
        }
        let slide_path = normalize(&rel.target, "ppt/");
        let Ok(slide_xml) = package.get_part_str(&slide_path) else {
            continue;
        };
        let slide_rels = package
            .get_part_str(&rels_path(&slide_path))
            .ok()
            .map(Relationships::parse)
            .transpose()?
            .unwrap_or_default();
        let layout_idx = if let Some(layout_path) = slide_rels
            .0
            .iter()
            .find(|rel| rel.rel_type == REL_LAYOUT)
            .map(|rel| normalize(&rel.target, &slide_path))
        {
            ensure_layout(
                &layout_path,
                &package,
                &mut doc,
                &mut layout_indices,
                &mut master_indices,
            )?
        } else {
            None
        };
        let theme = doc.theme.clone();
        let (mut slide, notes_path) = parse_slide(
            slide_xml,
            &slide_rels,
            &slide_path,
            &package,
            &mut doc.images,
            theme.as_ref(),
        )?;
        slide.layout_idx = layout_idx;
        if let Some(np) = notes_path
            && let Ok(nxml) = package.get_part_str(&np)
        {
            slide.notes = Some(extract_notes_text(nxml)?);
        }
        doc.slides.push(slide);
    }
    Ok(doc)
}

fn ensure_layout(
    layout_path: &str,
    package: &OpcPackage,
    doc: &mut SlidesDocument,
    layout_indices: &mut HashMap<String, usize>,
    master_indices: &mut HashMap<String, usize>,
) -> PboResult<Option<usize>> {
    if let Some(idx) = layout_indices.get(layout_path).copied() {
        return Ok(Some(idx));
    }

    let Ok(layout_xml) = package.get_part_str(layout_path) else {
        return Ok(None);
    };
    let layout_rels = package
        .get_part_str(&rels_path(layout_path))
        .ok()
        .map(Relationships::parse)
        .transpose()?
        .unwrap_or_default();
    let master_path = layout_rels
        .0
        .iter()
        .find(|rel| rel.rel_type == REL_MASTER)
        .map(|rel| normalize(&rel.target, layout_path));
    let theme = doc.theme.clone();
    let mut layout = parse_layout(
        layout_xml,
        &layout_rels,
        layout_path,
        package,
        &mut doc.images,
        theme.as_ref(),
    )?;
    layout.master_idx = if let Some(master_path) = master_path {
        ensure_master(&master_path, package, doc, master_indices)?
    } else {
        None
    };

    let idx = doc.layouts.len();
    doc.layouts.push(layout);
    layout_indices.insert(layout_path.to_owned(), idx);
    Ok(Some(idx))
}

fn ensure_master(
    master_path: &str,
    package: &OpcPackage,
    doc: &mut SlidesDocument,
    master_indices: &mut HashMap<String, usize>,
) -> PboResult<Option<usize>> {
    if let Some(idx) = master_indices.get(master_path).copied() {
        return Ok(Some(idx));
    }

    let Ok(master_xml) = package.get_part_str(master_path) else {
        return Ok(None);
    };
    let master_rels = package
        .get_part_str(&rels_path(master_path))
        .ok()
        .map(Relationships::parse)
        .transpose()?
        .unwrap_or_default();
    if doc.theme.is_none()
        && let Some(theme_path) = master_rels
            .0
            .iter()
            .find(|rel| rel.rel_type == REL_THEME)
            .map(|rel| normalize(&rel.target, master_path))
        && let Ok(theme_xml) = package.get_part_str(&theme_path)
    {
        doc.theme = Some(parse_theme(theme_xml)?);
    }
    let theme = doc.theme.clone();
    let master = parse_master(
        master_xml,
        &master_rels,
        master_path,
        package,
        &mut doc.images,
        theme.as_ref(),
    )?;

    let idx = doc.masters.len();
    doc.masters.push(master);
    master_indices.insert(master_path.to_owned(), idx);
    Ok(Some(idx))
}

#[cfg(any())]
mod duplicate_reader_helpers {
    use super::*;

    fn parse_theme(xml: &str) -> PboResult<Theme> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let mut theme = Theme::default();
        let mut current_role: Option<String> = None;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) | Event::Empty(e) => {
                    let name = local_name(e.name().as_ref());
                    if name == b"theme" {
                        if let Some(value) = attr_local(&e, b"name", &reader)? {
                            theme.name = value;
                        }
                    } else if is_theme_color_role(name) {
                        current_role = Some(String::from_utf8_lossy(name).into_owned());
                    } else if name == b"srgbClr" {
                        if let (Some(role), Some(value)) =
                            (current_role.clone(), attr_local(&e, b"val", &reader)?)
                        {
                            theme.colors.push(ThemeColor { role, value });
                        }
                    } else if name == b"sysClr"
                        && let (Some(role), Some(value)) =
                            (current_role.clone(), attr_local(&e, b"lastClr", &reader)?)
                    {
                        theme.colors.push(ThemeColor { role, value });
                    }
                }
                Event::End(e) if is_theme_color_role(local_name(e.name().as_ref())) => {
                    current_role = None;
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(theme)
    }

    fn is_theme_color_role(name: &[u8]) -> bool {
        matches!(
            name,
            b"dk1"
                | b"lt1"
                | b"dk2"
                | b"lt2"
                | b"accent1"
                | b"accent2"
                | b"accent3"
                | b"accent4"
                | b"accent5"
                | b"accent6"
                | b"hlink"
                | b"folHlink"
        )
    }

    fn resolve_theme_color(theme: Option<&Theme>, role: &str) -> Option<String> {
        let theme = theme?;
        let normalized_role = match role {
            "bg1" => "lt1",
            "tx1" => "dk1",
            "bg2" => "lt2",
            "tx2" => "dk2",
            other => other,
        };
        theme
            .colors
            .iter()
            .find(|color| color.role == normalized_role)
            .map(|color| color.value.clone())
    }

    /// Apply color modifiers (lumMod, lumOff, tint, shade, alpha) to an RGB hex string.
    fn apply_color_modifiers(
        hex: &str,
        lum_mod: Option<i64>,
        lum_off: Option<i64>,
        tint: Option<i64>,
    ) -> String {
        let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64;
        let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64;
        let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64;

        let (mut r, mut g, mut b) = (r, g, b);

        // Apply tint: tint towards white
        if let Some(t) = tint {
            let factor = t as f64 / 100000.0;
            r = r + (255.0 - r) * factor;
            g = g + (255.0 - g) * factor;
            b = b + (255.0 - b) * factor;
        }

        // Apply lumMod/lumOff: modulate luminance
        if lum_mod.is_some() || lum_off.is_some() {
            let mod_val = lum_mod.unwrap_or(100000) as f64 / 100000.0;
            let off_val = lum_off.unwrap_or(0) as f64 / 100000.0 * 255.0;
            r = (r * mod_val + off_val).clamp(0.0, 255.0);
            g = (g * mod_val + off_val).clamp(0.0, 255.0);
            b = (b * mod_val + off_val).clamp(0.0, 255.0);
        }

        format!("{:02X}{:02X}{:02X}", r as u8, g as u8, b as u8)
    }

    /// Resolve a color element (srgbClr or schemeClr with modifiers) reading child events.
    fn resolve_color_with_modifiers(
        reader: &mut Reader<&[u8]>,
        start_tag: &[u8],
        base_color: &str,
        theme: Option<&Theme>,
    ) -> PboResult<Option<String>> {
        let is_scheme = start_tag == b"schemeClr";
        let base = if is_scheme {
            resolve_theme_color(theme, base_color)
        } else {
            Some(base_color.to_string())
        };
        let Some(hex) = base else { return Ok(None) };

        let mut lum_mod = None;
        let mut lum_off = None;
        let mut tint = None;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"lumMod" => {
                        lum_mod = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok())
                    }
                    b"lumOff" => {
                        lum_off = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok())
                    }
                    b"tint" => tint = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
                    b"shade" => {
                        // shade is inverse of tint — darken
                        if let Some(v) =
                            attr_local(&e, b"val", reader)?.and_then(|v| v.parse::<i64>().ok())
                        {
                            lum_mod = Some(v);
                        }
                    }
                    _ => {}
                },
                Event::End(e) if local_name(e.name().as_ref()) == start_tag => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        if lum_mod.is_some() || lum_off.is_some() || tint.is_some() {
            Ok(Some(apply_color_modifiers(&hex, lum_mod, lum_off, tint)))
        } else {
            Ok(Some(hex))
        }
    }

    /// Parse `<p:style>` to extract fill and outline references from the theme style matrix.
    fn parse_style_refs(
        reader: &mut Reader<&[u8]>,
        theme: Option<&Theme>,
        style_fill: &mut Option<Fill>,
        style_outline: &mut Option<Outline>,
    ) -> PboResult<()> {
        let mut buf = Vec::new();
        let mut in_fill_ref = false;
        let mut in_ln_ref = false;
        let mut fill_ref_idx: Option<u32> = None;
        let mut ln_ref_idx: Option<u32> = None;
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"fillRef" => {
                        fill_ref_idx = attr_local(&e, b"idx", reader)?.and_then(|v| v.parse().ok());
                        in_fill_ref = true;
                    }
                    b"lnRef" => {
                        ln_ref_idx = attr_local(&e, b"idx", reader)?.and_then(|v| v.parse().ok());
                        in_ln_ref = true;
                    }
                    _ => {}
                },
                Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"schemeClr" => {
                        if let Some(val) = attr_local(&e, b"val", reader)?
                            && let Some(resolved) = resolve_theme_color(theme, &val)
                        {
                            if in_fill_ref && fill_ref_idx.is_some_and(|i| i > 0) {
                                *style_fill = Some(Fill::Solid(resolved.clone()));
                            }
                            if in_ln_ref && ln_ref_idx.is_some_and(|i| i > 0) {
                                *style_outline = Some(Outline {
                                    width_emu: 12700,
                                    color: Some(resolved),
                                    dash: None,
                                    head_end: None,
                                    tail_end: None,
                                });
                            }
                        }
                    }
                    b"srgbClr" => {
                        if let Some(val) = attr_local(&e, b"val", reader)? {
                            if in_fill_ref && fill_ref_idx.is_some_and(|i| i > 0) {
                                *style_fill = Some(Fill::Solid(val.clone()));
                            }
                            if in_ln_ref && ln_ref_idx.is_some_and(|i| i > 0) {
                                *style_outline = Some(Outline {
                                    width_emu: 12700,
                                    color: Some(val),
                                    dash: None,
                                    head_end: None,
                                    tail_end: None,
                                });
                            }
                        }
                    }
                    _ => {}
                },
                Event::End(e) => match local_name(e.name().as_ref()) {
                    b"fillRef" => in_fill_ref = false,
                    b"lnRef" => in_ln_ref = false,
                    b"style" => break,
                    _ => {}
                },
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(())
    }

    fn parse_gradient_fill(
        reader: &mut Reader<&[u8]>,
        theme: Option<&Theme>,
    ) -> PboResult<Option<Fill>> {
        let mut stops = Vec::new();
        let mut current_pos = None;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"gs" => {
                        current_pos =
                            attr_local(&e, b"pos", reader)?.and_then(|value| value.parse().ok());
                    }
                    b"srgbClr" => {
                        if let Some(pos) = current_pos
                            && let Some(val) = attr_local(&e, b"val", reader)?
                            && let Some(color) =
                                resolve_color_with_modifiers(reader, b"srgbClr", &val, theme)?
                        {
                            stops.push(GradientStop {
                                position: pos,
                                color,
                            });
                        }
                    }
                    b"schemeClr" => {
                        if let Some(pos) = current_pos
                            && let Some(role) = attr_local(&e, b"val", reader)?
                            && let Some(color) =
                                resolve_color_with_modifiers(reader, b"schemeClr", &role, theme)?
                        {
                            stops.push(GradientStop {
                                position: pos,
                                color,
                            });
                        }
                    }
                    _ => {}
                },
                Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"srgbClr" => {
                        if let (Some(pos), Some(color)) =
                            (current_pos, attr_local(&e, b"val", reader)?)
                        {
                            stops.push(GradientStop {
                                position: pos,
                                color,
                            });
                        }
                    }
                    b"schemeClr" => {
                        if let (Some(pos), Some(role)) =
                            (current_pos, attr_local(&e, b"val", reader)?)
                            && let Some(color) = resolve_theme_color(theme, &role)
                        {
                            stops.push(GradientStop {
                                position: pos,
                                color,
                            });
                        }
                    }
                    _ => {}
                },
                Event::End(e) if local_name(e.name().as_ref()) == b"gs" => current_pos = None,
                Event::End(e) if local_name(e.name().as_ref()) == b"gradFill" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        if stops.is_empty() {
            Ok(None)
        } else {
            Ok(Some(Fill::Gradient(stops)))
        }
    }

    fn parse_outline_empty(
        e: &quick_xml::events::BytesStart<'_>,
        reader: &Reader<&[u8]>,
    ) -> PboResult<Outline> {
        let width_emu = attr_local(e, b"w", reader)?
            .and_then(|value| value.parse().ok())
            .unwrap_or(9525);
        Ok(Outline {
            width_emu,
            color: None,
            dash: None,
            head_end: None,
            tail_end: None,
        })
    }

    fn parse_outline(
        reader: &mut Reader<&[u8]>,
        e: &quick_xml::events::BytesStart<'_>,
        theme: Option<&Theme>,
    ) -> PboResult<Outline> {
        let mut outline = parse_outline_empty(e, reader)?;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(event) => {
                    if local_name(event.name().as_ref()) == b"solidFill"
                        && let Some(Fill::Solid(color)) = parse_solid_fill(reader, theme)?
                    {
                        outline.color = Some(color);
                    }
                }
                Event::Empty(event) => match local_name(event.name().as_ref()) {
                    b"prstDash" => outline.dash = attr_local(&event, b"val", reader)?,
                    b"srgbClr" => {
                        if outline.color.is_none() {
                            outline.color = attr_local(&event, b"val", reader)?;
                        }
                    }
                    b"schemeClr" => {
                        if outline.color.is_none()
                            && let Some(role) = attr_local(&event, b"val", reader)?
                        {
                            outline.color = resolve_theme_color(theme, &role);
                        }
                    }
                    b"headEnd" => {
                        outline.head_end = attr_local(&event, b"type", reader)?;
                    }
                    b"tailEnd" => {
                        outline.tail_end = attr_local(&event, b"type", reader)?;
                    }
                    _ => {}
                },
                Event::End(event) if local_name(event.name().as_ref()) == b"ln" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(outline)
    }

    fn parse_sp(
        reader: &mut Reader<&[u8]>,
        theme: Option<&Theme>,
    ) -> PboResult<Option<SlideElement>> {
        let mut tf = Xf::default();
        let mut prst = None;
        let mut fill = None;
        let mut outline = None;
        let mut style_fill = None;
        let mut style_outline = None;
        let mut text = String::new();
        let mut bold = false;
        let mut italic = false;
        let mut color = None;
        let mut size = None;
        let mut align = None;
        let mut placeholder = None;
        let mut placeholder_idx = None;
        let mut paragraphs = Vec::new();
        let mut hidden = false;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"spPr" => parse_shape_props(
                        reader,
                        &mut tf,
                        &mut prst,
                        &mut fill,
                        &mut outline,
                        theme,
                    )?,
                    b"xfrm" => parse_xfrm(reader, &mut tf)?,
                    b"prstGeom" => prst = attr_local(&e, b"prst", reader)?,
                    b"cNvPr" => {
                        if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") {
                            hidden = true;
                        }
                    }
                    b"ph" => {
                        let raw_type = attr_local(&e, b"type", reader)?;
                        // Normalize ctrTitle -> title (matches OnlyOffice SlideLayout::GetLevelUp)
                        placeholder = raw_type
                            .map(|t| {
                                if t == "ctrTitle" {
                                    "title".to_string()
                                } else {
                                    t
                                }
                            })
                            .or(Some("body".into()));
                        placeholder_idx = attr_local(&e, b"idx", reader)?;
                    }
                    b"txBody" => {
                        paragraphs = parse_tx_body(
                            reader,
                            &mut text,
                            &mut bold,
                            &mut italic,
                            &mut color,
                            &mut size,
                            &mut align,
                            theme,
                        )?;
                    }
                    b"style" => {
                        parse_style_refs(reader, theme, &mut style_fill, &mut style_outline)?;
                    }
                    _ => {}
                },
                Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"spPr" => {} // empty <p:spPr/> - nothing to parse
                    b"prstGeom" => prst = attr_local(&e, b"prst", reader)?,
                    b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                        hidden = true;
                    }
                    b"cNvPr" => {}
                    b"ph" => {
                        let raw_type = attr_local(&e, b"type", reader)?;
                        placeholder = raw_type
                            .map(|t| {
                                if t == "ctrTitle" {
                                    "title".to_string()
                                } else {
                                    t
                                }
                            })
                            .or(Some("body".into()));
                        placeholder_idx = attr_local(&e, b"idx", reader)?;
                    }
                    _ => {}
                },
                Event::End(e) if local_name(e.name().as_ref()) == b"sp" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }

        // Skip hidden elements (e.g. design-time grid guides)
        if hidden {
            return Ok(None);
        }

        // Use style-based fill/outline as fallback when no explicit fill/outline
        if fill.is_none() {
            fill = style_fill;
        }
        if outline.is_none() {
            outline = style_outline;
        }

        // Placeholder shapes → TextBox (for layout/master geometry inheritance)
        if placeholder.is_some() {
            return Ok(Some(SlideElement::TextBox(TextBox {
                text,
                x: tf.x,
                y: tf.y,
                width: tf.w,
                height: tf.h,
                font_size: size,
                bold,
                italic,
                color,
                align,
                rotation: tf.rot,
                placeholder,
                placeholder_idx,
                paragraphs,
                fill,
                outline,
            })));
        }

        // Non-placeholder with fill, outline, or preset geometry → Shape
        if fill.is_some() || outline.is_some() || prst.is_some() {
            let shape_type = prst
                .as_deref()
                .map(map_prst)
                .unwrap_or(ShapeType::Rectangle);
            let inner = if text.is_empty() {
                None
            } else {
                Some(TextBox {
                    text,
                    x: 0.0,
                    y: 0.0,
                    width: tf.w,
                    height: tf.h,
                    font_size: size,
                    bold,
                    italic,
                    color,
                    align,
                    rotation: 0,
                    placeholder: None,
                    placeholder_idx: None,
                    paragraphs,
                    fill: None,
                    outline: None,
                })
            };
            return Ok(Some(SlideElement::Shape(Shape {
                shape_type,
                x: tf.x,
                y: tf.y,
                width: tf.w,
                height: tf.h,
                rotation: tf.rot,
                fill,
                outline,
                text: inner,
            })));
        }

        // Plain text without visual properties → TextBox
        if !text.is_empty() {
            return Ok(Some(SlideElement::TextBox(TextBox {
                text,
                x: tf.x,
                y: tf.y,
                width: tf.w,
                height: tf.h,
                font_size: size,
                bold,
                italic,
                color,
                align,
                rotation: tf.rot,
                placeholder: None,
                placeholder_idx: None,
                paragraphs,
                fill: None,
                outline: None,
            })));
        }

        Ok(None)
    }

    #[allow(clippy::too_many_arguments)]
    fn parse_tx_body(
        reader: &mut Reader<&[u8]>,
        text: &mut String,
        bold: &mut bool,
        italic: &mut bool,
        color: &mut Option<String>,
        size: &mut Option<f32>,
        align: &mut Option<TextAlign>,
        theme: Option<&Theme>,
    ) -> PboResult<Vec<TextParagraph>> {
        let mut paragraphs = Vec::new();
        let mut buf = Vec::new();
        let mut first_para = true;
        let mut first_run = true;
        // Current paragraph state
        let mut para_align: Option<TextAlign> = None;
        let mut para_runs: Vec<TextRun> = Vec::new();
        // Current run state
        let mut run_bold = false;
        let mut run_italic = false;
        let mut run_size: Option<f32> = None;
        let mut run_color: Option<String> = None;
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"p" => {
                        if !first_para {
                            text.push('\n');
                        }
                        first_para = false;
                        para_align = None;
                        para_runs = Vec::new();
                    }
                    b"pPr" => {
                        if let Some(a) = attr_local(&e, b"algn", reader)? {
                            let ta = match a.as_str() {
                                "ctr" => TextAlign::Center,
                                "r" => TextAlign::Right,
                                "just" => TextAlign::Justify,
                                _ => TextAlign::Left,
                            };
                            para_align = Some(ta);
                            if align.is_none() {
                                *align = Some(ta);
                            }
                        }
                    }
                    b"rPr" => {
                        run_bold = attr_local(&e, b"b", reader)?.as_deref() == Some("1");
                        run_italic = attr_local(&e, b"i", reader)?.as_deref() == Some("1");
                        run_size = attr_local(&e, b"sz", reader)?
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 100.0);
                        run_color = None;
                    }
                    b"solidFill" => {
                        // Inside rPr → solidFill
                        if let Some(Fill::Solid(c)) = parse_solid_fill(reader, theme)? {
                            run_color = Some(c);
                        }
                    }
                    b"srgbClr" => {
                        if let Some(v) = attr_local(&e, b"val", reader)?
                            && run_color.is_none()
                        {
                            run_color = Some(v);
                        }
                    }
                    b"schemeClr" => {
                        if let Some(role) = attr_local(&e, b"val", reader)?
                            && run_color.is_none()
                        {
                            run_color = resolve_theme_color(theme, &role);
                        }
                    }
                    b"t" => {
                        let mut run_text = String::new();
                        let mut b2 = Vec::new();
                        loop {
                            match reader.read_event_into(&mut b2).map_err(xml_err)? {
                                Event::Text(t) => run_text.push_str(
                                    &t.unescape().map_err(|e| PboError::Xml(e.to_string()))?,
                                ),
                                Event::End(ee) if local_name(ee.name().as_ref()) == b"t" => break,
                                Event::Eof => break,
                                _ => {}
                            }
                            b2.clear();
                        }
                        text.push_str(&run_text);
                        // Capture first run's properties for backward compat flat fields
                        if first_run && !run_text.is_empty() {
                            *bold = run_bold;
                            *italic = run_italic;
                            if run_size.is_some() {
                                *size = run_size;
                            }
                            if run_color.is_some() {
                                *color = run_color.clone();
                            }
                            first_run = false;
                        }
                        para_runs.push(TextRun {
                            text: run_text,
                            bold: run_bold,
                            italic: run_italic,
                            font_size: run_size,
                            color: run_color.clone(),
                        });
                    }
                    _ => {}
                },
                Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"br" => {
                        text.push('\n');
                        para_runs.push(TextRun {
                            text: "\n".into(),
                            ..Default::default()
                        });
                    }
                    b"rPr" => {
                        run_bold = attr_local(&e, b"b", reader)?.as_deref() == Some("1");
                        run_italic = attr_local(&e, b"i", reader)?.as_deref() == Some("1");
                        run_size = attr_local(&e, b"sz", reader)?
                            .and_then(|v| v.parse::<f32>().ok())
                            .map(|v| v / 100.0);
                        run_color = None;
                    }
                    b"srgbClr" => {
                        if let Some(v) = attr_local(&e, b"val", reader)?
                            && run_color.is_none()
                        {
                            run_color = Some(v);
                        }
                    }
                    b"schemeClr" => {
                        if let Some(role) = attr_local(&e, b"val", reader)?
                            && run_color.is_none()
                        {
                            run_color = resolve_theme_color(theme, &role);
                        }
                    }
                    b"pPr" => {
                        if let Some(a) = attr_local(&e, b"algn", reader)? {
                            let ta = match a.as_str() {
                                "ctr" => TextAlign::Center,
                                "r" => TextAlign::Right,
                                "just" => TextAlign::Justify,
                                _ => TextAlign::Left,
                            };
                            para_align = Some(ta);
                            if align.is_none() {
                                *align = Some(ta);
                            }
                        }
                    }
                    _ => {}
                },
                Event::End(e) => match local_name(e.name().as_ref()) {
                    b"p" if !para_runs.is_empty() => {
                        paragraphs.push(TextParagraph {
                            runs: std::mem::take(&mut para_runs),
                            align: para_align,
                        });
                    }
                    b"p" => {}
                    b"txBody" => {
                        // Flush any remaining paragraph
                        if !para_runs.is_empty() {
                            paragraphs.push(TextParagraph {
                                runs: std::mem::take(&mut para_runs),
                                align: para_align,
                            });
                        }
                        break;
                    }
                    _ => {}
                },
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(paragraphs)
    }

    fn map_prst(p: &str) -> ShapeType {
        match p {
            "rect" => ShapeType::Rectangle,
            "roundRect" => ShapeType::RoundedRectangle,
            "ellipse" | "oval" => ShapeType::Ellipse,
            "triangle" => ShapeType::Triangle,
            "rtTriangle" => ShapeType::RightTriangle,
            "diamond" => ShapeType::Diamond,
            "parallelogram" => ShapeType::Parallelogram,
            "trapezoid" => ShapeType::Trapezoid,
            "pentagon" => ShapeType::Pentagon,
            "hexagon" => ShapeType::Hexagon,
            "octagon" => ShapeType::Octagon,
            "star5" => ShapeType::Star5,
            "rightArrow" | "leftArrow" | "upArrow" | "downArrow" => ShapeType::Arrow,
            "line" => ShapeType::Line,
            s if s.starts_with("callout") || s.starts_with("wedgeCallout") => ShapeType::Callout,
            _ => ShapeType::Other,
        }
    }

    fn parse_pic(
        reader: &mut Reader<&[u8]>,
        rels: &Relationships,
        slide_path: &str,
        package: &OpcPackage,
        images: &mut Vec<SlidesImage>,
    ) -> PboResult<Option<SlideElement>> {
        let mut tf = Xf::default();
        let mut rel_id: Option<String> = None;
        let mut name: Option<String> = None;
        let mut hidden = false;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"blip" => {
                        // Check for r:embed directly on blip (standard images)
                        if let Some(rid) = attr_qualified(&e, b"r:embed", reader)? {
                            rel_id = Some(rid);
                        }
                    }
                    b"svgBlip" => {
                        // SVG images use <asvg:svgBlip r:embed="..."/> inside <a:blip>
                        if let Some(rid) = attr_qualified(&e, b"r:embed", reader)? {
                            rel_id = Some(rid);
                        }
                    }
                    b"cNvPr" => {
                        name = attr_local(&e, b"name", reader)?;
                        if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") {
                            hidden = true;
                        }
                    }
                    b"xfrm" => parse_xfrm(reader, &mut tf)?,
                    _ => {}
                },
                Event::End(e) if local_name(e.name().as_ref()) == b"pic" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        if hidden {
            return Ok(None);
        }
        let rid = match rel_id {
            Some(r) => r,
            None => return Ok(None),
        };
        let rel = match rels.find_by_id(&rid) {
            Some(r) if r.rel_type == REL_IMAGE => r,
            _ => return Ok(None),
        };
        let image_path = normalize(&rel.target, slide_path);
        let logical = ensure_image(package, &image_path, images);
        Ok(Some(SlideElement::Picture(Picture {
            image_id: logical,
            name,
            x: tf.x,
            y: tf.y,
            width: tf.w,
            height: tf.h,
            rotation: tf.rot,
        })))
    }

    fn ensure_image(package: &OpcPackage, path: &str, images: &mut Vec<SlidesImage>) -> String {
        if images.iter().any(|i| i.id == path) {
            return path.to_string();
        }
        let bytes = package.parts.get(path).cloned().unwrap_or_default();
        let ext = path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
        let mime = match ext.as_str() {
            "png" => "image/png",
            "jpg" | "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "svg" => "image/svg+xml",
            _ => "application/octet-stream",
        }
        .to_string();
        images.push(SlidesImage {
            id: path.to_string(),
            mime,
            data_base64: b64::encode(&bytes),
        });
        path.to_string()
    }

    fn parse_cxn(
        reader: &mut Reader<&[u8]>,
        theme: Option<&Theme>,
    ) -> PboResult<Option<SlideElement>> {
        let mut tf = Xf::default();
        let mut prst = String::new();
        let mut outline = None;
        let mut hidden = false;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                        hidden = true;
                    }
                    b"cNvPr" => {}
                    b"spPr" => {
                        let mut maybe_prst = if prst.is_empty() {
                            None
                        } else {
                            Some(prst.clone())
                        };
                        let mut fill = None;
                        parse_shape_props(
                            reader,
                            &mut tf,
                            &mut maybe_prst,
                            &mut fill,
                            &mut outline,
                            theme,
                        )?;
                        if let Some(parsed_prst) = maybe_prst {
                            prst = parsed_prst;
                        }
                    }
                    b"prstGeom" => {
                        if let Some(p) = attr_local(&e, b"prst", reader)? {
                            prst = p;
                        }
                    }
                    b"xfrm" => parse_xfrm(reader, &mut tf)?,
                    _ => {}
                },
                Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"spPr" => {} // empty <p:spPr/> - nothing to parse
                    b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                        hidden = true;
                    }
                    b"cNvPr" => {}
                    b"prstGeom" => {
                        if let Some(p) = attr_local(&e, b"prst", reader)? {
                            prst = p;
                        }
                    }
                    _ => {}
                },
                Event::End(e) if local_name(e.name().as_ref()) == b"cxnSp" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        if hidden {
            return Ok(None);
        }
        let connector_type = if prst.starts_with("bent") {
            ConnectorType::Bent
        } else if prst.starts_with("curved") {
            ConnectorType::Curved
        } else {
            ConnectorType::Straight
        };
        Ok(Some(SlideElement::Connector(Connector {
            connector_type,
            x: tf.x,
            y: tf.y,
            width: tf.w,
            height: tf.h,
            flip_h: tf.flip_h,
            flip_v: tf.flip_v,
            rotation: tf.rot,
            outline,
        })))
    }

    #[derive(Default)]
    struct GroupXf {
        off: Xf,
        ch_off_x: f64,
        ch_off_y: f64,
        ch_ext_cx: f64,
        ch_ext_cy: f64,
    }

    fn parse_group(
        reader: &mut Reader<&[u8]>,
        rels: &Relationships,
        slide_path: &str,
        package: &OpcPackage,
        images: &mut Vec<SlidesImage>,
        theme: Option<&Theme>,
    ) -> PboResult<Option<GroupShape>> {
        let mut gxf = GroupXf::default();
        let mut elements = Vec::new();
        let mut hidden = false;
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) => match local_name(e.name().as_ref()) {
                    b"grpSpPr" => parse_group_sp_pr(reader, &mut gxf)?,
                    b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                        hidden = true;
                    }
                    b"cNvPr" => {}
                    b"sp" => {
                        if let Some(el) = parse_sp(reader, theme)? {
                            elements.push(el);
                        }
                    }
                    b"pic" => {
                        if let Some(el) = parse_pic(reader, rels, slide_path, package, images)? {
                            elements.push(el);
                        }
                    }
                    b"cxnSp" => {
                        if let Some(el) = parse_cxn(reader, theme)? {
                            elements.push(el);
                        }
                    }
                    b"grpSp" => {
                        if let Some(g) =
                            parse_group(reader, rels, slide_path, package, images, theme)?
                        {
                            elements.push(SlideElement::Group(g));
                        }
                    }
                    _ => {}
                },
                Event::Empty(e) => {
                    if local_name(e.name().as_ref()) == b"cNvPr"
                        && attr_local(&e, b"hidden", reader)?.as_deref() == Some("1")
                    {
                        hidden = true;
                    }
                }
                Event::End(e) if local_name(e.name().as_ref()) == b"grpSp" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        if hidden {
            return Ok(None);
        }
        Ok(Some(GroupShape {
            x: gxf.off.x,
            y: gxf.off.y,
            width: gxf.off.w,
            height: gxf.off.h,
            ch_off_x: gxf.ch_off_x,
            ch_off_y: gxf.ch_off_y,
            ch_ext_cx: gxf.ch_ext_cx,
            ch_ext_cy: gxf.ch_ext_cy,
            elements,
        }))
    }

    fn parse_group_sp_pr(reader: &mut Reader<&[u8]>, gxf: &mut GroupXf) -> PboResult<()> {
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf).map_err(xml_err)? {
                Event::Start(e) => {
                    if local_name(e.name().as_ref()) == b"xfrm" {
                        if let Some(r) =
                            attr_local(&e, b"rot", reader)?.and_then(|v| v.parse().ok())
                        {
                            gxf.off.rot = r;
                        }
                        if attr_local(&e, b"flipH", reader)?.as_deref() == Some("1") {
                            gxf.off.flip_h = true;
                        }
                        if attr_local(&e, b"flipV", reader)?.as_deref() == Some("1") {
                            gxf.off.flip_v = true;
                        }
                    }
                }
                Event::Empty(e) => match local_name(e.name().as_ref()) {
                    b"off" => {
                        if let Some(x) = attr_i64(&e, b"x", reader)? {
                            gxf.off.x = x as f64;
                        }
                        if let Some(y) = attr_i64(&e, b"y", reader)? {
                            gxf.off.y = y as f64;
                        }
                    }
                    b"ext" => {
                        if let Some(cx) = attr_i64(&e, b"cx", reader)? {
                            gxf.off.w = cx as f64;
                        }
                        if let Some(cy) = attr_i64(&e, b"cy", reader)? {
                            gxf.off.h = cy as f64;
                        }
                    }
                    b"chOff" => {
                        if let Some(x) = attr_i64(&e, b"x", reader)? {
                            gxf.ch_off_x = x as f64;
                        }
                        if let Some(y) = attr_i64(&e, b"y", reader)? {
                            gxf.ch_off_y = y as f64;
                        }
                    }
                    b"chExt" => {
                        if let Some(cx) = attr_i64(&e, b"cx", reader)? {
                            gxf.ch_ext_cx = cx as f64;
                        }
                        if let Some(cy) = attr_i64(&e, b"cy", reader)? {
                            gxf.ch_ext_cy = cy as f64;
                        }
                    }
                    _ => {}
                },
                Event::End(e) if local_name(e.name().as_ref()) == b"grpSpPr" => break,
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(())
    }
}

fn normalize(target: &str, base_dir: &str) -> String {
    if let Some(rest) = target.strip_prefix('/') {
        return rest.into();
    }
    let dir = base_dir.trim_end_matches('/');
    let parent = dir.rsplit_once('/').map(|(p, _)| p).unwrap_or(dir);
    if let Some(rest) = target.strip_prefix("../") {
        let pp = parent.rsplit_once('/').map(|(p, _)| p).unwrap_or("");
        if pp.is_empty() {
            rest.into()
        } else {
            format!("{pp}/{rest}")
        }
    } else if parent.is_empty() {
        target.into()
    } else {
        format!("{parent}/{target}")
    }
}

fn rels_path(part: &str) -> String {
    let (dir, file) = part.rsplit_once('/').unwrap_or(("", part));
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

fn xml_err(e: quick_xml::Error) -> PboError {
    PboError::Xml(e.to_string())
}

fn parse_slide_size(xml: &str) -> PboResult<(f64, f64)> {
    let mut r = Reader::from_str(xml);
    r.config_mut().trim_text(false);
    let (mut w, mut h) = (9144000.0, 6858000.0);
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) if local_name(e.name().as_ref()) == b"sldSz" => {
                if let Some(cx) = attr_local(&e, b"cx", &r)?.and_then(|v| v.parse().ok()) {
                    w = cx;
                }
                if let Some(cy) = attr_local(&e, b"cy", &r)?.and_then(|v| v.parse().ok()) {
                    h = cy;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok((w, h))
}

fn parse_slide_id_list(xml: &str) -> PboResult<Vec<String>> {
    let mut r = Reader::from_str(xml);
    r.config_mut().trim_text(false);
    let mut out = Vec::new();
    let mut buf = Vec::new();
    loop {
        match r.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) if local_name(e.name().as_ref()) == b"sldId" => {
                if let Some(id) = attr_qualified(&e, b"r:id", &r)? {
                    out.push(id);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn parse_slide(
    xml: &str,
    rels: &Relationships,
    slide_path: &str,
    package: &OpcPackage,
    images: &mut Vec<SlidesImage>,
    theme: Option<&Theme>,
) -> PboResult<(Slide, Option<String>)> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut slide = Slide::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"spTree" => {
                slide.elements =
                    parse_sp_tree(&mut reader, rels, slide_path, package, images, theme)?;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    let notes = rels
        .0
        .iter()
        .find(|r| r.rel_type == REL_NOTES)
        .map(|r| normalize(&r.target, slide_path));
    Ok((slide, notes))
}

fn parse_layout(
    xml: &str,
    rels: &Relationships,
    layout_path: &str,
    package: &OpcPackage,
    images: &mut Vec<SlidesImage>,
    theme: Option<&Theme>,
) -> PboResult<SlideLayout> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut layout = SlideLayout::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"cSld" => {
                if let Some(name) = attr_local(&e, b"name", &reader)? {
                    layout.name = name;
                }
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"spTree" => {
                layout.elements =
                    parse_sp_tree(&mut reader, rels, layout_path, package, images, theme)?;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(layout)
}

fn parse_master(
    xml: &str,
    rels: &Relationships,
    master_path: &str,
    package: &OpcPackage,
    images: &mut Vec<SlidesImage>,
    theme: Option<&Theme>,
) -> PboResult<SlideMaster> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut master = SlideMaster::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"cSld" => {
                if let Some(name) = attr_local(&e, b"name", &reader)? {
                    master.name = name;
                }
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"spTree" => {
                master.elements =
                    parse_sp_tree(&mut reader, rels, master_path, package, images, theme)?;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(master)
}

fn parse_theme(xml: &str) -> PboResult<Theme> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut theme = Theme::default();
    let mut current_role: Option<String> = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name == b"theme" {
                    if let Some(value) = attr_local(&e, b"name", &reader)? {
                        theme.name = value;
                    }
                } else if is_theme_color_role(name) {
                    current_role = Some(String::from_utf8_lossy(name).into_owned());
                } else if name == b"srgbClr" {
                    if let (Some(role), Some(value)) =
                        (current_role.clone(), attr_local(&e, b"val", &reader)?)
                    {
                        theme.colors.push(ThemeColor { role, value });
                    }
                } else if name == b"sysClr"
                    && let (Some(role), Some(value)) =
                        (current_role.clone(), attr_local(&e, b"lastClr", &reader)?)
                {
                    theme.colors.push(ThemeColor { role, value });
                }
            }
            Event::End(e) if is_theme_color_role(local_name(e.name().as_ref())) => {
                current_role = None;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(theme)
}

fn is_theme_color_role(name: &[u8]) -> bool {
    matches!(
        name,
        b"dk1"
            | b"lt1"
            | b"dk2"
            | b"lt2"
            | b"accent1"
            | b"accent2"
            | b"accent3"
            | b"accent4"
            | b"accent5"
            | b"accent6"
            | b"hlink"
            | b"folHlink"
    )
}

fn resolve_theme_color(theme: Option<&Theme>, role: &str) -> Option<String> {
    let theme = theme?;
    let normalized_role = match role {
        "bg1" => "lt1",
        "tx1" => "dk1",
        "bg2" => "lt2",
        "tx2" => "dk2",
        other => other,
    };
    theme
        .colors
        .iter()
        .find(|color| color.role == normalized_role)
        .map(|color| color.value.clone())
}

/// Apply color modifiers (lumMod, lumOff, tint, shade, alpha) to an RGB hex string.
fn apply_color_modifiers(
    hex: &str,
    lum_mod: Option<i64>,
    lum_off: Option<i64>,
    tint: Option<i64>,
) -> String {
    let r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0) as f64;
    let g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0) as f64;
    let b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0) as f64;

    let (mut r, mut g, mut b) = (r, g, b);

    // Apply tint: tint towards white
    if let Some(t) = tint {
        let factor = t as f64 / 100000.0;
        r = r + (255.0 - r) * factor;
        g = g + (255.0 - g) * factor;
        b = b + (255.0 - b) * factor;
    }

    // Apply lumMod/lumOff: modulate luminance
    if lum_mod.is_some() || lum_off.is_some() {
        let mod_val = lum_mod.unwrap_or(100000) as f64 / 100000.0;
        let off_val = lum_off.unwrap_or(0) as f64 / 100000.0 * 255.0;
        r = (r * mod_val + off_val).clamp(0.0, 255.0);
        g = (g * mod_val + off_val).clamp(0.0, 255.0);
        b = (b * mod_val + off_val).clamp(0.0, 255.0);
    }

    format!("{:02X}{:02X}{:02X}", r as u8, g as u8, b as u8)
}

/// Resolve a color element (srgbClr or schemeClr with modifiers) reading child events.
fn resolve_color_with_modifiers(
    reader: &mut Reader<&[u8]>,
    start_tag: &[u8],
    base_color: &str,
    theme: Option<&Theme>,
) -> PboResult<Option<String>> {
    let is_scheme = start_tag == b"schemeClr";
    let base = if is_scheme {
        resolve_theme_color(theme, base_color)
    } else {
        Some(base_color.to_string())
    };
    let Some(hex) = base else { return Ok(None) };

    let mut lum_mod = None;
    let mut lum_off = None;
    let mut tint = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"lumMod" => lum_mod = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
                b"lumOff" => lum_off = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
                b"tint" => tint = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
                b"shade" => {
                    // shade is inverse of tint — darken
                    if let Some(v) =
                        attr_local(&e, b"val", reader)?.and_then(|v| v.parse::<i64>().ok())
                    {
                        lum_mod = Some(v);
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == start_tag => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if lum_mod.is_some() || lum_off.is_some() || tint.is_some() {
        Ok(Some(apply_color_modifiers(&hex, lum_mod, lum_off, tint)))
    } else {
        Ok(Some(hex))
    }
}

/// Parse `<p:style>` to extract fill and outline references from the theme style matrix.
fn parse_style_refs(
    reader: &mut Reader<&[u8]>,
    theme: Option<&Theme>,
    style_fill: &mut Option<Fill>,
    style_outline: &mut Option<Outline>,
) -> PboResult<()> {
    let mut buf = Vec::new();
    let mut in_fill_ref = false;
    let mut in_ln_ref = false;
    let mut fill_ref_idx: Option<u32> = None;
    let mut ln_ref_idx: Option<u32> = None;
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"fillRef" => {
                    fill_ref_idx = attr_local(&e, b"idx", reader)?.and_then(|v| v.parse().ok());
                    in_fill_ref = true;
                }
                b"lnRef" => {
                    ln_ref_idx = attr_local(&e, b"idx", reader)?.and_then(|v| v.parse().ok());
                    in_ln_ref = true;
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"schemeClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)?
                        && let Some(resolved) = resolve_theme_color(theme, &val)
                    {
                        if in_fill_ref && fill_ref_idx.is_some_and(|i| i > 0) {
                            *style_fill = Some(Fill::Solid(resolved.clone()));
                        }
                        if in_ln_ref && ln_ref_idx.is_some_and(|i| i > 0) {
                            *style_outline = Some(Outline {
                                width_emu: 12700,
                                color: Some(resolved),
                                dash: None,
                                head_end: None,
                                tail_end: None,
                            });
                        }
                    }
                }
                b"srgbClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)? {
                        if in_fill_ref && fill_ref_idx.is_some_and(|i| i > 0) {
                            *style_fill = Some(Fill::Solid(val.clone()));
                        }
                        if in_ln_ref && ln_ref_idx.is_some_and(|i| i > 0) {
                            *style_outline = Some(Outline {
                                width_emu: 12700,
                                color: Some(val),
                                dash: None,
                                head_end: None,
                                tail_end: None,
                            });
                        }
                    }
                }
                _ => {}
            },
            Event::End(e) => match local_name(e.name().as_ref()) {
                b"fillRef" => in_fill_ref = false,
                b"lnRef" => in_ln_ref = false,
                b"style" => break,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_sp_tree(
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    slide_path: &str,
    package: &OpcPackage,
    images: &mut Vec<SlidesImage>,
    theme: Option<&Theme>,
) -> PboResult<Vec<SlideElement>> {
    let mut out = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"sp" => {
                    if let Some(el) = parse_sp(reader, theme)? {
                        out.push(el);
                    }
                }
                b"pic" => {
                    if let Some(el) = parse_pic(reader, rels, slide_path, package, images)? {
                        out.push(el);
                    }
                }
                b"cxnSp" => {
                    if let Some(el) = parse_cxn(reader, theme)? {
                        out.push(el);
                    }
                }
                b"grpSp" => {
                    if let Some(g) = parse_group(reader, rels, slide_path, package, images, theme)?
                    {
                        out.push(SlideElement::Group(g));
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"spTree" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

#[derive(Default)]
struct Xf {
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    rot: i32,
    flip_h: bool,
    flip_v: bool,
}

fn parse_xfrm(reader: &mut Reader<&[u8]>, tf: &mut Xf) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) => {
                let qname = e.name();
                let name = local_name(qname.as_ref());
                if name == b"off" {
                    if let Some(x) = attr_i64(&e, b"x", reader)? {
                        tf.x = x as f64;
                    }
                    if let Some(y) = attr_i64(&e, b"y", reader)? {
                        tf.y = y as f64;
                    }
                } else if name == b"ext" {
                    if let Some(cx) = attr_i64(&e, b"cx", reader)? {
                        tf.w = cx as f64;
                    }
                    if let Some(cy) = attr_i64(&e, b"cy", reader)? {
                        tf.h = cy as f64;
                    }
                } else if name == b"xfrm" {
                    if let Some(r) = attr_local(&e, b"rot", reader)?.and_then(|v| v.parse().ok()) {
                        tf.rot = r;
                    }
                    if attr_local(&e, b"flipH", reader)?.as_deref() == Some("1") {
                        tf.flip_h = true;
                    }
                    if attr_local(&e, b"flipV", reader)?.as_deref() == Some("1") {
                        tf.flip_v = true;
                    }
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"xfrm" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_shape_props(
    reader: &mut Reader<&[u8]>,
    tf: &mut Xf,
    prst: &mut Option<String>,
    fill: &mut Option<Fill>,
    outline: &mut Option<Outline>,
    theme: Option<&Theme>,
) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"xfrm" => parse_xfrm(reader, tf)?,
                b"prstGeom" => *prst = attr_local(&e, b"prst", reader)?,
                b"solidFill" => *fill = parse_solid_fill(reader, theme)?,
                b"gradFill" => *fill = parse_gradient_fill(reader, theme)?,
                b"ln" => *outline = Some(parse_outline(reader, &e, theme)?),
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"prstGeom" => *prst = attr_local(&e, b"prst", reader)?,
                b"noFill" => *fill = Some(Fill::None),
                b"ln" => *outline = Some(parse_outline_empty(&e, reader)?),
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"spPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_solid_fill(reader: &mut Reader<&[u8]>, theme: Option<&Theme>) -> PboResult<Option<Fill>> {
    let mut color = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"srgbClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)? {
                        color = resolve_color_with_modifiers(reader, b"srgbClr", &val, theme)?;
                    }
                }
                b"schemeClr" => {
                    if let Some(role) = attr_local(&e, b"val", reader)? {
                        color = resolve_color_with_modifiers(reader, b"schemeClr", &role, theme)?;
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"srgbClr" => {
                    color = attr_local(&e, b"val", reader)?;
                }
                b"schemeClr" => {
                    if let Some(role) = attr_local(&e, b"val", reader)? {
                        color = resolve_theme_color(theme, &role);
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"solidFill" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(color.map(Fill::Solid))
}

fn parse_gradient_fill(
    reader: &mut Reader<&[u8]>,
    theme: Option<&Theme>,
) -> PboResult<Option<Fill>> {
    let mut stops = Vec::new();
    let mut current_pos = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"gs" => {
                    current_pos =
                        attr_local(&e, b"pos", reader)?.and_then(|value| value.parse().ok());
                }
                b"srgbClr" => {
                    if let Some(pos) = current_pos
                        && let Some(val) = attr_local(&e, b"val", reader)?
                        && let Some(color) =
                            resolve_color_with_modifiers(reader, b"srgbClr", &val, theme)?
                    {
                        stops.push(GradientStop {
                            position: pos,
                            color,
                        });
                    }
                }
                b"schemeClr" => {
                    if let Some(pos) = current_pos
                        && let Some(role) = attr_local(&e, b"val", reader)?
                        && let Some(color) =
                            resolve_color_with_modifiers(reader, b"schemeClr", &role, theme)?
                    {
                        stops.push(GradientStop {
                            position: pos,
                            color,
                        });
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"srgbClr" => {
                    if let (Some(pos), Some(color)) = (current_pos, attr_local(&e, b"val", reader)?)
                    {
                        stops.push(GradientStop {
                            position: pos,
                            color,
                        });
                    }
                }
                b"schemeClr" => {
                    if let (Some(pos), Some(role)) = (current_pos, attr_local(&e, b"val", reader)?)
                        && let Some(color) = resolve_theme_color(theme, &role)
                    {
                        stops.push(GradientStop {
                            position: pos,
                            color,
                        });
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"gs" => current_pos = None,
            Event::End(e) if local_name(e.name().as_ref()) == b"gradFill" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    if stops.is_empty() {
        Ok(None)
    } else {
        Ok(Some(Fill::Gradient(stops)))
    }
}

fn parse_outline_empty(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<&[u8]>,
) -> PboResult<Outline> {
    let width_emu = attr_local(e, b"w", reader)?
        .and_then(|value| value.parse().ok())
        .unwrap_or(9525);
    Ok(Outline {
        width_emu,
        color: None,
        dash: None,
        head_end: None,
        tail_end: None,
    })
}

fn parse_outline(
    reader: &mut Reader<&[u8]>,
    e: &quick_xml::events::BytesStart<'_>,
    theme: Option<&Theme>,
) -> PboResult<Outline> {
    let mut outline = parse_outline_empty(e, reader)?;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(event) if local_name(event.name().as_ref()) == b"solidFill" => {
                if let Some(Fill::Solid(color)) = parse_solid_fill(reader, theme)? {
                    outline.color = Some(color);
                }
            }
            Event::Empty(event) if local_name(event.name().as_ref()) == b"prstDash" => {
                outline.dash = attr_local(&event, b"val", reader)?;
            }
            Event::Empty(event)
                if local_name(event.name().as_ref()) == b"srgbClr" && outline.color.is_none() =>
            {
                outline.color = attr_local(&event, b"val", reader)?;
            }
            Event::Empty(event)
                if local_name(event.name().as_ref()) == b"schemeClr" && outline.color.is_none() =>
            {
                if let Some(role) = attr_local(&event, b"val", reader)? {
                    outline.color = resolve_theme_color(theme, &role);
                }
            }
            Event::Empty(event) if local_name(event.name().as_ref()) == b"headEnd" => {
                outline.head_end = attr_local(&event, b"type", reader)?;
            }
            Event::Empty(event) if local_name(event.name().as_ref()) == b"tailEnd" => {
                outline.tail_end = attr_local(&event, b"type", reader)?;
            }
            Event::End(event) if local_name(event.name().as_ref()) == b"ln" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(outline)
}

fn parse_sp(reader: &mut Reader<&[u8]>, theme: Option<&Theme>) -> PboResult<Option<SlideElement>> {
    let mut tf = Xf::default();
    let mut prst: Option<String> = None;
    let mut fill = None;
    let mut outline = None;
    let mut text = String::new();
    let mut bold = false;
    let mut italic = false;
    let mut color: Option<String> = None;
    let mut size: Option<f32> = None;
    let mut align: Option<TextAlign> = None;
    let mut placeholder: Option<String> = None;
    let mut placeholder_idx: Option<String> = None;
    let mut paragraphs = Vec::new();
    let mut hidden = false;
    let mut style_fill: Option<Fill> = None;
    let mut style_outline: Option<Outline> = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"spPr" => {
                    parse_shape_props(reader, &mut tf, &mut prst, &mut fill, &mut outline, theme)?
                }
                b"xfrm" => parse_xfrm(reader, &mut tf)?,
                b"prstGeom" => prst = attr_local(&e, b"prst", reader)?,
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"cNvPr" => {}
                b"ph" => {
                    let raw_type = attr_local(&e, b"type", reader)?;
                    // Normalize ctrTitle -> title (matches OnlyOffice SlideLayout::GetLevelUp)
                    placeholder = raw_type
                        .map(|t| {
                            if t == "ctrTitle" {
                                "title".to_string()
                            } else {
                                t
                            }
                        })
                        .or(Some("body".into()));
                    placeholder_idx = attr_local(&e, b"idx", reader)?;
                }
                b"txBody" => {
                    paragraphs = parse_tx_body(
                        reader,
                        &mut text,
                        &mut bold,
                        &mut italic,
                        &mut color,
                        &mut size,
                        &mut align,
                        theme,
                    )?;
                }
                b"style" => {
                    parse_style_refs(reader, theme, &mut style_fill, &mut style_outline)?;
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"spPr" => {} // empty <p:spPr/> - nothing to parse
                b"prstGeom" => prst = attr_local(&e, b"prst", reader)?,
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"cNvPr" => {}
                b"ph" => {
                    let raw_type = attr_local(&e, b"type", reader)?;
                    placeholder = raw_type
                        .map(|t| {
                            if t == "ctrTitle" {
                                "title".to_string()
                            } else {
                                t
                            }
                        })
                        .or(Some("body".into()));
                    placeholder_idx = attr_local(&e, b"idx", reader)?;
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"sp" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    // Skip hidden elements (e.g. design-time grid guides)
    if hidden {
        return Ok(None);
    }

    // Use style-based fill/outline as fallback when no explicit fill/outline
    if fill.is_none() {
        fill = style_fill;
    }
    if outline.is_none() {
        outline = style_outline;
    }

    // Placeholder shapes → TextBox (for layout/master geometry inheritance)
    if placeholder.is_some() {
        return Ok(Some(SlideElement::TextBox(TextBox {
            text,
            x: tf.x,
            y: tf.y,
            width: tf.w,
            height: tf.h,
            font_size: size,
            bold,
            italic,
            color,
            align,
            rotation: tf.rot,
            placeholder,
            placeholder_idx,
            paragraphs,
            fill,
            outline,
        })));
    }

    // Non-placeholder with fill, outline, or preset geometry → Shape
    if fill.is_some() || outline.is_some() || prst.is_some() {
        let shape_type = prst
            .as_deref()
            .map(map_prst)
            .unwrap_or(ShapeType::Rectangle);
        let inner = if text.is_empty() {
            None
        } else {
            Some(TextBox {
                text,
                x: 0.0,
                y: 0.0,
                width: tf.w,
                height: tf.h,
                font_size: size,
                bold,
                italic,
                color,
                align,
                rotation: 0,
                placeholder: None,
                placeholder_idx: None,
                paragraphs,
                fill: None,
                outline: None,
            })
        };
        return Ok(Some(SlideElement::Shape(Shape {
            shape_type,
            x: tf.x,
            y: tf.y,
            width: tf.w,
            height: tf.h,
            rotation: tf.rot,
            fill,
            outline,
            text: inner,
        })));
    }

    // Plain text without visual properties → TextBox
    if !text.is_empty() {
        return Ok(Some(SlideElement::TextBox(TextBox {
            text,
            x: tf.x,
            y: tf.y,
            width: tf.w,
            height: tf.h,
            font_size: size,
            bold,
            italic,
            color,
            align,
            rotation: tf.rot,
            placeholder: None,
            placeholder_idx: None,
            paragraphs,
            fill: None,
            outline: None,
        })));
    }

    Ok(None)
}

#[allow(clippy::too_many_arguments)]
fn parse_tx_body(
    reader: &mut Reader<&[u8]>,
    text: &mut String,
    bold: &mut bool,
    italic: &mut bool,
    color: &mut Option<String>,
    size: &mut Option<f32>,
    align: &mut Option<TextAlign>,
    theme: Option<&Theme>,
) -> PboResult<Vec<TextParagraph>> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();
    let mut first_para = true;
    let mut first_run = true;
    // Current paragraph state
    let mut para_align: Option<TextAlign> = None;
    let mut para_runs: Vec<TextRun> = Vec::new();
    // Current run state
    let mut run_bold = false;
    let mut run_italic = false;
    let mut run_size: Option<f32> = None;
    let mut run_color: Option<String> = None;
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"p" => {
                    if !first_para {
                        text.push('\n');
                    }
                    first_para = false;
                    para_align = None;
                    para_runs = Vec::new();
                }
                b"pPr" => {
                    if let Some(a) = attr_local(&e, b"algn", reader)? {
                        let ta = match a.as_str() {
                            "ctr" => TextAlign::Center,
                            "r" => TextAlign::Right,
                            "just" => TextAlign::Justify,
                            _ => TextAlign::Left,
                        };
                        para_align = Some(ta);
                        if align.is_none() {
                            *align = Some(ta);
                        }
                    }
                }
                b"rPr" => {
                    run_bold = attr_local(&e, b"b", reader)?.as_deref() == Some("1");
                    run_italic = attr_local(&e, b"i", reader)?.as_deref() == Some("1");
                    run_size = attr_local(&e, b"sz", reader)?
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|v| v / 100.0);
                    run_color = None;
                }
                b"solidFill" => {
                    // Inside rPr → solidFill
                    if let Some(Fill::Solid(c)) = parse_solid_fill(reader, theme)? {
                        run_color = Some(c);
                    }
                }
                b"srgbClr" => {
                    if run_color.is_none()
                        && let Some(v) = attr_local(&e, b"val", reader)?
                    {
                        run_color = Some(v);
                    }
                }
                b"schemeClr" => {
                    if run_color.is_none()
                        && let Some(role) = attr_local(&e, b"val", reader)?
                    {
                        run_color = resolve_theme_color(theme, &role);
                    }
                }
                b"t" => {
                    let mut run_text = String::new();
                    let mut b2 = Vec::new();
                    loop {
                        match reader.read_event_into(&mut b2).map_err(xml_err)? {
                            Event::Text(t) => run_text
                                .push_str(&t.unescape().map_err(|e| PboError::Xml(e.to_string()))?),
                            Event::End(ee) if local_name(ee.name().as_ref()) == b"t" => break,
                            Event::Eof => break,
                            _ => {}
                        }
                        b2.clear();
                    }
                    text.push_str(&run_text);
                    // Capture first run's properties for backward compat flat fields
                    if first_run && !run_text.is_empty() {
                        *bold = run_bold;
                        *italic = run_italic;
                        if run_size.is_some() {
                            *size = run_size;
                        }
                        if run_color.is_some() {
                            *color = run_color.clone();
                        }
                        first_run = false;
                    }
                    para_runs.push(TextRun {
                        text: run_text,
                        bold: run_bold,
                        italic: run_italic,
                        font_size: run_size,
                        color: run_color.clone(),
                    });
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"br" => {
                    text.push('\n');
                    para_runs.push(TextRun {
                        text: "\n".into(),
                        ..Default::default()
                    });
                }
                b"rPr" => {
                    run_bold = attr_local(&e, b"b", reader)?.as_deref() == Some("1");
                    run_italic = attr_local(&e, b"i", reader)?.as_deref() == Some("1");
                    run_size = attr_local(&e, b"sz", reader)?
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|v| v / 100.0);
                    run_color = None;
                }
                b"srgbClr" => {
                    if run_color.is_none()
                        && let Some(v) = attr_local(&e, b"val", reader)?
                    {
                        run_color = Some(v);
                    }
                }
                b"schemeClr" => {
                    if run_color.is_none()
                        && let Some(role) = attr_local(&e, b"val", reader)?
                    {
                        run_color = resolve_theme_color(theme, &role);
                    }
                }
                b"pPr" => {
                    if let Some(a) = attr_local(&e, b"algn", reader)? {
                        let ta = match a.as_str() {
                            "ctr" => TextAlign::Center,
                            "r" => TextAlign::Right,
                            "just" => TextAlign::Justify,
                            _ => TextAlign::Left,
                        };
                        para_align = Some(ta);
                        if align.is_none() {
                            *align = Some(ta);
                        }
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"p" && !para_runs.is_empty() => {
                paragraphs.push(TextParagraph {
                    runs: std::mem::take(&mut para_runs),
                    align: para_align,
                });
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"txBody" => {
                // Flush any remaining paragraph
                if !para_runs.is_empty() {
                    paragraphs.push(TextParagraph {
                        runs: std::mem::take(&mut para_runs),
                        align: para_align,
                    });
                }
                break;
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(paragraphs)
}

fn map_prst(p: &str) -> ShapeType {
    match p {
        "rect" => ShapeType::Rectangle,
        "roundRect" => ShapeType::RoundedRectangle,
        "ellipse" | "oval" => ShapeType::Ellipse,
        "triangle" => ShapeType::Triangle,
        "rtTriangle" => ShapeType::RightTriangle,
        "diamond" => ShapeType::Diamond,
        "parallelogram" => ShapeType::Parallelogram,
        "trapezoid" => ShapeType::Trapezoid,
        "pentagon" => ShapeType::Pentagon,
        "hexagon" => ShapeType::Hexagon,
        "octagon" => ShapeType::Octagon,
        "star5" => ShapeType::Star5,
        "rightArrow" | "leftArrow" | "upArrow" | "downArrow" => ShapeType::Arrow,
        "line" => ShapeType::Line,
        s if s.starts_with("callout") || s.starts_with("wedgeCallout") => ShapeType::Callout,
        _ => ShapeType::Other,
    }
}

fn parse_pic(
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    slide_path: &str,
    package: &OpcPackage,
    images: &mut Vec<SlidesImage>,
) -> PboResult<Option<SlideElement>> {
    let mut tf = Xf::default();
    let mut rel_id: Option<String> = None;
    let mut name: Option<String> = None;
    let mut hidden = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"blip" => {
                    // Check for r:embed directly on blip (standard images)
                    if let Some(rid) = attr_qualified(&e, b"r:embed", reader)? {
                        rel_id = Some(rid);
                    }
                }
                b"svgBlip" => {
                    // SVG images use <asvg:svgBlip r:embed="..."/> inside <a:blip>
                    if let Some(rid) = attr_qualified(&e, b"r:embed", reader)? {
                        rel_id = Some(rid);
                    }
                }
                b"cNvPr" => {
                    name = attr_local(&e, b"name", reader)?;
                    if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") {
                        hidden = true;
                    }
                }
                b"xfrm" => parse_xfrm(reader, &mut tf)?,
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"pic" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    if hidden {
        return Ok(None);
    }
    let rid = match rel_id {
        Some(r) => r,
        None => return Ok(None),
    };
    let rel = match rels.find_by_id(&rid) {
        Some(r) if r.rel_type == REL_IMAGE => r,
        _ => return Ok(None),
    };
    let image_path = normalize(&rel.target, slide_path);
    let logical = ensure_image(package, &image_path, images);
    Ok(Some(SlideElement::Picture(Picture {
        image_id: logical,
        name,
        x: tf.x,
        y: tf.y,
        width: tf.w,
        height: tf.h,
        rotation: tf.rot,
    })))
}

fn ensure_image(package: &OpcPackage, path: &str, images: &mut Vec<SlidesImage>) -> String {
    if images.iter().any(|i| i.id == path) {
        return path.to_string();
    }
    let bytes = package.parts.get(path).cloned().unwrap_or_default();
    let ext = path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
    let mime = match ext.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "svg" => "image/svg+xml",
        _ => "application/octet-stream",
    }
    .to_string();
    images.push(SlidesImage {
        id: path.to_string(),
        mime,
        data_base64: b64::encode(&bytes),
    });
    path.to_string()
}

fn parse_cxn(reader: &mut Reader<&[u8]>, theme: Option<&Theme>) -> PboResult<Option<SlideElement>> {
    let mut tf = Xf::default();
    let mut prst = String::new();
    let mut outline = None;
    let mut hidden = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"cNvPr" => {}
                b"spPr" => {
                    let mut maybe_prst = if prst.is_empty() {
                        None
                    } else {
                        Some(prst.clone())
                    };
                    let mut fill = None;
                    parse_shape_props(
                        reader,
                        &mut tf,
                        &mut maybe_prst,
                        &mut fill,
                        &mut outline,
                        theme,
                    )?;
                    if let Some(parsed_prst) = maybe_prst {
                        prst = parsed_prst;
                    }
                }
                b"prstGeom" if let Some(p) = attr_local(&e, b"prst", reader)? => prst = p,
                b"prstGeom" => {}
                b"xfrm" => parse_xfrm(reader, &mut tf)?,
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"spPr" => {} // empty <p:spPr/> - nothing to parse
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"cNvPr" => {}
                b"prstGeom" if let Some(p) = attr_local(&e, b"prst", reader)? => prst = p,
                b"prstGeom" => {}
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"cxnSp" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    if hidden {
        return Ok(None);
    }
    let connector_type = if prst.starts_with("bent") {
        ConnectorType::Bent
    } else if prst.starts_with("curved") {
        ConnectorType::Curved
    } else {
        ConnectorType::Straight
    };
    Ok(Some(SlideElement::Connector(Connector {
        connector_type,
        x: tf.x,
        y: tf.y,
        width: tf.w,
        height: tf.h,
        flip_h: tf.flip_h,
        flip_v: tf.flip_v,
        rotation: tf.rot,
        outline,
    })))
}

#[derive(Default)]
struct GroupXf {
    off: Xf,
    ch_off_x: f64,
    ch_off_y: f64,
    ch_ext_cx: f64,
    ch_ext_cy: f64,
}

fn parse_group(
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    slide_path: &str,
    package: &OpcPackage,
    images: &mut Vec<SlidesImage>,
    theme: Option<&Theme>,
) -> PboResult<Option<GroupShape>> {
    let mut gxf = GroupXf::default();
    let mut elements = Vec::new();
    let mut hidden = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"grpSpPr" => parse_group_sp_pr(reader, &mut gxf)?,
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"cNvPr" => {}
                b"sp" => {
                    if let Some(el) = parse_sp(reader, theme)? {
                        elements.push(el);
                    }
                }
                b"pic" => {
                    if let Some(el) = parse_pic(reader, rels, slide_path, package, images)? {
                        elements.push(el);
                    }
                }
                b"cxnSp" => {
                    if let Some(el) = parse_cxn(reader, theme)? {
                        elements.push(el);
                    }
                }
                b"grpSp" => {
                    elements.extend(
                        parse_group(reader, rels, slide_path, package, images, theme)?
                            .map(SlideElement::Group),
                    );
                }
                _ => {}
            },
            Event::Empty(e)
                if local_name(e.name().as_ref()) == b"cNvPr"
                    && attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") =>
            {
                hidden = true;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"grpSp" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    if hidden {
        return Ok(None);
    }
    Ok(Some(GroupShape {
        x: gxf.off.x,
        y: gxf.off.y,
        width: gxf.off.w,
        height: gxf.off.h,
        ch_off_x: gxf.ch_off_x,
        ch_off_y: gxf.ch_off_y,
        ch_ext_cx: gxf.ch_ext_cx,
        ch_ext_cy: gxf.ch_ext_cy,
        elements,
    }))
}

fn parse_group_sp_pr(reader: &mut Reader<&[u8]>, gxf: &mut GroupXf) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"xfrm" => {
                if let Some(r) = attr_local(&e, b"rot", reader)?.and_then(|v| v.parse().ok()) {
                    gxf.off.rot = r;
                }
                if attr_local(&e, b"flipH", reader)?.as_deref() == Some("1") {
                    gxf.off.flip_h = true;
                }
                if attr_local(&e, b"flipV", reader)?.as_deref() == Some("1") {
                    gxf.off.flip_v = true;
                }
            }
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"off" => {
                    if let Some(x) = attr_i64(&e, b"x", reader)? {
                        gxf.off.x = x as f64;
                    }
                    if let Some(y) = attr_i64(&e, b"y", reader)? {
                        gxf.off.y = y as f64;
                    }
                }
                b"ext" => {
                    if let Some(cx) = attr_i64(&e, b"cx", reader)? {
                        gxf.off.w = cx as f64;
                    }
                    if let Some(cy) = attr_i64(&e, b"cy", reader)? {
                        gxf.off.h = cy as f64;
                    }
                }
                b"chOff" => {
                    if let Some(x) = attr_i64(&e, b"x", reader)? {
                        gxf.ch_off_x = x as f64;
                    }
                    if let Some(y) = attr_i64(&e, b"y", reader)? {
                        gxf.ch_off_y = y as f64;
                    }
                }
                b"chExt" => {
                    if let Some(cx) = attr_i64(&e, b"cx", reader)? {
                        gxf.ch_ext_cx = cx as f64;
                    }
                    if let Some(cy) = attr_i64(&e, b"cy", reader)? {
                        gxf.ch_ext_cy = cy as f64;
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"grpSpPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn extract_notes_text(xml: &str) -> PboResult<String> {
    let mut r = Reader::from_str(xml);
    r.config_mut().trim_text(false);
    let mut out = String::new();
    let mut buf = Vec::new();
    let mut first = true;
    loop {
        match r.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"p" => {
                    if !first {
                        out.push('\n');
                    }
                    first = false;
                }
                b"t" => {
                    let mut b2 = Vec::new();
                    loop {
                        match r.read_event_into(&mut b2).map_err(xml_err)? {
                            Event::Text(t) => out
                                .push_str(&t.unescape().map_err(|e| PboError::Xml(e.to_string()))?),
                            Event::End(ee) if local_name(ee.name().as_ref()) == b"t" => break,
                            Event::Eof => break,
                            _ => {}
                        }
                        b2.clear();
                    }
                }
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn parse_core(xml: &str) -> PboResult<PresentationMetadata> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut md = PresentationMetadata::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => {
                let qname = e.name().as_ref().to_vec();
                let mut text = String::new();
                let mut inner = Vec::new();
                loop {
                    match reader.read_event_into(&mut inner).map_err(xml_err)? {
                        Event::Text(t) => {
                            text.push_str(&t.unescape().map_err(|e| PboError::Xml(e.to_string()))?)
                        }
                        Event::End(ee) if local_name(ee.name().as_ref()) == local_name(&qname) => {
                            break;
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                    inner.clear();
                }
                let t = if text.is_empty() { None } else { Some(text) };
                match local_name(&qname) {
                    b"title" => md.title = t,
                    b"creator" => md.author = t,
                    b"lastModifiedBy" => md.last_modified_by = t,
                    b"created" => md.created = t,
                    b"modified" => md.modified = t,
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(md)
}
