// PPTX reader: parses a .pptx OPC package into a SlidesDocument.
// Covers slides with text boxes, preset shapes, pictures, connectors,
// groups; notes slides; core metadata.

use pbo_common::b64;
use pbo_common::model::slides::{
    Connector, ConnectorType, CustomGeometry, Fill, GeometryAdjustment, GlowEffect, GradientStop,
    GroupShape, Outline, PatternFill, Picture, PictureCrop, PresentationMetadata, ShadowEffect,
    Shape, ShapeType, Slide, SlideElement, SlideLayout, SlideMaster, SlideTable, SlideTableCell,
    SlideTableRow, SlidesDocument, SlidesImage, SolidFill, TextAlign, TextBodyProperties, TextBox,
    TextParagraph, TextRun, Theme, ThemeColor, ThemeColorModifier, ThemeFillStop, ThemeFillStyle,
    VisualEffects,
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
        let theme = layout_idx
            .and_then(|idx| doc.layouts.get(idx))
            .and_then(|layout| layout.master_idx)
            .and_then(|idx| doc.masters.get(idx))
            .and_then(|master| master.theme.as_ref())
            .or(doc.theme.as_ref());
        let (mut slide, notes_path) = parse_slide(
            slide_xml,
            &slide_rels,
            &slide_path,
            &package,
            &mut doc.images,
            theme,
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
    let master_idx = if let Some(master_path) = master_path {
        ensure_master(&master_path, package, doc, master_indices)?
    } else {
        None
    };

    let theme = master_idx
        .and_then(|idx| doc.masters.get(idx))
        .and_then(|master| master.theme.as_ref())
        .or(doc.theme.as_ref());
    let mut layout = parse_layout(
        layout_xml,
        &layout_rels,
        layout_path,
        package,
        &mut doc.images,
        theme,
    )?;
    layout.master_idx = master_idx;

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
    let master_theme = if let Some(theme_path) = master_rels
        .0
        .iter()
        .find(|rel| rel.rel_type == REL_THEME)
        .map(|rel| normalize(&rel.target, master_path))
        && let Ok(theme_xml) = package.get_part_str(&theme_path)
    {
        let theme = parse_theme(theme_xml)?;
        if doc.theme.is_none() {
            doc.theme = Some(theme.clone());
        }
        Some(theme)
    } else {
        None
    };
    let theme = master_theme.as_ref().or(doc.theme.as_ref());
    let mut master = parse_master(
        master_xml,
        &master_rels,
        master_path,
        package,
        &mut doc.images,
        theme,
    )?;
    master.theme = master_theme;

    let idx = doc.masters.len();
    doc.masters.push(master);
    master_indices.insert(master_path.to_owned(), idx);
    Ok(Some(idx))
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
    let mapped_theme = theme_with_color_map(xml, theme)?;
    let theme = mapped_theme.as_ref().or(theme);
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut slide = Slide::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"bgPr" => {
                slide.background = parse_bg_pr(&mut reader, theme)?;
            }
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
    let mapped_theme = theme_with_color_map(xml, theme)?;
    let theme = mapped_theme.as_ref().or(theme);
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
    let mapped_theme = theme_with_color_map(xml, theme)?;
    let theme = mapped_theme.as_ref().or(theme);
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
                } else if name == b"fillStyleLst" {
                    theme.fill_styles = parse_theme_fill_style_list(&mut reader)?;
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

fn parse_theme_fill_style_list(reader: &mut Reader<&[u8]>) -> PboResult<Vec<ThemeFillStyle>> {
    let mut styles = Vec::new();
    let mut idx = 1;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"solidFill" => {
                    let stops = parse_theme_solid_fill(reader)?;
                    styles.push(ThemeFillStyle {
                        idx,
                        is_gradient: false,
                        stops,
                    });
                    idx += 1;
                }
                b"gradFill" => {
                    let stops = parse_theme_gradient_fill(reader)?;
                    styles.push(ThemeFillStyle {
                        idx,
                        is_gradient: true,
                        stops,
                    });
                    idx += 1;
                }
                b"blipFill" | b"pattFill" => {
                    let tag = local_name(e.name().as_ref()).to_vec();
                    skip_to_end(reader, &tag)?;
                    styles.push(ThemeFillStyle {
                        idx,
                        is_gradient: false,
                        stops: Vec::new(),
                    });
                    idx += 1;
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"solidFill" | b"gradFill" | b"blipFill" | b"pattFill" => {
                    styles.push(ThemeFillStyle {
                        idx,
                        is_gradient: local_name(e.name().as_ref()) == b"gradFill",
                        stops: Vec::new(),
                    });
                    idx += 1;
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"fillStyleLst" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(styles)
}

fn parse_theme_solid_fill(reader: &mut Reader<&[u8]>) -> PboResult<Vec<ThemeFillStop>> {
    let mut stop = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"srgbClr" => {
                    stop = Some(parse_theme_color_ref(reader, b"srgbClr", &e, 0)?);
                }
                b"schemeClr" => {
                    stop = Some(parse_theme_color_ref(reader, b"schemeClr", &e, 0)?);
                }
                b"prstClr" => {
                    stop = Some(parse_theme_color_ref(reader, b"prstClr", &e, 0)?);
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"srgbClr" => stop = Some(theme_color_ref_from_empty(&e, reader, 0)?),
                b"schemeClr" => stop = Some(theme_color_ref_from_empty(&e, reader, 0)?),
                b"prstClr" => stop = Some(theme_color_ref_from_empty(&e, reader, 0)?),
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"solidFill" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(stop.into_iter().collect())
}

fn parse_theme_gradient_fill(reader: &mut Reader<&[u8]>) -> PboResult<Vec<ThemeFillStop>> {
    let mut stops = Vec::new();
    let mut current_pos = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"gs" => current_pos = attr_local(&e, b"pos", reader)?.and_then(|v| v.parse().ok()),
                b"srgbClr" => {
                    if let Some(pos) = current_pos {
                        stops.push(parse_theme_color_ref(reader, b"srgbClr", &e, pos)?);
                    }
                }
                b"schemeClr" => {
                    if let Some(pos) = current_pos {
                        stops.push(parse_theme_color_ref(reader, b"schemeClr", &e, pos)?);
                    }
                }
                b"prstClr" => {
                    if let Some(pos) = current_pos {
                        stops.push(parse_theme_color_ref(reader, b"prstClr", &e, pos)?);
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"srgbClr" | b"schemeClr" | b"prstClr" => {
                    if let Some(pos) = current_pos {
                        stops.push(theme_color_ref_from_empty(&e, reader, pos)?);
                    }
                }
                _ => {}
            },
            Event::End(e) => match local_name(e.name().as_ref()) {
                b"gs" => current_pos = None,
                b"gradFill" => break,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(stops)
}

fn parse_theme_color_ref(
    reader: &mut Reader<&[u8]>,
    start_tag: &[u8],
    e: &quick_xml::events::BytesStart<'_>,
    position: u32,
) -> PboResult<ThemeFillStop> {
    let mut stop = theme_color_ref_from_empty(e, reader, position)?;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(child) => {
                let child_name = child.name();
                let name = local_name(child_name.as_ref());
                if name == b"alpha" {
                    stop.alpha = attr_local(&child, b"val", reader)?.and_then(|v| v.parse().ok());
                } else if let Some(modifier) = color_modifier_name(name)
                    && let Some(value) =
                        attr_local(&child, b"val", reader)?.and_then(|v| v.parse().ok())
                {
                    stop.modifiers.push(ThemeColorModifier {
                        name: modifier.to_string(),
                        value,
                    });
                }
            }
            Event::End(child) if local_name(child.name().as_ref()) == start_tag => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(stop)
}

fn theme_color_ref_from_empty(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<&[u8]>,
    position: u32,
) -> PboResult<ThemeFillStop> {
    let mut stop = ThemeFillStop {
        position,
        ..Default::default()
    };
    match local_name(e.name().as_ref()) {
        b"srgbClr" => stop.color = attr_local(e, b"val", reader)?,
        b"schemeClr" => stop.role = attr_local(e, b"val", reader)?,
        b"prstClr" => {
            stop.color =
                attr_local(e, b"val", reader)?.and_then(|role| resolve_preset_color(&role));
        }
        _ => {}
    }
    Ok(stop)
}

fn skip_to_end(reader: &mut Reader<&[u8]>, end_tag: &[u8]) -> PboResult<()> {
    let mut depth = 1usize;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == end_tag => depth += 1,
            Event::End(e) if local_name(e.name().as_ref()) == end_tag => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
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
    if let Some(color) = theme.colors.iter().find(|color| color.role == role) {
        return Some(color.value.clone());
    }
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

fn theme_with_color_map(xml: &str, theme: Option<&Theme>) -> PboResult<Option<Theme>> {
    let Some(theme) = theme else {
        return Ok(None);
    };
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut mapped = theme.clone();
    let mut found = false;
    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e)
                if matches!(
                    local_name(e.name().as_ref()),
                    b"overrideClrMapping" | b"clrMap"
                ) =>
            {
                for attr in e.attributes().with_checks(false).filter_map(Result::ok) {
                    let logical =
                        String::from_utf8_lossy(local_name(attr.key.as_ref())).into_owned();
                    let Ok(actual) = attr.decode_and_unescape_value(reader.decoder()) else {
                        continue;
                    };
                    if let Some(value) = resolve_theme_color(Some(theme), actual.as_ref()) {
                        if let Some(color) =
                            mapped.colors.iter_mut().find(|color| color.role == logical)
                        {
                            color.value = value;
                        } else {
                            mapped.colors.push(ThemeColor {
                                role: logical,
                                value,
                            });
                        }
                        found = true;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(found.then_some(mapped))
}

/// Apply color modifiers in OOXML order, following the C++ `ColorBase::ApplyModifiers` behavior.
fn apply_color_modifiers(hex: &str, modifiers: &[(String, i64)]) -> String {
    let mut r = u8::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let mut g = u8::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let mut b = u8::from_str_radix(&hex[4..6], 16).unwrap_or(0);

    for (name, raw) in modifiers {
        let val = *raw as f64 / 100000.0;
        match name.as_str() {
            "red" => r = if val > 1.0 { 255 } else { (255.0 * val) as u8 },
            "green" => g = if val > 1.0 { 255 } else { (255.0 * val) as u8 },
            "blue" => b = if val > 1.0 { 255 } else { (255.0 * val) as u8 },
            "redMod" => r = ((r as f64 * val).min(255.0)) as u8,
            "greenMod" => g = ((g as f64 * val).min(255.0)) as u8,
            "blueMod" => b = ((b as f64 * val).min(255.0)) as u8,
            "redOff" => r = (r as f64 + val * 255.0).clamp(0.0, 255.0) as u8,
            "greenOff" => g = (g as f64 + val * 255.0).clamp(0.0, 255.0) as u8,
            "blueOff" => b = (b as f64 + val * 255.0).clamp(0.0, 255.0) as u8,
            "lumMod" => {
                let (h, s, mut l) = rgb_to_hsl240(r, g, b);
                l = ((l as f64 * val).min(240.0)) as u8;
                (r, g, b) = hsl240_to_rgb(h, s, l);
            }
            "lumOff" => {
                let (h, s, l) = rgb_to_hsl240(r, g, b);
                let l = (l as f64 + val * 240.0).clamp(0.0, 240.0) as u8;
                (r, g, b) = hsl240_to_rgb(h, s, l);
            }
            "satMod" => {
                let (h, mut s, l) = rgb_to_hsl240(r, g, b);
                s = ((s as f64 * val).min(240.0)) as u8;
                (r, g, b) = hsl240_to_rgb(h, s, l);
            }
            "satOff" => {
                let (h, s, l) = rgb_to_hsl240(r, g, b);
                let s = (s as f64 + val * 240.0).clamp(0.0, 240.0) as u8;
                (r, g, b) = hsl240_to_rgb(h, s, l);
            }
            "shade" => {
                r = (r as f64 * val) as u8;
                g = (g as f64 * val) as u8;
                b = (b as f64 * val) as u8;
            }
            "tint" => {
                r = (255.0 - (255.0 - r as f64) * val) as u8;
                g = (255.0 - (255.0 - g as f64) * val) as u8;
                b = (255.0 - (255.0 - b as f64) * val) as u8;
            }
            "inv" => {
                r ^= 0xFF;
                g ^= 0xFF;
                b ^= 0xFF;
            }
            _ => {}
        }
    }

    format!("{r:02X}{g:02X}{b:02X}")
}

fn color_modifier_name(name: &[u8]) -> Option<&'static str> {
    match name {
        b"red" => Some("red"),
        b"green" => Some("green"),
        b"blue" => Some("blue"),
        b"redMod" => Some("redMod"),
        b"greenMod" => Some("greenMod"),
        b"blueMod" => Some("blueMod"),
        b"redOff" => Some("redOff"),
        b"greenOff" => Some("greenOff"),
        b"blueOff" => Some("blueOff"),
        b"lumMod" => Some("lumMod"),
        b"lumOff" => Some("lumOff"),
        b"satMod" => Some("satMod"),
        b"satOff" => Some("satOff"),
        b"shade" => Some("shade"),
        b"tint" => Some("tint"),
        b"inv" => Some("inv"),
        _ => None,
    }
}

fn rgb_to_hsl240(r: u8, g: u8, b: u8) -> (u8, u8, u8) {
    const ONE_THIRD: f64 = 1.0 / 3.0;
    const TWO_THIRDS: f64 = 2.0 / 3.0;
    let min = r.min(g).min(b) as i32;
    let max = r.max(g).max(b) as i32;
    let delta = max - min;
    let d_max = (max + min) as f64 / 255.0;
    let mut d_delta = delta as f64 / 255.0;
    let l = d_max / 2.0;
    let (mut h, s) = if delta == 0 {
        (0.0, 0.0)
    } else {
        let s = if l < 0.5 {
            d_delta / d_max
        } else {
            d_delta / (2.0 - d_max)
        };
        d_delta *= 1530.0;
        let dr = (max - r as i32) as f64 / d_delta;
        let dg = (max - g as i32) as f64 / d_delta;
        let db = (max - b as i32) as f64 / d_delta;
        let h = if r as i32 == max {
            db - dg
        } else if g as i32 == max {
            ONE_THIRD + dr - db
        } else {
            TWO_THIRDS + dg - dr
        };
        (h, s)
    };
    if h < 0.0 {
        h += 1.0;
    }
    if h > 1.0 {
        h -= 1.0;
    }
    ((h * 240.0) as u8, (s * 240.0) as u8, (l * 240.0) as u8)
}

fn hsl240_to_rgb(h: u8, s: u8, l: u8) -> (u8, u8, u8) {
    const ONE_THIRD: f64 = 1.0 / 3.0;
    let h = h as f64 / 240.0;
    let s = s as f64 / 240.0;
    let l = l as f64 / 240.0;
    let v2 = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - s * l
    };
    let v1 = 2.0 * l - v2;
    (
        (255.0 * hue_to_rgb(v1, v2, h + ONE_THIRD)) as u8,
        (255.0 * hue_to_rgb(v1, v2, h)) as u8,
        (255.0 * hue_to_rgb(v1, v2, h - ONE_THIRD)) as u8,
    )
}

fn hue_to_rgb(v1: f64, v2: f64, mut vh: f64) -> f64 {
    if vh < 0.0 {
        vh += 1.0;
    }
    if vh > 1.0 {
        vh -= 1.0;
    }
    if vh < 1.0 / 6.0 {
        return v1 + (v2 - v1) * 6.0 * vh;
    }
    if vh < 0.5 {
        return v2;
    }
    if vh < 2.0 / 3.0 {
        return v1 + (v2 - v1) * (2.0 / 3.0 - vh) * 6.0;
    }
    v1
}

#[derive(Debug, Clone, Default)]
struct ParsedColor {
    color: Option<String>,
    alpha: Option<i64>,
}

/// Resolve a color element (srgbClr or schemeClr with modifiers) reading child events.
fn resolve_color_with_modifiers_and_alpha(
    reader: &mut Reader<&[u8]>,
    start_tag: &[u8],
    base_color: &str,
    theme: Option<&Theme>,
) -> PboResult<ParsedColor> {
    let is_scheme = start_tag == b"schemeClr";
    let base = if is_scheme {
        resolve_theme_color(theme, base_color)
    } else {
        Some(base_color.to_string())
    };
    let Some(hex) = base else {
        return Ok(ParsedColor::default());
    };

    let mut alpha = None;
    let mut modifiers = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"alpha" => alpha = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
                name => {
                    if let Some(modifier) = color_modifier_name(name)
                        && let Some(value) =
                            attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok())
                    {
                        modifiers.push((modifier.to_string(), value));
                    }
                }
            },
            Event::End(e) if local_name(e.name().as_ref()) == start_tag => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    let color = if modifiers.is_empty() {
        hex
    } else {
        apply_color_modifiers(&hex, &modifiers)
    };
    Ok(ParsedColor {
        color: Some(color),
        alpha,
    })
}

fn resolve_color_with_modifiers(
    reader: &mut Reader<&[u8]>,
    start_tag: &[u8],
    base_color: &str,
    theme: Option<&Theme>,
) -> PboResult<Option<String>> {
    Ok(resolve_color_with_modifiers_and_alpha(reader, start_tag, base_color, theme)?.color)
}

fn resolve_color_from_empty(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<&[u8]>,
    theme: Option<&Theme>,
) -> PboResult<Option<String>> {
    match local_name(e.name().as_ref()) {
        b"srgbClr" => Ok(attr_local(e, b"val", reader)?),
        b"schemeClr" => {
            Ok(attr_local(e, b"val", reader)?.and_then(|role| resolve_theme_color(theme, &role)))
        }
        b"sysClr" => Ok(attr_local(e, b"lastClr", reader)?),
        b"prstClr" => {
            Ok(attr_local(e, b"val", reader)?.and_then(|role| resolve_preset_color(&role)))
        }
        _ => Ok(None),
    }
}

fn resolve_preset_color(role: &str) -> Option<String> {
    match role {
        "black" => Some("000000".into()),
        "white" => Some("FFFFFF".into()),
        "red" => Some("FF0000".into()),
        "green" => Some("008000".into()),
        "blue" => Some("0000FF".into()),
        "yellow" => Some("FFFF00".into()),
        "cyan" => Some("00FFFF".into()),
        "magenta" => Some("FF00FF".into()),
        _ => None,
    }
}

fn parse_color_container(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
    theme: Option<&Theme>,
) -> PboResult<(Option<String>, Option<i64>)> {
    let mut color = None;
    let mut alpha = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"srgbClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)? {
                        let parsed = resolve_color_with_modifiers_and_alpha(
                            reader, b"srgbClr", &val, theme,
                        )?;
                        color = parsed.color;
                        alpha = alpha.or(parsed.alpha);
                    }
                }
                b"schemeClr" => {
                    if let Some(role) = attr_local(&e, b"val", reader)? {
                        let parsed = resolve_color_with_modifiers_and_alpha(
                            reader,
                            b"schemeClr",
                            &role,
                            theme,
                        )?;
                        color = parsed.color;
                        alpha = alpha.or(parsed.alpha);
                    }
                }
                b"sysClr" => color = attr_local(&e, b"lastClr", reader)?,
                b"prstClr" => {
                    if let Some(role) = attr_local(&e, b"val", reader)? {
                        color = resolve_preset_color(&role);
                    }
                }
                b"alpha" => alpha = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
                _ => {}
            },
            Event::Empty(e)
                if matches!(
                    local_name(e.name().as_ref()),
                    b"srgbClr" | b"schemeClr" | b"sysClr" | b"prstClr"
                ) && color.is_none() =>
            {
                color = resolve_color_from_empty(&e, reader, theme)?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"alpha" => {
                alpha = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok());
            }
            Event::End(e) if local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok((color, alpha))
}

fn parse_body_properties(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    empty: bool,
) -> PboResult<TextBodyProperties> {
    let mut body = TextBodyProperties {
        inset_left: attr_i64(start, b"lIns", reader)?.map(|v| v as f64),
        inset_top: attr_i64(start, b"tIns", reader)?.map(|v| v as f64),
        inset_right: attr_i64(start, b"rIns", reader)?.map(|v| v as f64),
        inset_bottom: attr_i64(start, b"bIns", reader)?.map(|v| v as f64),
        anchor: attr_local(start, b"anchor", reader)?,
        rotation: attr_local(start, b"rot", reader)?.and_then(|v| v.parse().ok()),
        vertical: attr_local(start, b"vert", reader)?,
        wrap: attr_local(start, b"wrap", reader)?,
        vertical_overflow: attr_local(start, b"vertOverflow", reader)?,
        horizontal_overflow: attr_local(start, b"horzOverflow", reader)?,
    };
    if empty {
        return Ok(body);
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()) == b"prstTxWarp" && body.vertical.is_none() =>
            {
                body.vertical = attr_local(&e, b"prst", reader)?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"bodyPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(body)
}

fn parse_pattern_fill(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    empty: bool,
    theme: Option<&Theme>,
) -> PboResult<Option<Fill>> {
    let mut pattern = PatternFill {
        preset: attr_local(start, b"prst", reader)?,
        ..Default::default()
    };
    if empty {
        return Ok(Some(Fill::Pattern(pattern)));
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"fgClr" => {
                pattern.foreground = parse_color_container(reader, b"fgClr", theme)?.0;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"bgClr" => {
                pattern.background = parse_color_container(reader, b"bgClr", theme)?.0;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"pattFill" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(Some(Fill::Pattern(pattern)))
}

fn parse_effects(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
    theme: Option<&Theme>,
) -> PboResult<VisualEffects> {
    let mut effects = VisualEffects::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"outerShdw" => {
                effects.shadow = Some(parse_shadow_effect(reader, &e, false, theme)?);
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"glow" => {
                effects.glow = Some(parse_glow_effect(reader, &e, false, theme)?);
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"outerShdw" => {
                effects.shadow = Some(parse_shadow_effect(reader, &e, true, theme)?);
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"glow" => {
                effects.glow = Some(parse_glow_effect(reader, &e, true, theme)?);
            }
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"softEdge" => {
                effects.soft_edge_radius = attr_i64(&e, b"rad", reader)?.map(|v| v as f64);
            }
            Event::End(e) if local_name(e.name().as_ref()) == end_tag => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(effects)
}

fn parse_shadow_effect(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    empty: bool,
    theme: Option<&Theme>,
) -> PboResult<ShadowEffect> {
    let mut shadow = ShadowEffect {
        blur_radius: attr_i64(start, b"blurRad", reader)?.unwrap_or_default() as f64,
        distance: attr_i64(start, b"dist", reader)?.unwrap_or_default() as f64,
        direction: attr_local(start, b"dir", reader)?
            .and_then(|v| v.parse().ok())
            .unwrap_or(2_700_000),
        color: None,
        alpha: None,
    };
    if empty {
        return Ok(shadow);
    }
    let (color, alpha) = parse_color_container(reader, b"outerShdw", theme)?;
    shadow.color = color;
    shadow.alpha = alpha;
    Ok(shadow)
}

fn parse_glow_effect(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    empty: bool,
    theme: Option<&Theme>,
) -> PboResult<GlowEffect> {
    let mut glow = GlowEffect {
        radius: attr_i64(start, b"rad", reader)?.unwrap_or_default() as f64,
        color: None,
        alpha: None,
    };
    if empty {
        return Ok(glow);
    }
    let (color, alpha) = parse_color_container(reader, b"glow", theme)?;
    glow.color = color;
    glow.alpha = alpha;
    Ok(glow)
}

fn parse_crop_rect_empty(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<&[u8]>,
    crop: &mut PictureCrop,
) -> PboResult<()> {
    if let Some(v) = attr_local(e, b"l", reader)?.and_then(|v| v.parse().ok()) {
        crop.left = v;
    }
    if let Some(v) = attr_local(e, b"t", reader)?.and_then(|v| v.parse().ok()) {
        crop.top = v;
    }
    if let Some(v) = attr_local(e, b"r", reader)?.and_then(|v| v.parse().ok()) {
        crop.right = v;
    }
    if let Some(v) = attr_local(e, b"b", reader)?.and_then(|v| v.parse().ok()) {
        crop.bottom = v;
    }
    Ok(())
}

#[derive(Debug, Clone, Default)]
struct RunStyle {
    bold: Option<bool>,
    italic: Option<bool>,
    size: Option<f32>,
    color: Option<String>,
    font_family: Option<String>,
    underline: Option<bool>,
    strike: Option<bool>,
    baseline: Option<i32>,
}

impl RunStyle {
    fn merge(&mut self, other: RunStyle) {
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.size.is_some() {
            self.size = other.size;
        }
        if other.color.is_some() {
            self.color = other.color;
        }
        if other.font_family.is_some() {
            self.font_family = other.font_family;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.strike.is_some() {
            self.strike = other.strike;
        }
        if other.baseline.is_some() {
            self.baseline = other.baseline;
        }
    }
}

fn parse_run_style_attrs(
    start: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<&[u8]>,
) -> PboResult<RunStyle> {
    Ok(RunStyle {
        bold: attr_local(start, b"b", reader)?.map(|value| value == "1"),
        italic: attr_local(start, b"i", reader)?.map(|value| value == "1"),
        size: attr_local(start, b"sz", reader)?
            .and_then(|value| value.parse::<f32>().ok())
            .map(|value| value / 100.0),
        color: None,
        font_family: None,
        underline: attr_local(start, b"u", reader)?.map(|value| value != "none"),
        strike: attr_local(start, b"strike", reader)?.map(|value| value != "noStrike"),
        baseline: attr_local(start, b"baseline", reader)?.and_then(|value| value.parse().ok()),
    })
}

fn apply_run_style_attrs(
    start: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<&[u8]>,
    style: &mut RunStyle,
) -> PboResult<()> {
    style.merge(parse_run_style_attrs(start, reader)?);
    Ok(())
}

fn parse_default_run_properties(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    empty: bool,
    theme: Option<&Theme>,
) -> PboResult<RunStyle> {
    let mut style = parse_run_style_attrs(start, reader)?;
    if empty {
        return Ok(style);
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"solidFill" => {
                    if let Some(fill) = parse_solid_fill(reader, theme)?
                        && let Some(color) = solid_fill_color(&fill)
                    {
                        style.color = Some(color.to_owned());
                    }
                }
                b"gradFill" => {
                    if let Some(Fill::Gradient(stops)) = parse_gradient_fill(reader, theme)?
                        && let Some(stop) = stops.first()
                    {
                        style.color = Some(stop.color.clone());
                    }
                }
                b"latin" | b"ea" | b"cs" if style.font_family.is_none() => {
                    style.font_family = attr_local(&e, b"typeface", reader)?;
                }
                b"srgbClr" => {
                    if style.color.is_none()
                        && let Some(value) = attr_local(&e, b"val", reader)?
                    {
                        style.color = Some(value);
                    }
                }
                b"schemeClr" => {
                    if style.color.is_none()
                        && let Some(role) = attr_local(&e, b"val", reader)?
                    {
                        style.color = resolve_theme_color(theme, &role);
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"latin" | b"ea" | b"cs" if style.font_family.is_none() => {
                    style.font_family = attr_local(&e, b"typeface", reader)?;
                }
                b"srgbClr" => {
                    if style.color.is_none()
                        && let Some(value) = attr_local(&e, b"val", reader)?
                    {
                        style.color = Some(value);
                    }
                }
                b"schemeClr" => {
                    if style.color.is_none()
                        && let Some(role) = attr_local(&e, b"val", reader)?
                    {
                        style.color = resolve_theme_color(theme, &role);
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"defRPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(style)
}

fn merge_effects(target: &mut Option<VisualEffects>, parsed: VisualEffects) {
    if parsed.shadow.is_some() || parsed.glow.is_some() || parsed.soft_edge_radius.is_some() {
        *target = Some(parsed);
    }
}

#[allow(clippy::too_many_arguments)]
fn apply_style_ref_color(
    color: String,
    alpha: Option<i64>,
    in_fill_ref: bool,
    fill_ref_idx: Option<u32>,
    in_ln_ref: bool,
    ln_ref_idx: Option<u32>,
    theme: Option<&Theme>,
    style_fill: &mut Option<Fill>,
    style_outline: &mut Option<Outline>,
) {
    if in_fill_ref && fill_ref_idx.is_some_and(|i| i > 0) {
        *style_fill = resolve_theme_fill_style(theme, fill_ref_idx.unwrap_or(0), &color, alpha)
            .or_else(|| {
                Some(match alpha {
                    Some(alpha) => Fill::SolidAlpha(SolidFill {
                        color: color.clone(),
                        alpha: Some(alpha),
                    }),
                    None => Fill::Solid(color.clone()),
                })
            });
    }
    if in_ln_ref && ln_ref_idx.is_some_and(|i| i > 0) {
        *style_outline = Some(Outline {
            width_emu: 12700,
            color: Some(color),
            alpha,
            dash: None,
            head_end: None,
            tail_end: None,
            ..Default::default()
        });
    }
}

fn resolve_theme_fill_style(
    theme: Option<&Theme>,
    idx: u32,
    ph_color: &str,
    alpha: Option<i64>,
) -> Option<Fill> {
    let theme = theme?;
    let style = theme.fill_styles.iter().find(|style| style.idx == idx)?;
    if style.stops.is_empty() {
        return None;
    }

    if style.is_gradient || style.stops.len() > 1 {
        let stops = style
            .stops
            .iter()
            .filter_map(|stop| {
                resolve_theme_fill_stop(theme, stop, ph_color).map(|color| GradientStop {
                    position: stop.position,
                    color,
                })
            })
            .collect::<Vec<_>>();
        if stops.is_empty() {
            None
        } else {
            Some(Fill::Gradient(stops))
        }
    } else {
        let stop = style.stops.first()?;
        let color = resolve_theme_fill_stop(theme, stop, ph_color)?;
        match stop.alpha.or(alpha) {
            Some(alpha) => Some(Fill::SolidAlpha(SolidFill {
                color,
                alpha: Some(alpha),
            })),
            None => Some(Fill::Solid(color)),
        }
    }
}

fn resolve_theme_fill_stop(theme: &Theme, stop: &ThemeFillStop, ph_color: &str) -> Option<String> {
    let mut color = if let Some(color) = &stop.color {
        Some(color.clone())
    } else if stop.role.as_deref() == Some("phClr") {
        Some(ph_color.to_owned())
    } else if let Some(role) = &stop.role {
        resolve_theme_color(Some(theme), role)
    } else {
        None
    }?;

    if !stop.modifiers.is_empty() {
        let modifiers = stop
            .modifiers
            .iter()
            .map(|modifier| (modifier.name.clone(), modifier.value))
            .collect::<Vec<_>>();
        color = apply_color_modifiers(&color, &modifiers);
    }
    Some(color)
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
                b"schemeClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)? {
                        let parsed = resolve_color_with_modifiers_and_alpha(
                            reader,
                            b"schemeClr",
                            &val,
                            theme,
                        )?;
                        if let Some(color) = parsed.color {
                            apply_style_ref_color(
                                color,
                                parsed.alpha,
                                in_fill_ref,
                                fill_ref_idx,
                                in_ln_ref,
                                ln_ref_idx,
                                theme,
                                style_fill,
                                style_outline,
                            );
                        }
                    }
                }
                b"srgbClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)? {
                        let parsed = resolve_color_with_modifiers_and_alpha(
                            reader, b"srgbClr", &val, theme,
                        )?;
                        if let Some(color) = parsed.color {
                            apply_style_ref_color(
                                color,
                                parsed.alpha,
                                in_fill_ref,
                                fill_ref_idx,
                                in_ln_ref,
                                ln_ref_idx,
                                theme,
                                style_fill,
                                style_outline,
                            );
                        }
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"schemeClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)?
                        && let Some(resolved) = resolve_theme_color(theme, &val)
                    {
                        apply_style_ref_color(
                            resolved,
                            None,
                            in_fill_ref,
                            fill_ref_idx,
                            in_ln_ref,
                            ln_ref_idx,
                            theme,
                            style_fill,
                            style_outline,
                        );
                    }
                }
                b"srgbClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)? {
                        apply_style_ref_color(
                            val,
                            None,
                            in_fill_ref,
                            fill_ref_idx,
                            in_ln_ref,
                            ln_ref_idx,
                            theme,
                            style_fill,
                            style_outline,
                        );
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
                b"graphicFrame" => {
                    if let Some(el) =
                        parse_graphic_frame(reader, rels, slide_path, package, images, theme)?
                    {
                        out.push(el);
                    }
                }
                b"AlternateContent" => {
                    out.extend(parse_alternate_content(
                        reader, rels, slide_path, package, images, theme,
                    )?);
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

fn parse_alternate_content(
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    slide_path: &str,
    package: &OpcPackage,
    images: &mut Vec<SlidesImage>,
    theme: Option<&Theme>,
) -> PboResult<Vec<SlideElement>> {
    let mut choice: Option<Vec<SlideElement>> = None;
    let mut fallback: Option<Vec<SlideElement>> = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"Choice" => {
                let elements = parse_alternate_branch(
                    reader, b"Choice", rels, slide_path, package, images, theme,
                )?;
                if choice.as_ref().is_none_or(Vec::is_empty) && !elements.is_empty() {
                    choice = Some(elements);
                }
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"Fallback" => {
                let elements = parse_alternate_branch(
                    reader,
                    b"Fallback",
                    rels,
                    slide_path,
                    package,
                    images,
                    theme,
                )?;
                if fallback.as_ref().is_none_or(Vec::is_empty) && !elements.is_empty() {
                    fallback = Some(elements);
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"AlternateContent" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(choice
        .filter(|elements| !elements.is_empty())
        .or(fallback)
        .unwrap_or_default())
}

fn parse_alternate_branch(
    reader: &mut Reader<&[u8]>,
    end_tag: &[u8],
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
                b"graphicFrame" => {
                    if let Some(el) =
                        parse_graphic_frame(reader, rels, slide_path, package, images, theme)?
                    {
                        out.push(el);
                    }
                }
                b"AlternateContent" => {
                    out.extend(parse_alternate_content(
                        reader, rels, slide_path, package, images, theme,
                    )?);
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == end_tag => break,
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

fn parse_xfrm_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    reader: &Reader<&[u8]>,
    tf: &mut Xf,
) -> PboResult<()> {
    if let Some(r) = attr_local(e, b"rot", reader)?.and_then(|v| v.parse().ok()) {
        tf.rot = r;
    }
    if attr_local(e, b"flipH", reader)?.as_deref() == Some("1") {
        tf.flip_h = true;
    }
    if attr_local(e, b"flipV", reader)?.as_deref() == Some("1") {
        tf.flip_v = true;
    }
    Ok(())
}

fn parse_xfrm_start(
    reader: &mut Reader<&[u8]>,
    e: &quick_xml::events::BytesStart<'_>,
    tf: &mut Xf,
) -> PboResult<()> {
    parse_xfrm_attrs(e, reader, tf)?;
    parse_xfrm(reader, tf)
}

fn parse_xfrm(reader: &mut Reader<&[u8]>, tf: &mut Xf) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) => {
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
                    parse_xfrm_attrs(&e, reader, tf)?;
                    break;
                }
            }
            Event::Start(e) => {
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
                    parse_xfrm_attrs(&e, reader, tf)?;
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

#[allow(clippy::too_many_arguments)]
fn parse_shape_props(
    reader: &mut Reader<&[u8]>,
    tf: &mut Xf,
    prst: &mut Option<String>,
    adjustments: &mut Vec<GeometryAdjustment>,
    fill: &mut Option<Fill>,
    outline: &mut Option<Outline>,
    custom_geometry: &mut Option<CustomGeometry>,
    effects: &mut Option<VisualEffects>,
    theme: Option<&Theme>,
) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"xfrm" => parse_xfrm_start(reader, &e, tf)?,
                b"prstGeom" => parse_preset_geometry(reader, &e, false, prst, adjustments)?,
                b"solidFill" => *fill = parse_solid_fill(reader, theme)?,
                b"gradFill" => *fill = parse_gradient_fill(reader, theme)?,
                b"pattFill" => *fill = parse_pattern_fill(reader, &e, false, theme)?,
                b"ln" => *outline = Some(parse_outline(reader, &e, theme)?),
                b"custGeom" => *custom_geometry = Some(parse_custom_geometry(reader)?),
                b"effectLst" => merge_effects(effects, parse_effects(reader, b"effectLst", theme)?),
                b"effectDag" => merge_effects(effects, parse_effects(reader, b"effectDag", theme)?),
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"xfrm" => parse_xfrm_attrs(&e, reader, tf)?,
                b"prstGeom" => parse_preset_geometry(reader, &e, true, prst, adjustments)?,
                b"custGeom" => *custom_geometry = Some(CustomGeometry::default()),
                b"noFill" => *fill = Some(Fill::None),
                b"pattFill" => *fill = parse_pattern_fill(reader, &e, true, theme)?,
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

fn parse_preset_geometry(
    reader: &mut Reader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    empty: bool,
    prst: &mut Option<String>,
    adjustments: &mut Vec<GeometryAdjustment>,
) -> PboResult<()> {
    *prst = attr_local(start, b"prst", reader)?;
    if empty {
        return Ok(());
    }

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"gd" => {
                let name = attr_local(&e, b"name", reader)?.unwrap_or_default();
                let formula = attr_local(&e, b"fmla", reader)?.unwrap_or_default();
                adjustments.push(GeometryAdjustment {
                    name,
                    value: preset_adjustment_value(&formula),
                    formula,
                });
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"prstGeom" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn preset_adjustment_value(formula: &str) -> Option<f64> {
    formula
        .strip_prefix("val ")
        .and_then(|value| value.parse::<f64>().ok())
}

fn paragraph_style_level(name: &[u8]) -> Option<u8> {
    if name.len() == 7 && name.starts_with(b"lvl") && name.ends_with(b"pPr") {
        return (name[3] as char).to_digit(10).map(|value| value as u8);
    }
    None
}

fn parse_custom_geometry(reader: &mut Reader<&[u8]>) -> PboResult<CustomGeometry> {
    let mut geometry = CustomGeometry::default();
    let mut current_path = String::new();
    let mut command: Option<&'static str> = None;
    let mut cubic_points = Vec::<(i64, i64)>::new();
    let mut quad_points = Vec::<(i64, i64)>::new();
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"path" => {
                    current_path.clear();
                    if let Some(w) = attr_i64(&e, b"w", reader)? {
                        geometry.width = w as f64;
                    }
                    if let Some(h) = attr_i64(&e, b"h", reader)? {
                        geometry.height = h as f64;
                    }
                }
                b"moveTo" => command = Some("M"),
                b"lnTo" => command = Some("L"),
                b"cubicBezTo" => {
                    command = Some("C");
                    cubic_points.clear();
                }
                b"quadBezTo" => {
                    command = Some("Q");
                    quad_points.clear();
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"pt" => {
                    let x = attr_i64(&e, b"x", reader)?.unwrap_or_default();
                    let y = attr_i64(&e, b"y", reader)?.unwrap_or_default();
                    match command {
                        Some("M") | Some("L") => {
                            let command = command.unwrap_or("L");
                            current_path.push_str(command);
                            current_path.push(' ');
                            current_path.push_str(&x.to_string());
                            current_path.push(' ');
                            current_path.push_str(&y.to_string());
                            current_path.push(' ');
                        }
                        Some("C") => {
                            cubic_points.push((x, y));
                            if cubic_points.len() == 3 {
                                current_path.push_str("C ");
                                for (px, py) in cubic_points.drain(..) {
                                    current_path.push_str(&px.to_string());
                                    current_path.push(' ');
                                    current_path.push_str(&py.to_string());
                                    current_path.push(' ');
                                }
                            }
                        }
                        Some("Q") => {
                            quad_points.push((x, y));
                            if quad_points.len() == 2 {
                                current_path.push_str("Q ");
                                for (px, py) in quad_points.drain(..) {
                                    current_path.push_str(&px.to_string());
                                    current_path.push(' ');
                                    current_path.push_str(&py.to_string());
                                    current_path.push(' ');
                                }
                            }
                        }
                        _ => {}
                    }
                }
                b"close" => current_path.push_str("Z "),
                _ => {}
            },
            Event::End(e) => match local_name(e.name().as_ref()) {
                b"moveTo" | b"lnTo" | b"cubicBezTo" | b"quadBezTo" => command = None,
                b"path" => {
                    let path = current_path.trim();
                    if !path.is_empty() {
                        geometry.paths.push(path.to_owned());
                    }
                    current_path.clear();
                }
                b"custGeom" => break,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    Ok(geometry)
}

fn parse_solid_fill(reader: &mut Reader<&[u8]>, theme: Option<&Theme>) -> PboResult<Option<Fill>> {
    let mut color = None;
    let mut alpha = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"srgbClr" => {
                    if let Some(val) = attr_local(&e, b"val", reader)? {
                        let parsed = resolve_color_with_modifiers_and_alpha(
                            reader, b"srgbClr", &val, theme,
                        )?;
                        color = parsed.color;
                        alpha = alpha.or(parsed.alpha);
                    }
                }
                b"schemeClr" => {
                    if let Some(role) = attr_local(&e, b"val", reader)? {
                        let parsed = resolve_color_with_modifiers_and_alpha(
                            reader,
                            b"schemeClr",
                            &role,
                            theme,
                        )?;
                        color = parsed.color;
                        alpha = alpha.or(parsed.alpha);
                    }
                }
                b"prstClr" => {
                    if let Some(role) = attr_local(&e, b"val", reader)? {
                        color = resolve_preset_color(&role);
                    }
                }
                b"alpha" => alpha = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
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
                b"prstClr" => {
                    if let Some(role) = attr_local(&e, b"val", reader)? {
                        color = resolve_preset_color(&role);
                    }
                }
                b"alpha" => alpha = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"solidFill" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(color.map(|color| match alpha {
        Some(alpha) => Fill::SolidAlpha(SolidFill {
            color,
            alpha: Some(alpha),
        }),
        None => Fill::Solid(color),
    }))
}

fn solid_fill_color(fill: &Fill) -> Option<&str> {
    match fill {
        Fill::Solid(color) => Some(color.as_str()),
        Fill::SolidAlpha(fill) => Some(fill.color.as_str()),
        _ => None,
    }
}

fn solid_fill_alpha(fill: &Fill) -> Option<i64> {
    match fill {
        Fill::SolidAlpha(fill) => fill.alpha,
        _ => None,
    }
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
        alpha: None,
        dash: None,
        head_end: None,
        tail_end: None,
        cap: attr_local(e, b"cap", reader)?,
        join: None,
        compound: attr_local(e, b"cmpd", reader)?,
        align: attr_local(e, b"algn", reader)?,
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
                if let Some(fill) = parse_solid_fill(reader, theme)?
                    && let Some(color) = solid_fill_color(&fill)
                {
                    outline.color = Some(color.to_owned());
                    outline.alpha = solid_fill_alpha(&fill);
                }
            }
            Event::Start(event) if local_name(event.name().as_ref()) == b"round" => {
                outline.join = Some("round".into());
            }
            Event::Start(event) if local_name(event.name().as_ref()) == b"bevel" => {
                outline.join = Some("bevel".into());
            }
            Event::Start(event) if local_name(event.name().as_ref()) == b"miter" => {
                outline.join = Some("miter".into());
            }
            Event::Empty(event) if local_name(event.name().as_ref()) == b"prstDash" => {
                outline.dash = attr_local(&event, b"val", reader)?;
            }
            Event::Empty(event) if local_name(event.name().as_ref()) == b"round" => {
                outline.join = Some("round".into());
            }
            Event::Empty(event) if local_name(event.name().as_ref()) == b"bevel" => {
                outline.join = Some("bevel".into());
            }
            Event::Empty(event) if local_name(event.name().as_ref()) == b"miter" => {
                outline.join = Some("miter".into());
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
    let mut adjustments = Vec::new();
    let mut fill = None;
    let mut outline = None;
    let mut custom_geometry = None;
    let mut text = String::new();
    let mut bold = false;
    let mut italic = false;
    let mut color: Option<String> = None;
    let mut size: Option<f32> = None;
    let mut align: Option<TextAlign> = None;
    let mut body = TextBodyProperties::default();
    let mut placeholder: Option<String> = None;
    let mut placeholder_idx: Option<String> = None;
    let mut paragraphs = Vec::new();
    let mut hidden = false;
    let mut is_text_box = false;
    let mut style_fill: Option<Fill> = None;
    let mut style_outline: Option<Outline> = None;
    let mut effects: Option<VisualEffects> = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"spPr" => parse_shape_props(
                    reader,
                    &mut tf,
                    &mut prst,
                    &mut adjustments,
                    &mut fill,
                    &mut outline,
                    &mut custom_geometry,
                    &mut effects,
                    theme,
                )?,
                b"xfrm" => parse_xfrm_start(reader, &e, &mut tf)?,
                b"prstGeom" => {
                    parse_preset_geometry(reader, &e, false, &mut prst, &mut adjustments)?
                }
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"cNvPr" => {}
                b"cNvSpPr" => {
                    is_text_box = attr_local(&e, b"txBox", reader)?.as_deref() == Some("1");
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
                        &mut body,
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
                b"prstGeom" => {
                    parse_preset_geometry(reader, &e, true, &mut prst, &mut adjustments)?
                }
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"cNvPr" => {}
                b"cNvSpPr" => {
                    is_text_box = attr_local(&e, b"txBox", reader)?.as_deref() == Some("1");
                }
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

    // Shapes with fill, outline, preset geometry, or custom geometry keep their
    // visual geometry even when they also contain placeholder or text-box text.
    if fill.is_some()
        || outline.is_some()
        || custom_geometry.is_some()
        || effects.is_some()
        || (prst.is_some() && placeholder.is_none() && !is_text_box)
    {
        let shape_type = prst
            .as_deref()
            .map(map_prst)
            .unwrap_or(ShapeType::Rectangle);
        let inner = if text.is_empty() && placeholder.is_none() && !is_text_box {
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
                placeholder,
                placeholder_idx,
                paragraphs,
                fill: None,
                outline: None,
                body,
                flip_h: tf.flip_h,
                flip_v: tf.flip_v,
            })
        };
        let shape_effects = effects.unwrap_or_default();
        return Ok(Some(SlideElement::Shape(Shape {
            shape_type,
            preset: prst,
            adjustments,
            x: tf.x,
            y: tf.y,
            width: tf.w,
            height: tf.h,
            rotation: tf.rot,
            fill,
            outline,
            text: inner,
            custom_geometry,
            flip_h: tf.flip_h,
            flip_v: tf.flip_v,
            effects: shape_effects,
        })));
    }

    // Placeholder shapes and explicit DrawingML text boxes without visual shape
    // properties can be represented as TextBox elements.
    if placeholder.is_some() || is_text_box {
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
            body,
            flip_h: tf.flip_h,
            flip_v: tf.flip_v,
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
            body,
            flip_h: tf.flip_h,
            flip_v: tf.flip_v,
        })));
    }

    Ok(None)
}

fn parse_graphic_frame(
    reader: &mut Reader<&[u8]>,
    rels: &Relationships,
    slide_path: &str,
    package: &OpcPackage,
    images: &mut Vec<SlidesImage>,
    theme: Option<&Theme>,
) -> PboResult<Option<SlideElement>> {
    let mut tf = Xf::default();
    let mut table = None;
    let mut fallback_element = None;
    let mut hidden = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"xfrm" => parse_xfrm_start(reader, &e, &mut tf)?,
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"tbl" => table = Some(parse_table(reader, theme)?),
                b"pic" => {
                    if fallback_element.is_none() {
                        fallback_element = parse_pic(reader, rels, slide_path, package, images)?;
                    } else {
                        let _ = parse_pic(reader, rels, slide_path, package, images)?;
                    }
                }
                b"AlternateContent" => {
                    let elements =
                        parse_alternate_content(reader, rels, slide_path, package, images, theme)?;
                    if fallback_element.is_none() {
                        fallback_element = elements.into_iter().next();
                    }
                }
                _ => {}
            },
            Event::Empty(e)
                if local_name(e.name().as_ref()) == b"cNvPr"
                    && attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") =>
            {
                hidden = true;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"xfrm" => {
                parse_xfrm_attrs(&e, reader, &mut tf)?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"graphicFrame" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    if hidden {
        return Ok(None);
    }
    if let Some(mut table) = table {
        table.x = tf.x;
        table.y = tf.y;
        table.width = tf.w;
        table.height = tf.h;
        return Ok(Some(SlideElement::Table(table)));
    }
    if let Some(mut element) = fallback_element {
        apply_frame_transform(&mut element, &tf);
        return Ok(Some(element));
    }
    Ok(None)
}

fn apply_frame_transform(element: &mut SlideElement, tf: &Xf) {
    if tf.w <= 0.0 || tf.h <= 0.0 {
        return;
    }
    match element {
        SlideElement::TextBox(text) => {
            text.x = tf.x;
            text.y = tf.y;
            text.width = tf.w;
            text.height = tf.h;
        }
        SlideElement::Shape(shape) => {
            shape.x = tf.x;
            shape.y = tf.y;
            shape.width = tf.w;
            shape.height = tf.h;
        }
        SlideElement::Picture(picture) => {
            picture.x = tf.x;
            picture.y = tf.y;
            picture.width = tf.w;
            picture.height = tf.h;
        }
        SlideElement::Connector(connector) => {
            connector.x = tf.x;
            connector.y = tf.y;
            connector.width = tf.w;
            connector.height = tf.h;
        }
        SlideElement::Group(group) => {
            group.x = tf.x;
            group.y = tf.y;
            group.width = tf.w;
            group.height = tf.h;
        }
        SlideElement::Table(table) => {
            table.x = tf.x;
            table.y = tf.y;
            table.width = tf.w;
            table.height = tf.h;
        }
    }
}

fn parse_table(reader: &mut Reader<&[u8]>, theme: Option<&Theme>) -> PboResult<SlideTable> {
    let mut table = SlideTable::default();
    let mut rows = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) if local_name(e.name().as_ref()) == b"gridCol" => {
                if let Some(w) = attr_i64(&e, b"w", reader)? {
                    table.columns.push(w as f64);
                }
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"tr" => {
                let height = attr_i64(&e, b"h", reader)?.unwrap_or_default() as f64;
                rows.push(parse_table_row(reader, height, theme)?);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"tbl" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    table.rows = rows;
    if table.width == 0.0 {
        table.width = table.columns.iter().sum();
    }
    if table.height == 0.0 {
        table.height = table.rows.iter().map(|row| row.height).sum();
    }
    Ok(table)
}

fn parse_table_row(
    reader: &mut Reader<&[u8]>,
    height: f64,
    theme: Option<&Theme>,
) -> PboResult<SlideTableRow> {
    let mut row = SlideTableRow {
        height,
        cells: Vec::new(),
    };
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"tc" => {
                row.cells.push(parse_table_cell(reader, theme)?);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"tr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(row)
}

fn parse_table_cell(
    reader: &mut Reader<&[u8]>,
    theme: Option<&Theme>,
) -> PboResult<SlideTableCell> {
    let mut cell = SlideTableCell {
        col_span: 1,
        row_span: 1,
        ..Default::default()
    };
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"txBody" => {
                    let mut bold = false;
                    let mut italic = false;
                    let mut color = None;
                    let mut size = None;
                    let mut align = None;
                    let mut body = TextBodyProperties::default();
                    let _ = parse_tx_body(
                        reader,
                        &mut cell.text,
                        &mut bold,
                        &mut italic,
                        &mut color,
                        &mut size,
                        &mut align,
                        &mut body,
                        theme,
                    )?;
                }
                b"tcPr" => parse_table_cell_props(reader, &mut cell, theme)?,
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"gridSpan" => {
                    if let Some(v) = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()) {
                        cell.col_span = v;
                    }
                }
                b"rowSpan" => {
                    if let Some(v) = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()) {
                        cell.row_span = v;
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"tc" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(cell)
}

fn parse_table_cell_props(
    reader: &mut Reader<&[u8]>,
    cell: &mut SlideTableCell,
    theme: Option<&Theme>,
) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"solidFill" => {
                cell.fill = parse_solid_fill(reader, theme)?;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"gradFill" => {
                cell.fill = parse_gradient_fill(reader, theme)?;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"pattFill" => {
                cell.fill = parse_pattern_fill(reader, &e, false, theme)?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"noFill" => {
                cell.fill = Some(Fill::None);
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"pattFill" => {
                cell.fill = parse_pattern_fill(reader, &e, true, theme)?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"tcPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_bg_pr(reader: &mut Reader<&[u8]>, theme: Option<&Theme>) -> PboResult<Option<Fill>> {
    let mut fill = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"solidFill" => {
                fill = parse_solid_fill(reader, theme)?;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"gradFill" => {
                fill = parse_gradient_fill(reader, theme)?;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"pattFill" => {
                fill = parse_pattern_fill(reader, &e, false, theme)?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"noFill" => {
                fill = Some(Fill::None);
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"pattFill" => {
                fill = parse_pattern_fill(reader, &e, true, theme)?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"bgPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(fill)
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
    body: &mut TextBodyProperties,
    theme: Option<&Theme>,
) -> PboResult<Vec<TextParagraph>> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();
    let mut first_para = true;
    let mut first_run = true;
    // Current paragraph state
    let mut para_align: Option<TextAlign> = None;
    let mut para_bullet: Option<String> = None;
    let mut para_margin_left: Option<f64> = None;
    let mut para_indent: Option<f64> = None;
    let mut para_runs: Vec<TextRun> = Vec::new();
    let mut default_run_style = RunStyle::default();
    let mut list_style_level: Option<u8> = None;
    // Current run state
    let mut run_bold = false;
    let mut run_italic = false;
    let mut run_size: Option<f32> = None;
    let mut run_color: Option<String> = None;
    let mut run_font_family: Option<String> = None;
    let mut run_underline = false;
    let mut run_strike = false;
    let mut run_baseline: Option<i32> = None;
    let mut tx_body_depth = 1usize;
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"txBody" => tx_body_depth += 1,
                b"bodyPr" => *body = parse_body_properties(reader, &e, false)?,
                name if paragraph_style_level(name).is_some() => {
                    list_style_level = paragraph_style_level(name);
                }
                b"p" => {
                    if !first_para {
                        text.push('\n');
                    }
                    first_para = false;
                    para_align = None;
                    para_bullet = None;
                    para_margin_left = None;
                    para_indent = None;
                    para_runs = Vec::new();
                }
                b"pPr" => {
                    para_margin_left = attr_i64(&e, b"marL", reader)?.map(|v| v as f64);
                    para_indent = attr_i64(&e, b"indent", reader)?.map(|v| v as f64);
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
                b"defRPr" => {
                    let parsed = parse_default_run_properties(reader, &e, false, theme)?;
                    if list_style_level.is_none_or(|level| level == 1) {
                        default_run_style.merge(parsed);
                    }
                }
                b"rPr" => {
                    let mut style = default_run_style.clone();
                    apply_run_style_attrs(&e, reader, &mut style)?;
                    run_bold = style.bold.unwrap_or(false);
                    run_italic = style.italic.unwrap_or(false);
                    run_underline = style.underline.unwrap_or(false);
                    run_strike = style.strike.unwrap_or(false);
                    run_baseline = style.baseline;
                    run_size = style.size;
                    run_color = style.color;
                    run_font_family = style.font_family;
                }
                b"latin" | b"ea" | b"cs" if run_font_family.is_none() => {
                    run_font_family = attr_local(&e, b"typeface", reader)?;
                }
                b"solidFill" => {
                    // Inside rPr → solidFill
                    if let Some(fill) = parse_solid_fill(reader, theme)?
                        && let Some(color) = solid_fill_color(&fill)
                    {
                        run_color = Some(color.to_owned());
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
                b"br" => {
                    text.push('\n');
                    para_runs.push(TextRun {
                        text: "\n".into(),
                        ..Default::default()
                    });
                    skip_to_end(reader, b"br")?;
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
                        font_family: run_font_family.clone(),
                        underline: run_underline,
                        strike: run_strike,
                        baseline: run_baseline,
                    });
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"bodyPr" => {
                    *body = TextBodyProperties {
                        inset_left: attr_i64(&e, b"lIns", reader)?.map(|v| v as f64),
                        inset_top: attr_i64(&e, b"tIns", reader)?.map(|v| v as f64),
                        inset_right: attr_i64(&e, b"rIns", reader)?.map(|v| v as f64),
                        inset_bottom: attr_i64(&e, b"bIns", reader)?.map(|v| v as f64),
                        anchor: attr_local(&e, b"anchor", reader)?,
                        rotation: attr_local(&e, b"rot", reader)?.and_then(|v| v.parse().ok()),
                        vertical: attr_local(&e, b"vert", reader)?,
                        wrap: attr_local(&e, b"wrap", reader)?,
                        vertical_overflow: attr_local(&e, b"vertOverflow", reader)?,
                        horizontal_overflow: attr_local(&e, b"horzOverflow", reader)?,
                    };
                }
                b"defRPr" => {
                    let parsed = parse_default_run_properties(reader, &e, true, theme)?;
                    if list_style_level.is_none_or(|level| level == 1) {
                        default_run_style.merge(parsed);
                    }
                }
                b"br" => {
                    text.push('\n');
                    para_runs.push(TextRun {
                        text: "\n".into(),
                        ..Default::default()
                    });
                }
                b"rPr" => {
                    let mut style = default_run_style.clone();
                    apply_run_style_attrs(&e, reader, &mut style)?;
                    run_bold = style.bold.unwrap_or(false);
                    run_italic = style.italic.unwrap_or(false);
                    run_underline = style.underline.unwrap_or(false);
                    run_strike = style.strike.unwrap_or(false);
                    run_baseline = style.baseline;
                    run_size = style.size;
                    run_color = style.color;
                    run_font_family = style.font_family;
                }
                b"latin" | b"ea" | b"cs" if run_font_family.is_none() => {
                    run_font_family = attr_local(&e, b"typeface", reader)?;
                }
                b"buNone" => para_bullet = None,
                b"buChar" => para_bullet = attr_local(&e, b"char", reader)?,
                b"buAutoNum" => para_bullet = Some("1.".into()),
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
                    para_margin_left = attr_i64(&e, b"marL", reader)?.map(|v| v as f64);
                    para_indent = attr_i64(&e, b"indent", reader)?.map(|v| v as f64);
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
            Event::End(e) if paragraph_style_level(local_name(e.name().as_ref())).is_some() => {
                list_style_level = None;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"p" && !para_runs.is_empty() => {
                paragraphs.push(TextParagraph {
                    runs: std::mem::take(&mut para_runs),
                    align: para_align,
                    bullet: para_bullet.clone(),
                    margin_left: para_margin_left,
                    indent: para_indent,
                });
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"txBody" => {
                tx_body_depth = tx_body_depth.saturating_sub(1);
                if tx_body_depth == 0 {
                    // Flush any remaining paragraph
                    if !para_runs.is_empty() {
                        paragraphs.push(TextParagraph {
                            runs: std::mem::take(&mut para_runs),
                            align: para_align,
                            bullet: para_bullet.clone(),
                            margin_left: para_margin_left,
                            indent: para_indent,
                        });
                    }
                    break;
                }
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
    let mut prst = None;
    let mut fill = None;
    let mut outline = None;
    let mut custom_geometry = None;
    let mut effects: Option<VisualEffects> = None;
    let mut crop = PictureCrop::default();
    let mut hidden = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"spPr" => parse_shape_props(
                    reader,
                    &mut tf,
                    &mut prst,
                    &mut Vec::new(),
                    &mut fill,
                    &mut outline,
                    &mut custom_geometry,
                    &mut effects,
                    None,
                )?,
                b"srcRect" => parse_crop_rect_empty(&e, reader, &mut crop)?,
                b"tile" => crop.tile = true,
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
                b"xfrm" => parse_xfrm_start(reader, &e, &mut tf)?,
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
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
                b"xfrm" => parse_xfrm_attrs(&e, reader, &mut tf)?,
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
        custom_geometry,
        crop,
        flip_h: tf.flip_h,
        flip_v: tf.flip_v,
        effects: effects.unwrap_or_default(),
    })))
}

fn ensure_image(package: &OpcPackage, path: &str, images: &mut Vec<SlidesImage>) -> String {
    if images.iter().any(|i| i.id == path) {
        return path.to_string();
    }
    let bytes = package.parts.get(path).cloned().unwrap_or_default();
    let ext = path.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
    let (mime, data_base64) = match ext.as_str() {
        "emf" => match emf_to_svg(&bytes) {
            Some(svg) => ("image/svg+xml".to_string(), b64::encode(svg.as_bytes())),
            None => ("image/x-emf".to_string(), b64::encode(&bytes)),
        },
        "png" => ("image/png".to_string(), b64::encode(&bytes)),
        "jpg" | "jpeg" => ("image/jpeg".to_string(), b64::encode(&bytes)),
        "gif" => ("image/gif".to_string(), b64::encode(&bytes)),
        "bmp" => ("image/bmp".to_string(), b64::encode(&bytes)),
        "webp" => ("image/webp".to_string(), b64::encode(&bytes)),
        "svg" => ("image/svg+xml".to_string(), b64::encode(&bytes)),
        "wmf" => ("image/x-wmf".to_string(), b64::encode(&bytes)),
        "wdp" => ("image/vnd.ms-photo".to_string(), b64::encode(&bytes)),
        _ => ("application/octet-stream".to_string(), b64::encode(&bytes)),
    };
    images.push(SlidesImage {
        id: path.to_string(),
        mime,
        data_base64,
    });
    path.to_string()
}

fn le_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    bytes
        .get(offset..offset + 2)
        .and_then(|value| value.try_into().ok())
        .map(u16::from_le_bytes)
}

fn le_i16(bytes: &[u8], offset: usize) -> Option<i16> {
    le_u16(bytes, offset).map(|value| value as i16)
}

fn le_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    bytes
        .get(offset..offset + 4)
        .and_then(|value| value.try_into().ok())
        .map(u32::from_le_bytes)
}

fn le_i32(bytes: &[u8], offset: usize) -> Option<i32> {
    le_u32(bytes, offset).map(|value| value as i32)
}

#[derive(Debug, Clone)]
struct EmfState {
    window_org: (f64, f64),
    window_ext: (f64, f64),
    viewport_org: (f64, f64),
    viewport_ext: (f64, f64),
    current_brush: String,
    current_pen: String,
    brushes: HashMap<u32, String>,
    pens: HashMap<u32, String>,
    in_path: bool,
    path: String,
}

impl Default for EmfState {
    fn default() -> Self {
        Self {
            window_org: (0.0, 0.0),
            window_ext: (1.0, 1.0),
            viewport_org: (0.0, 0.0),
            viewport_ext: (1.0, 1.0),
            current_brush: "#000000".into(),
            current_pen: "#000000".into(),
            brushes: HashMap::new(),
            pens: HashMap::new(),
            in_path: false,
            path: String::new(),
        }
    }
}

impl EmfState {
    fn map_point(&self, x: i32, y: i32) -> (f64, f64) {
        let sx = if self.window_ext.0.abs() > f64::EPSILON {
            self.viewport_ext.0 / self.window_ext.0
        } else {
            1.0
        };
        let sy = if self.window_ext.1.abs() > f64::EPSILON {
            self.viewport_ext.1 / self.window_ext.1
        } else {
            1.0
        };
        (
            self.viewport_org.0 + (x as f64 - self.window_org.0) * sx,
            self.viewport_org.1 + (y as f64 - self.window_org.1) * sy,
        )
    }

    fn push_move(&mut self, x: f64, y: f64) {
        self.path.push_str(&format!("M {x:.3} {y:.3} "));
    }

    fn push_line(&mut self, x: f64, y: f64) {
        self.path.push_str(&format!("L {x:.3} {y:.3} "));
    }
}

fn emf_colorref_to_css(color: u32) -> String {
    let r = color & 0xFF;
    let g = (color >> 8) & 0xFF;
    let b = (color >> 16) & 0xFF;
    format!("#{r:02X}{g:02X}{b:02X}")
}

fn svg_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn emf_polyline16_points(bytes: &[u8], offset: usize) -> Option<Vec<(f64, f64)>> {
    let count = le_u32(bytes, offset + 24)? as usize;
    let mut points = Vec::with_capacity(count);
    let mut pos = offset + 28;
    for _ in 0..count {
        let x = le_i16(bytes, pos)? as i32;
        let y = le_i16(bytes, pos + 2)? as i32;
        points.push((x as f64, y as f64));
        pos += 4;
    }
    Some(points)
}

fn emf_polypolyline16_points(bytes: &[u8], offset: usize) -> Option<Vec<Vec<(i32, i32)>>> {
    let poly_count = le_u32(bytes, offset + 24)? as usize;
    let mut pos = offset + 32;
    let mut counts = Vec::with_capacity(poly_count);
    for _ in 0..poly_count {
        let count = le_u32(bytes, pos)? as usize;
        counts.push(count);
        pos += 4;
    }
    let mut polys = Vec::with_capacity(poly_count);
    for count in counts {
        let mut points = Vec::with_capacity(count);
        for _ in 0..count {
            let x = le_i16(bytes, pos)? as i32;
            let y = le_i16(bytes, pos + 2)? as i32;
            points.push((x, y));
            pos += 4;
        }
        polys.push(points);
    }
    Some(polys)
}

fn svg_points(points: &[(f64, f64)]) -> String {
    points
        .iter()
        .map(|(x, y)| format!("{x:.3},{y:.3}"))
        .collect::<Vec<_>>()
        .join(" ")
}

fn emf_to_svg(bytes: &[u8]) -> Option<String> {
    if le_u32(bytes, 0)? != 1 || le_u32(bytes, 4)? < 88 {
        return None;
    }
    let left = le_i32(bytes, 8)? as f64;
    let top = le_i32(bytes, 12)? as f64;
    let right = le_i32(bytes, 16)? as f64;
    let bottom = le_i32(bytes, 20)? as f64;
    let width = (right - left).abs().max(1.0);
    let height = (bottom - top).abs().max(1.0);
    let mut state = EmfState::default();
    let mut body = String::new();
    let mut offset = 0usize;

    while offset + 8 <= bytes.len() {
        let record_type = le_u32(bytes, offset)?;
        let record_size = le_u32(bytes, offset + 4)? as usize;
        if record_size < 8 || offset + record_size > bytes.len() {
            break;
        }

        match record_type {
            9 => {
                state.window_ext = (
                    le_i32(bytes, offset + 8)? as f64,
                    le_i32(bytes, offset + 12)? as f64,
                );
            }
            10 => {
                state.window_org = (
                    le_i32(bytes, offset + 8)? as f64,
                    le_i32(bytes, offset + 12)? as f64,
                );
            }
            11 => {
                state.viewport_ext = (
                    le_i32(bytes, offset + 8)? as f64,
                    le_i32(bytes, offset + 12)? as f64,
                );
            }
            12 => {
                state.viewport_org = (
                    le_i32(bytes, offset + 8)? as f64,
                    le_i32(bytes, offset + 12)? as f64,
                );
            }
            27 => {
                let (x, y) =
                    state.map_point(le_i32(bytes, offset + 8)?, le_i32(bytes, offset + 12)?);
                if state.in_path {
                    state.push_move(x, y);
                }
            }
            37 => {
                let handle = le_u32(bytes, offset + 8)?;
                if let Some(brush) = state.brushes.get(&handle) {
                    state.current_brush.clone_from(brush);
                }
                if let Some(pen) = state.pens.get(&handle) {
                    state.current_pen.clone_from(pen);
                }
            }
            39 => {
                let handle = le_u32(bytes, offset + 8)?;
                let color = le_u32(bytes, offset + 16)?;
                state.brushes.insert(handle, emf_colorref_to_css(color));
            }
            40 => {
                let handle = le_u32(bytes, offset + 8)?;
                state.brushes.remove(&handle);
                state.pens.remove(&handle);
            }
            59 => {
                state.in_path = true;
                state.path.clear();
            }
            60 => {
                state.in_path = false;
            }
            61 if state.in_path => state.path.push_str("Z "),
            61 => {}
            62 => {
                if !state.path.trim().is_empty() {
                    body.push_str(&format!(
                        r#"<path d="{}" fill="{}" stroke="none"/>"#,
                        svg_escape(state.path.trim()),
                        state.current_brush
                    ));
                }
                state.path.clear();
            }
            64 => {
                if !state.path.trim().is_empty() {
                    body.push_str(&format!(
                        r#"<path d="{}" fill="none" stroke="{}" stroke-width="1"/>"#,
                        svg_escape(state.path.trim()),
                        state.current_pen
                    ));
                }
                state.path.clear();
            }
            88 => {
                if let Some(raw_points) = emf_polyline16_points(bytes, offset) {
                    let points = raw_points
                        .iter()
                        .map(|(x, y)| state.map_point(*x as i32, *y as i32))
                        .collect::<Vec<_>>();
                    if state.in_path {
                        if let Some((x, y)) = points.first() {
                            state.push_move(*x, *y);
                        }
                        for (x, y) in points.iter().skip(1) {
                            state.push_line(*x, *y);
                        }
                    } else if !points.is_empty() {
                        body.push_str(&format!(
                            r#"<polyline points="{}" fill="none" stroke="{}" stroke-width="1"/>"#,
                            svg_points(&points),
                            state.current_pen
                        ));
                    }
                }
            }
            91 => {
                if let Some(polys) = emf_polypolyline16_points(bytes, offset) {
                    for poly in polys {
                        let points = poly
                            .iter()
                            .map(|(x, y)| state.map_point(*x, *y))
                            .collect::<Vec<_>>();
                        if points.len() < 2 {
                            continue;
                        }
                        let closed = poly.first() == poly.last();
                        if closed {
                            body.push_str(&format!(
                                r#"<polygon points="{}" fill="{}" stroke="{}" stroke-width="0.5"/>"#,
                                svg_points(&points),
                                state.current_brush,
                                state.current_brush
                            ));
                        } else {
                            body.push_str(&format!(
                                r#"<polyline points="{}" fill="none" stroke="{}" stroke-width="1"/>"#,
                                svg_points(&points),
                                state.current_pen
                            ));
                        }
                    }
                }
            }
            95 => {
                let handle = le_u32(bytes, offset + 8)?;
                let color = le_u32(bytes, offset + 40).unwrap_or_default();
                state.pens.insert(handle, emf_colorref_to_css(color));
            }
            14 => break,
            _ => {}
        }

        offset += record_size;
    }

    if body.is_empty() {
        return None;
    }
    Some(format!(
        r#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="{left:.3} {top:.3} {width:.3} {height:.3}">{body}</svg>"#
    ))
}

fn parse_cxn(reader: &mut Reader<&[u8]>, theme: Option<&Theme>) -> PboResult<Option<SlideElement>> {
    let mut tf = Xf::default();
    let mut prst = String::new();
    let mut adjustments = Vec::new();
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
                    let mut custom_geometry = None;
                    let mut effects = None;
                    parse_shape_props(
                        reader,
                        &mut tf,
                        &mut maybe_prst,
                        &mut adjustments,
                        &mut fill,
                        &mut outline,
                        &mut custom_geometry,
                        &mut effects,
                        theme,
                    )?;
                    if let Some(parsed_prst) = maybe_prst {
                        prst = parsed_prst;
                    }
                }
                b"prstGeom" => {
                    let mut maybe_prst = Some(prst.clone()).filter(|value| !value.is_empty());
                    parse_preset_geometry(reader, &e, false, &mut maybe_prst, &mut adjustments)?;
                    if let Some(parsed_prst) = maybe_prst {
                        prst = parsed_prst;
                    }
                }
                b"xfrm" => parse_xfrm_start(reader, &e, &mut tf)?,
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"spPr" => {} // empty <p:spPr/> - nothing to parse
                b"cNvPr" if attr_local(&e, b"hidden", reader)?.as_deref() == Some("1") => {
                    hidden = true;
                }
                b"cNvPr" => {}
                b"prstGeom" => {
                    let mut maybe_prst = Some(prst.clone()).filter(|value| !value.is_empty());
                    parse_preset_geometry(reader, &e, true, &mut maybe_prst, &mut adjustments)?;
                    if let Some(parsed_prst) = maybe_prst {
                        prst = parsed_prst;
                    }
                }
                b"xfrm" => parse_xfrm_attrs(&e, reader, &mut tf)?,
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
        preset_geometry: if prst.is_empty() { None } else { Some(prst) },
        adjustments,
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
                b"graphicFrame" => {
                    if let Some(el) =
                        parse_graphic_frame(reader, rels, slide_path, package, images, theme)?
                    {
                        elements.push(el);
                    }
                }
                b"AlternateContent" => {
                    elements.extend(parse_alternate_content(
                        reader, rels, slide_path, package, images, theme,
                    )?);
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
                parse_xfrm_attrs(&e, reader, &mut gxf.off)?;
            }
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"xfrm" => parse_xfrm_attrs(&e, reader, &mut gxf.off)?,
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
