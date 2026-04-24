//! Read a `.xlsx` file into a `CalcDocument`.

use pbo_common::model::calc::{
    BorderEdge, CalcDocument, Cell, CellBorder, CellFill, CellFont, CellValue, CellXf, ColumnDim,
    DefinedName, MergedRange, NumFmt, Sheet, SheetHyperlink, SheetView, SpreadsheetMetadata,
    WorkbookStyles,
};
use pbo_common::{PboError, PboResult};
use pbo_ooxml::xmlio::{attr_local, attr_qualified, attr_u32, local_name};
use pbo_ooxml::{OpcPackage, Relationships};
use quick_xml::Reader;
use quick_xml::events::Event;

const WORKBOOK_PART: &str = "xl/workbook.xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const SHARED_STRINGS_PART: &str = "xl/sharedStrings.xml";
const STYLES_PART: &str = "xl/styles.xml";
const CORE_PART: &str = "docProps/core.xml";
const REL_SHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

pub fn read_xlsx(data: &[u8]) -> PboResult<CalcDocument> {
    let package = OpcPackage::from_bytes(data)?;
    let mut doc = CalcDocument::default();

    if let Ok(xml) = package.get_part_str(CORE_PART) {
        doc.metadata = parse_core_properties(xml)?;
    }
    if let Ok(xml) = package.get_part_str(STYLES_PART) {
        doc.styles = parse_styles(xml)?;
    }
    let shared = package
        .get_part_str(SHARED_STRINGS_PART)
        .ok()
        .map(parse_shared_strings)
        .transpose()?
        .unwrap_or_default();

    let wb_xml = package.get_part_str(WORKBOOK_PART)?;
    let (sheet_refs, defined_names) = parse_workbook(wb_xml)?;
    doc.defined_names = defined_names;

    let rels = package
        .get_part_str(WORKBOOK_RELS)
        .ok()
        .map(Relationships::parse)
        .transpose()?
        .unwrap_or_default();

    for sr in sheet_refs {
        let Some(rel) = rels.find_by_id(&sr.rel_id) else {
            continue;
        };
        if rel.rel_type != REL_SHEET {
            continue;
        }
        let target = normalize_target(&rel.target);
        let Ok(sheet_xml) = package.get_part_str(&target) else {
            continue;
        };
        let sheet_rels = package
            .get_part_str(&sheet_rels_path_for(&target))
            .ok()
            .map(Relationships::parse)
            .transpose()?
            .unwrap_or_default();
        let sheet = parse_sheet(sheet_xml, &sr.name, &shared, &sheet_rels, &doc.styles)?;
        doc.sheets.push(sheet);
    }
    Ok(doc)
}

fn normalize_target(target: &str) -> String {
    if let Some(rest) = target.strip_prefix('/') {
        rest.into()
    } else {
        format!("xl/{target}")
    }
}

fn sheet_rels_path_for(sheet_path: &str) -> String {
    let (dir, file) = sheet_path.rsplit_once('/').unwrap_or(("", sheet_path));
    if dir.is_empty() {
        format!("_rels/{file}.rels")
    } else {
        format!("{dir}/_rels/{file}.rels")
    }
}

fn xml_err(e: quick_xml::Error) -> PboError {
    PboError::Xml(e.to_string())
}

// ── workbook.xml ────────────────────────────────────
struct SheetRef {
    name: String,
    rel_id: String,
}

fn parse_workbook(xml: &str) -> PboResult<(Vec<SheetRef>, Vec<DefinedName>)> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut sheets = Vec::new();
    let mut defined_names = Vec::new();
    let mut pending: Option<(String, Option<u32>)> = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) => match local_name(e.name().as_ref()) {
                b"sheet" => {
                    let name = attr_local(&e, b"name", &reader)?.unwrap_or_default();
                    let rel_id = attr_qualified(&e, b"r:id", &reader)?.unwrap_or_default();
                    sheets.push(SheetRef { name, rel_id });
                }
                b"definedName" => {
                    let name = attr_local(&e, b"name", &reader)?.unwrap_or_default();
                    let local = attr_u32(&e, b"localSheetId", &reader)?;
                    pending = Some((name, local));
                }
                _ => {}
            },
            Event::Text(t) => {
                if let Some((name, local)) = pending.take() {
                    let formula = t
                        .unescape()
                        .map_err(|e| PboError::Xml(e.to_string()))?
                        .into_owned();
                    defined_names.push(DefinedName {
                        name,
                        formula,
                        local_sheet_id: local,
                    });
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"definedName" => pending = None,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok((sheets, defined_names))
}

// ── sharedStrings.xml ───────────────────────────────
fn parse_shared_strings(xml: &str) -> PboResult<Vec<String>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"si" => {
                out.push(parse_si(&mut reader)?);
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn parse_si(reader: &mut Reader<&[u8]>) -> PboResult<String> {
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"t" => read_text(reader, &mut text, b"t")?,
                b"r" => parse_rich_run(reader, &mut text)?,
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"si" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(text)
}

fn parse_rich_run(reader: &mut Reader<&[u8]>, out: &mut String) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"t" => {
                read_text(reader, out, b"t")?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"r" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn read_text(reader: &mut Reader<&[u8]>, out: &mut String, closing: &[u8]) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Text(t) => {
                out.push_str(&t.unescape().map_err(|e| PboError::Xml(e.to_string()))?);
            }
            Event::End(e) if local_name(e.name().as_ref()) == closing => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

// ── sheet XML ───────────────────────────────────────
fn parse_sheet(
    xml: &str,
    name: &str,
    shared: &[String],
    rels: &Relationships,
    styles: &WorkbookStyles,
) -> PboResult<Sheet> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut sheet = Sheet {
        name: name.to_string(),
        ..Default::default()
    };
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"sheetData" => parse_sheet_data(&mut reader, &mut sheet, shared, styles)?,
                b"mergeCells" => parse_merges(&mut reader, &mut sheet)?,
                b"cols" => parse_cols(&mut reader, &mut sheet)?,
                b"sheetViews" => parse_sheet_views(&mut reader, &mut sheet)?,
                b"hyperlinks" => parse_hyperlinks(&mut reader, &mut sheet, rels)?,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(sheet)
}

fn parse_cols(reader: &mut Reader<&[u8]>, sheet: &mut Sheet) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) if local_name(e.name().as_ref()) == b"col" => {
                let min = attr_u32(&e, b"min", reader)?.unwrap_or(1);
                let max = attr_u32(&e, b"max", reader)?.unwrap_or(min);
                let width = attr_local(&e, b"width", reader)?
                    .and_then(|v| v.parse().ok())
                    .unwrap_or(0.0);
                let custom = attr_local(&e, b"customWidth", reader)?.as_deref() == Some("1");
                let hidden = attr_local(&e, b"hidden", reader)?.as_deref() == Some("1");
                sheet.columns.push(ColumnDim {
                    min,
                    max,
                    width,
                    custom_width: custom,
                    hidden,
                });
                for c in min..=max {
                    sheet.column_widths.insert(c, width);
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"cols" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_merges(reader: &mut Reader<&[u8]>, sheet: &mut Sheet) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e)
                if local_name(e.name().as_ref()) == b"mergeCell" =>
            {
                if let Some(r) = attr_local(&e, b"ref", reader)?
                    && let Some(m) = parse_merge_ref(&r)
                {
                    sheet.merges.push(m);
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"mergeCells" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_merge_ref(r: &str) -> Option<MergedRange> {
    let (a, b) = r.split_once(':')?;
    let (c0, r0) = parse_ref(a)?;
    let (c1, r1) = parse_ref(b)?;
    Some(MergedRange {
        row_start: r0,
        col_start: c0,
        row_end: r1,
        col_end: c1,
    })
}

fn parse_sheet_views(reader: &mut Reader<&[u8]>, sheet: &mut Sheet) -> PboResult<()> {
    let mut view = SheetView {
        show_grid_lines: true,
        zoom_scale: 100,
        ..Default::default()
    };
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) => match local_name(e.name().as_ref()) {
                b"sheetView" => {
                    if let Some(v) = attr_local(&e, b"showGridLines", reader)? {
                        view.show_grid_lines = v != "0";
                    }
                    if let Some(v) = attr_u32(&e, b"zoomScale", reader)? {
                        view.zoom_scale = v;
                    }
                }
                b"pane" => {
                    view.freeze_col = attr_u32(&e, b"xSplit", reader)?.unwrap_or(0);
                    view.freeze_row = attr_u32(&e, b"ySplit", reader)?.unwrap_or(0);
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetViews" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    sheet.view = Some(view);
    Ok(())
}

fn parse_hyperlinks(
    reader: &mut Reader<&[u8]>,
    sheet: &mut Sheet,
    rels: &Relationships,
) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e)
                if local_name(e.name().as_ref()) == b"hyperlink" =>
            {
                let cell_ref = attr_local(&e, b"ref", reader)?.unwrap_or_default();
                let target = match attr_qualified(&e, b"r:id", reader)? {
                    Some(rid) => rels
                        .find_by_id(&rid)
                        .map(|r| r.target.clone())
                        .unwrap_or_default(),
                    None => attr_local(&e, b"location", reader)?.unwrap_or_default(),
                };
                let display = attr_local(&e, b"display", reader)?;
                if !cell_ref.is_empty() && !target.is_empty() {
                    sheet.hyperlinks.push(SheetHyperlink {
                        cell_ref,
                        target,
                        display,
                    });
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"hyperlinks" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_sheet_data(
    reader: &mut Reader<&[u8]>,
    sheet: &mut Sheet,
    shared: &[String],
    styles: &WorkbookStyles,
) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"row" => {
                let ht = attr_local(&e, b"ht", reader)?.and_then(|v| v.parse().ok());
                let row_idx = attr_u32(&e, b"r", reader)?.unwrap_or(0);
                if let (Some(h), true) = (ht, row_idx > 0) {
                    sheet.row_heights.insert(row_idx, h);
                }
                parse_row(reader, sheet, row_idx, shared, styles)?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetData" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_row(
    reader: &mut Reader<&[u8]>,
    sheet: &mut Sheet,
    default_row: u32,
    shared: &[String],
    styles: &WorkbookStyles,
) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"c" => {
                let cell_ref = attr_local(&e, b"r", reader)?;
                let t_attr = attr_local(&e, b"t", reader)?;
                let s_attr = attr_u32(&e, b"s", reader)?;
                let (col, row) = match &cell_ref {
                    Some(s) => parse_ref(s).unwrap_or((0, default_row)),
                    None => (0, default_row),
                };
                let (value, formula) =
                    parse_cell_body(reader, t_attr.as_deref(), shared, styles, s_attr)?;
                if col > 0 && row > 0 {
                    sheet.cells.entry(row).or_default().insert(
                        col,
                        Cell {
                            value,
                            formula,
                            style_id: s_attr,
                        },
                    );
                }
            }
            Event::Empty(_) => {}
            Event::End(e) if local_name(e.name().as_ref()) == b"row" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_cell_body(
    reader: &mut Reader<&[u8]>,
    t_attr: Option<&str>,
    shared: &[String],
    styles: &WorkbookStyles,
    s_attr: Option<u32>,
) -> PboResult<(CellValue, Option<String>)> {
    let mut value: Option<String> = None;
    let mut formula: Option<String> = None;
    let mut inline_text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"v" => {
                    let mut t = String::new();
                    read_text(reader, &mut t, b"v")?;
                    value = Some(t);
                }
                b"f" => {
                    let mut t = String::new();
                    read_text(reader, &mut t, b"f")?;
                    formula = Some(t);
                }
                b"is" => parse_inline_is(reader, &mut inline_text)?,
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"c" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok((
        build_value(t_attr, value.as_deref(), &inline_text, shared, styles, s_attr),
        formula,
    ))
}

fn parse_inline_is(reader: &mut Reader<&[u8]>, out: &mut String) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"t" => read_text(reader, out, b"t")?,
                b"r" => parse_rich_run(reader, out)?,
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"is" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn build_value(
    t: Option<&str>,
    v: Option<&str>,
    inline_text: &str,
    shared: &[String],
    styles: &WorkbookStyles,
    s_attr: Option<u32>,
) -> CellValue {
    match t {
        Some("s") => match v
            .and_then(|s| s.parse::<usize>().ok())
            .and_then(|i| shared.get(i))
        {
            Some(s) => CellValue::Text(s.clone()),
            None => CellValue::Empty,
        },
        Some("str") => CellValue::Text(v.unwrap_or("").into()),
        Some("inlineStr") => CellValue::Text(inline_text.to_string()),
        Some("b") => CellValue::Boolean(matches!(v, Some("1") | Some("true"))),
        Some("e") => CellValue::Error(v.unwrap_or("").into()),
        _ => match v.and_then(|s| s.parse::<f64>().ok()) {
            Some(n) => {
                if let Some(idx) = s_attr
                    && let Some(fmt) = styles.number_format_for(idx)
                    && looks_like_date_format(fmt)
                {
                    return CellValue::Date(excel_serial_to_iso(n));
                }
                CellValue::Number(n)
            }
            None => CellValue::Empty,
        },
    }
}

fn looks_like_date_format(fmt: &str) -> bool {
    let l = fmt.to_ascii_lowercase();
    (l.contains('d') && l.contains('m')) || l.contains("yy") || l.contains("hh:mm")
}

fn excel_serial_to_iso(serial: f64) -> String {
    let days = serial.floor() as i64;
    let (y, m, d) = gregorian_from_days(days + 693594);
    format!("{y:04}-{m:02}-{d:02}")
}

fn gregorian_from_days(days: i64) -> (i32, u32, u32) {
    let a = days + 32045;
    let b = (4 * a + 3) / 146097;
    let c = a - (146097 * b) / 4;
    let d = (4 * c + 3) / 1461;
    let e = c - (1461 * d) / 4;
    let m = (5 * e + 2) / 153;
    let day = (e - (153 * m + 2) / 5 + 1) as u32;
    let month = (m + 3 - 12 * (m / 10)) as u32;
    let year = (100 * b + d - 4800 + (m / 10)) as i32;
    (year, month, day)
}

fn parse_ref(s: &str) -> Option<(u32, u32)> {
    let mut col = 0u32;
    let mut row_str = String::new();
    for c in s.chars() {
        if c == '$' {
            continue;
        }
        if c.is_ascii_alphabetic() {
            col = col * 26 + (c.to_ascii_uppercase() as u32 - b'A' as u32 + 1);
        } else {
            row_str.push(c);
        }
    }
    let row: u32 = row_str.parse().ok()?;
    Some((col, row))
}

// ── styles.xml ──────────────────────────────────────
fn parse_styles(xml: &str) -> PboResult<WorkbookStyles> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut out = WorkbookStyles::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"numFmts" => parse_num_fmts(&mut reader, &mut out.num_fmts)?,
                b"fonts" => parse_fonts(&mut reader, &mut out.fonts)?,
                b"fills" => parse_fills(&mut reader, &mut out.fills)?,
                b"borders" => parse_borders(&mut reader, &mut out.borders)?,
                b"cellXfs" => parse_cell_xfs(&mut reader, &mut out.cell_xfs)?,
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn parse_num_fmts(reader: &mut Reader<&[u8]>, out: &mut Vec<NumFmt>) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) if local_name(e.name().as_ref()) == b"numFmt" => {
                let id = attr_u32(&e, b"numFmtId", reader)?.unwrap_or(0);
                let code = attr_local(&e, b"formatCode", reader)?.unwrap_or_default();
                out.push(NumFmt { id, format_code: code });
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"numFmts" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_fonts(reader: &mut Reader<&[u8]>, out: &mut Vec<CellFont>) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"font" => {
                out.push(parse_font(reader)?);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"fonts" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_font(reader: &mut Reader<&[u8]>) -> PboResult<CellFont> {
    let mut f = CellFont::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) => match local_name(e.name().as_ref()) {
                b"b" => f.bold = true,
                b"i" => f.italic = true,
                b"u" => f.underline = true,
                b"strike" => f.strike = true,
                b"sz" => f.size = attr_local(&e, b"val", reader)?.and_then(|v| v.parse().ok()),
                b"name" | b"rFont" => f.name = attr_local(&e, b"val", reader)?,
                b"color" => {
                    f.color = attr_local(&e, b"rgb", reader)?
                        .or(attr_local(&e, b"val", reader)?);
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"font" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(f)
}

fn parse_fills(reader: &mut Reader<&[u8]>, out: &mut Vec<CellFill>) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"fill" => {
                out.push(parse_fill(reader)?);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"fills" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_fill(reader: &mut Reader<&[u8]>) -> PboResult<CellFill> {
    let mut fill = CellFill::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) => match local_name(e.name().as_ref()) {
                b"patternFill" => {
                    fill.pattern_type =
                        attr_local(&e, b"patternType", reader)?.unwrap_or_default();
                }
                b"fgColor" => {
                    fill.fg_color = attr_local(&e, b"rgb", reader)?
                        .or(attr_local(&e, b"val", reader)?);
                }
                b"bgColor" => {
                    fill.bg_color = attr_local(&e, b"rgb", reader)?
                        .or(attr_local(&e, b"val", reader)?);
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"fill" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(fill)
}

fn parse_borders(reader: &mut Reader<&[u8]>, out: &mut Vec<CellBorder>) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"border" => {
                out.push(parse_border(reader)?);
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"border" => {
                out.push(CellBorder::default());
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"borders" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_border(reader: &mut Reader<&[u8]>) -> PboResult<CellBorder> {
    let mut b = CellBorder::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                let side = matches!(name.as_slice(), b"top" | b"bottom" | b"left" | b"right");
                if side {
                    let style = attr_local(&e, b"style", reader)?;
                    let edge = style.map(|s| BorderEdge { style: s, color: None });
                    match name.as_slice() {
                        b"top" => b.top = edge,
                        b"bottom" => b.bottom = edge,
                        b"left" => b.left = edge,
                        b"right" => b.right = edge,
                        _ => {}
                    }
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"border" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(b)
}

fn parse_cell_xfs(reader: &mut Reader<&[u8]>, out: &mut Vec<CellXf>) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"xf" => {
                let xf = CellXf {
                    num_fmt_id: attr_u32(&e, b"numFmtId", reader)?,
                    font_id: attr_u32(&e, b"fontId", reader)?,
                    fill_id: attr_u32(&e, b"fillId", reader)?,
                    border_id: attr_u32(&e, b"borderId", reader)?,
                    alignment: None,
                    apply_number_format: attr_local(&e, b"applyNumberFormat", reader)?.as_deref()
                        == Some("1"),
                    apply_font: attr_local(&e, b"applyFont", reader)?.as_deref() == Some("1"),
                    apply_fill: attr_local(&e, b"applyFill", reader)?.as_deref() == Some("1"),
                    apply_border: attr_local(&e, b"applyBorder", reader)?.as_deref() == Some("1"),
                    apply_alignment: attr_local(&e, b"applyAlignment", reader)?.as_deref()
                        == Some("1"),
                };
                out.push(xf);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"cellXfs" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

// ── core.xml ────────────────────────────────────────
fn parse_core_properties(xml: &str) -> PboResult<SpreadsheetMetadata> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut md = SpreadsheetMetadata::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => {
                let qname = e.name().as_ref().to_vec();
                let mut text = String::new();
                read_text(&mut reader, &mut text, local_name(&qname))?;
                match local_name(&qname) {
                    b"title" => md.title = non_empty(text),
                    b"creator" => md.author = non_empty(text),
                    b"lastModifiedBy" => md.last_modified_by = non_empty(text),
                    b"created" => md.created = non_empty(text),
                    b"modified" => md.modified = non_empty(text),
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

fn non_empty(s: String) -> Option<String> {
    if s.is_empty() { None } else { Some(s) }
}