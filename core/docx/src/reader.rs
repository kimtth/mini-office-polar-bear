// DOCX reader: parses a .docx OPC package into a WriterDocument.
// Covers paragraphs, runs, rPr, pPr (jc, pStyle headings), tables
// (cells, grid spans, borders), hyperlinks, and core metadata.

use pbo_common::b64;
use pbo_common::model::writer::{
    Alignment, Block, BorderLine, CellBorders, DocumentMetadata, EmbeddedImage, Hyperlink,
    HyperlinkTarget, Paragraph, ParagraphStyle, Run, RunElement, RunStyle, Table, TableCell,
    TableRow, VMerge, WriterDocument,
};
use pbo_common::{PboError, PboResult};
use pbo_ooxml::xmlio::{attr_local, attr_qualified, attr_u32, local_name, on_off};
use pbo_ooxml::{OpcPackage, Relationships};
use quick_xml::Reader;
use quick_xml::events::Event;

const DOCUMENT_PART: &str = "word/document.xml";
const DOCUMENT_RELS: &str = "word/_rels/document.xml.rels";
const CORE_PART: &str = "docProps/core.xml";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

pub fn read_docx(data: &[u8]) -> PboResult<WriterDocument> {
    let package = OpcPackage::from_bytes(data)?;
    let mut doc = WriterDocument::default();

    if let Ok(xml) = package.get_part_str(CORE_PART) {
        doc.metadata = parse_core(xml)?;
    }

    let rels = package
        .get_part_str(DOCUMENT_RELS)
        .ok()
        .map(Relationships::parse)
        .transpose()?
        .unwrap_or_default();

    for r in &rels.0 {
        if r.rel_type == REL_HYPERLINK {
            doc.resources.hyperlinks.push(HyperlinkTarget {
                id: r.id.clone(),
                url: r.target.clone(),
            });
        } else if r.rel_type == REL_IMAGE {
            let path = if let Some(rest) = r.target.strip_prefix('/') {
                rest.to_string()
            } else {
                format!("word/{}", r.target)
            };
            let bytes = package.parts.get(&path).cloned().unwrap_or_default();
            let mime = mime_for(&r.target);
            doc.resources.images.push(EmbeddedImage {
                id: r.id.clone(),
                mime,
                data_base64: b64::encode(&bytes),
            });
        }
    }

    let body_xml = package.get_part_str(DOCUMENT_PART)?;
    doc.body = parse_body(body_xml)?;
    Ok(doc)
}

fn mime_for(target: &str) -> String {
    let ext = target.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
    match ext.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "svg" => "image/svg+xml",
        _ => "application/octet-stream",
    }
    .to_string()
}

fn xml_err(e: quick_xml::Error) -> PboError {
    PboError::Xml(e.to_string())
}

fn parse_body(xml: &str) -> PboResult<Vec<Block>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut out = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"p" => out.push(Block::Paragraph(parse_paragraph(&mut reader)?)),
                b"tbl" => out.push(Block::Table(parse_table(&mut reader)?)),
                _ => {}
            },
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn parse_paragraph(reader: &mut Reader<&[u8]>) -> PboResult<Paragraph> {
    let mut p = Paragraph::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"pPr" => p.style = Some(parse_ppr(reader)?),
                b"r" => p.runs.push(RunElement::Run(parse_run(reader)?)),
                b"hyperlink" => {
                    let rel_id = attr_qualified(&e, b"r:id", reader)?;
                    let anchor = attr_local(&e, b"anchor", reader)?;
                    let hl = parse_hyperlink(reader, rel_id, anchor)?;
                    p.runs.push(RunElement::Hyperlink(hl));
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"p" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(p)
}

fn parse_ppr(reader: &mut Reader<&[u8]>) -> PboResult<ParagraphStyle> {
    let mut style = ParagraphStyle::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"jc" => {
                    style.alignment = attr_local(&e, b"val", reader)?.map(|v| match v.as_str() {
                        "center" => Alignment::Center,
                        "right" | "end" => Alignment::Right,
                        "both" | "justify" => Alignment::Justify,
                        _ => Alignment::Left,
                    });
                }
                b"pStyle" => {
                    if let Some(id) = attr_local(&e, b"val", reader)? {
                            if let Some(rest) = id.strip_prefix("Heading")
                                && let Ok(n) = rest.parse::<u8>()
                            {
                                style.heading_level = Some(n);
                        }
                        style.style_id = Some(id);
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"pPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(style)
}

fn parse_run(reader: &mut Reader<&[u8]>) -> PboResult<Run> {
    let mut run = Run::default();
    let mut text = String::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"rPr" => run.style = Some(parse_rpr(reader)?),
                b"t" => {
                    let mut buf2 = Vec::new();
                    loop {
                        match reader.read_event_into(&mut buf2).map_err(xml_err)? {
                            Event::Text(t) => text.push_str(
                                &t.unescape().map_err(|e| PboError::Xml(e.to_string()))?,
                            ),
                            Event::End(ee) if local_name(ee.name().as_ref()) == b"t" => break,
                            Event::Eof => break,
                            _ => {}
                        }
                        buf2.clear();
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"tab" => text.push('\t'),
                b"br" => text.push('\n'),
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"r" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    run.text = text;
    Ok(run)
}

fn parse_rpr(reader: &mut Reader<&[u8]>) -> PboResult<RunStyle> {
    let mut style = RunStyle::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) | Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"b" => style.bold = on_off(&e, reader)?,
                b"i" => style.italic = on_off(&e, reader)?,
                b"u" => {
                    style.underline = attr_local(&e, b"val", reader)?.as_deref() != Some("none");
                }
                b"strike" | b"dstrike" => style.strike = on_off(&e, reader)?,
                b"sz" => {
                    style.font_size = attr_local(&e, b"val", reader)?
                        .and_then(|v| v.parse::<f32>().ok())
                        .map(|v| v / 2.0);
                }
                b"rFonts" => style.font_family = attr_local(&e, b"ascii", reader)?,
                b"color" => style.color = attr_local(&e, b"val", reader)?,
                b"highlight" => style.highlight = attr_local(&e, b"val", reader)?,
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"rPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(style)
}

fn parse_hyperlink(
    reader: &mut Reader<&[u8]>,
    rel_id: Option<String>,
    anchor: Option<String>,
) -> PboResult<Hyperlink> {
    let mut runs = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"r" => {
                runs.push(parse_run(reader)?);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"hyperlink" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(Hyperlink { rel_id, anchor, runs })
}

fn parse_table(reader: &mut Reader<&[u8]>) -> PboResult<Table> {
    let mut table = Table::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"tblGrid" => table.grid_cols = parse_tbl_grid(reader)?,
                b"tr" => table.rows.push(parse_tr(reader)?),
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"tbl" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(table)
}

fn parse_tbl_grid(reader: &mut Reader<&[u8]>) -> PboResult<Vec<i32>> {
    let mut cols = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e)
                if local_name(e.name().as_ref()) == b"gridCol" =>
            {
                cols.push(
                    attr_local(&e, b"w", reader)?
                        .and_then(|v| v.parse::<i32>().ok())
                        .unwrap_or(0),
                );
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"tblGrid" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(cols)
}

fn parse_tr(reader: &mut Reader<&[u8]>) -> PboResult<TableRow> {
    let mut row = TableRow::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"trPr" => parse_tr_pr(reader, &mut row)?,
                b"tc" => row.cells.push(parse_tc(reader)?),
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"tr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(row)
}

fn parse_tr_pr(reader: &mut Reader<&[u8]>, row: &mut TableRow) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) => match local_name(e.name().as_ref()) {
                b"trHeight" => {
                    row.height = attr_local(&e, b"val", reader)?
                        .and_then(|v| v.parse::<i32>().ok())
                        .unwrap_or(0);
                }
                b"tblHeader" => row.header = on_off(&e, reader)?,
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"trPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_tc(reader: &mut Reader<&[u8]>) -> PboResult<TableCell> {
    let mut cell = TableCell::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"tcPr" => parse_tc_pr(reader, &mut cell)?,
                b"p" => cell.content.push(Block::Paragraph(parse_paragraph(reader)?)),
                b"tbl" => cell.content.push(Block::Table(parse_table(reader)?)),
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

fn parse_tc_pr(reader: &mut Reader<&[u8]>, cell: &mut TableCell) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) => match local_name(e.name().as_ref()) {
                b"tcW" => {
                    cell.width = attr_local(&e, b"w", reader)?
                        .and_then(|v| v.parse::<i32>().ok())
                        .unwrap_or(0);
                }
                b"gridSpan" => {
                    cell.grid_span = attr_u32(&e, b"val", reader)?.unwrap_or(1);
                }
                b"vMerge" => {
                    cell.v_merge = match attr_local(&e, b"val", reader)?.as_deref() {
                        Some("restart") => VMerge::Restart,
                        _ => VMerge::Continue,
                    };
                }
                b"tcBorders" => {
                    let mut borders = cell.borders.take().unwrap_or_default();
                    parse_tc_borders(reader, &mut borders)?;
                    cell.borders = Some(borders);
                }
                b"shd" => cell.shading = attr_local(&e, b"fill", reader)?,
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"tcPr" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_tc_borders(reader: &mut Reader<&[u8]>, b: &mut CellBorders) -> PboResult<()> {
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Empty(e) | Event::Start(e) => {
                let name = local_name(e.name().as_ref()).to_vec();
                if matches!(name.as_slice(), b"top" | b"bottom" | b"left" | b"right") {
                    let style = attr_local(&e, b"val", reader)?.unwrap_or_default();
                    let size = attr_u32(&e, b"sz", reader)?.unwrap_or(4);
                    let color = attr_local(&e, b"color", reader)?;
                    let bl = Some(BorderLine { style, size, color });
                    match name.as_slice() {
                        b"top" => b.top = bl,
                        b"bottom" => b.bottom = bl,
                        b"left" => b.left = bl,
                        b"right" => b.right = bl,
                        _ => {}
                    }
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"tcBorders" => break,
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(())
}

fn parse_core(xml: &str) -> PboResult<DocumentMetadata> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut md = DocumentMetadata::default();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).map_err(xml_err)? {
            Event::Start(e) => {
                let qname = e.name().as_ref().to_vec();
                let mut text = String::new();
                let mut inner = Vec::new();
                loop {
                    match reader.read_event_into(&mut inner).map_err(xml_err)? {
                        Event::Text(t) => text.push_str(
                            &t.unescape().map_err(|e| PboError::Xml(e.to_string()))?,
                        ),
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
                    b"description" => md.description = t,
                    b"lastModifiedBy" => md.last_modified_by = t,
                    b"created" => md.created = t,
                    b"modified" => md.modified = t,
                    b"keywords" => md.keywords = t,
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

