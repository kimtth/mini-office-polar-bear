// DOCX writer: emits a WriterDocument as a .docx OPC package.

use pbo_common::b64;
use pbo_common::model::writer::{
    Alignment, Block, DocumentMetadata, Hyperlink, Paragraph, ParagraphStyle, Run, RunElement,
    RunStyle, Table, TableCell, TableRow, VMerge, WriterDocument,
};
use pbo_common::{PboError, PboResult};
use pbo_ooxml::xml::{escape_attr, escape_text};
use pbo_ooxml::{ContentTypes, OpcPackage, Relationships};

const NS_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_OFFICE_DOC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE: &str = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const CT_DOC: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";

pub fn write_docx(doc: &WriterDocument) -> PboResult<Vec<u8>> {
    let mut pkg = OpcPackage::default();
    let mut ct = ContentTypes::default();
    ct.add_default("rels", "application/vnd.openxmlformats-package.relationships+xml");
    ct.add_default("xml", "application/xml");
    ct.add_override("/word/document.xml", CT_DOC);
    ct.add_override("/docProps/core.xml", CT_CORE);

    // Root rels
    let mut root_rels = Relationships::default();
    root_rels.add("rId1", REL_OFFICE_DOC, "word/document.xml");
    root_rels.add("rId2", REL_CORE, "docProps/core.xml");
    pkg.set_part("_rels/.rels", root_rels.to_xml().into_bytes());
    pkg.set_part("docProps/core.xml", render_core(&doc.metadata).into_bytes());

    // Document rels (hyperlinks + images).
    let mut doc_rels = Relationships::default();
    for h in &doc.resources.hyperlinks {
        let mut r = pbo_ooxml::Relationship {
            id: h.id.clone(),
            rel_type: REL_HYPERLINK.to_string(),
            target: h.url.clone(),
            target_mode: Some("External".to_string()),
        };
        if r.id.is_empty() {
            r.id = format!("hId{}", doc_rels.0.len() + 1);
        }
        doc_rels.0.push(r);
    }
    for img in &doc.resources.images {
        let ext = if let Some(sub) = img.mime.strip_prefix("image/") {
            match sub {
                "jpeg" => "jpg",
                "svg+xml" => "svg",
                other => other,
            }
        } else {
            "bin"
        };
        let file_name = format!("{}.{ext}", img.id);
        ct.add_default(ext, &img.mime);
        let bytes = b64::decode(&img.data_base64)
            .map_err(|e| PboError::InvalidStructure(format!("bad image b64: {e}")))?;
        pkg.set_part(format!("word/media/{file_name}"), bytes);
        doc_rels.add(&img.id, REL_IMAGE, &format!("media/{file_name}"));
    }
    pkg.set_part("word/_rels/document.xml.rels", doc_rels.to_xml().into_bytes());

    pkg.set_part("word/document.xml", render_document(doc).into_bytes());
    pkg.set_part("[Content_Types].xml", ct.to_xml().into_bytes());
    pkg.to_bytes()
}

fn render_document(doc: &WriterDocument) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<w:document xmlns:w="{NS_W}" xmlns:r="{NS_R}"><w:body>"#
    ));
    for b in &doc.body {
        render_block(b, &mut xml);
    }
    xml.push_str("</w:body></w:document>");
    xml
}

fn render_block(b: &Block, out: &mut String) {
    match b {
        Block::Paragraph(p) => render_paragraph(p, out),
        Block::Table(t) => render_table(t, out),
        Block::PageBreak => {
            out.push_str(r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#);
        }
    }
}

fn render_paragraph(p: &Paragraph, out: &mut String) {
    out.push_str("<w:p>");
    if let Some(style) = &p.style {
        render_ppr(style, out);
    }
    for r in &p.runs {
        match r {
            RunElement::Run(run) => render_run(run, out),
            RunElement::Hyperlink(h) => render_hyperlink(h, out),
            RunElement::InlineImage(_) => {}
        }
    }
    out.push_str("</w:p>");
}

fn render_ppr(style: &ParagraphStyle, out: &mut String) {
    out.push_str("<w:pPr>");
    if let Some(id) = &style.style_id {
        out.push_str(&format!(
            r#"<w:pStyle w:val="{}"/>"#,
            escape_attr(id)
        ));
    } else if let Some(n) = style.heading_level {
        out.push_str(&format!(r#"<w:pStyle w:val="Heading{n}"/>"#));
    }
    if let Some(a) = style.alignment {
        let v = match a {
            Alignment::Left => "left",
            Alignment::Center => "center",
            Alignment::Right => "right",
            Alignment::Justify => "both",
        };
        out.push_str(&format!(r#"<w:jc w:val="{v}"/>"#));
    }
    out.push_str("</w:pPr>");
}

fn render_run(run: &Run, out: &mut String) {
    out.push_str("<w:r>");
    if let Some(s) = &run.style {
        render_rpr(s, out);
    }
    // Split text on \n and \t for proper OOXML handling.
    let mut first = true;
    for part in run.text.split_inclusive(['\n', '\t']) {
        let (text, delim) = match part.chars().last() {
            Some('\n') => (&part[..part.len() - 1], Some('\n')),
            Some('\t') => (&part[..part.len() - 1], Some('\t')),
            _ => (part, None),
        };
        if !text.is_empty() {
            if !first {
                // fallthrough
            }
            out.push_str(&format!(
                r#"<w:t xml:space="preserve">{}</w:t>"#,
                escape_text(text)
            ));
        }
        match delim {
            Some('\n') => out.push_str("<w:br/>"),
            Some('\t') => out.push_str("<w:tab/>"),
            _ => {}
        }
        first = false;
    }
    out.push_str("</w:r>");
}

fn render_rpr(s: &RunStyle, out: &mut String) {
    out.push_str("<w:rPr>");
    if s.bold {
        out.push_str("<w:b/>");
    }
    if s.italic {
        out.push_str("<w:i/>");
    }
    if s.underline {
        out.push_str(r#"<w:u w:val="single"/>"#);
    }
    if s.strike {
        out.push_str("<w:strike/>");
    }
    if let Some(sz) = s.font_size {
        let half = (sz * 2.0).round() as i32;
        out.push_str(&format!(r#"<w:sz w:val="{half}"/>"#));
    }
    if let Some(f) = &s.font_family {
        out.push_str(&format!(
            r#"<w:rFonts w:ascii="{}" w:hAnsi="{}"/>"#,
            escape_attr(f),
            escape_attr(f)
        ));
    }
    if let Some(c) = &s.color {
        out.push_str(&format!(r#"<w:color w:val="{}"/>"#, escape_attr(c)));
    }
    if let Some(h) = &s.highlight {
        out.push_str(&format!(r#"<w:highlight w:val="{}"/>"#, escape_attr(h)));
    }
    out.push_str("</w:rPr>");
}

fn render_hyperlink(h: &Hyperlink, out: &mut String) {
    out.push_str("<w:hyperlink");
    if let Some(rid) = &h.rel_id {
        out.push_str(&format!(r#" r:id="{}""#, escape_attr(rid)));
    }
    if let Some(a) = &h.anchor {
        out.push_str(&format!(r#" w:anchor="{}""#, escape_attr(a)));
    }
    out.push('>');
    for r in &h.runs {
        render_run(r, out);
    }
    out.push_str("</w:hyperlink>");
}

fn render_table(t: &Table, out: &mut String) {
    out.push_str("<w:tbl><w:tblPr/>");
    if !t.grid_cols.is_empty() {
        out.push_str("<w:tblGrid>");
        for w in &t.grid_cols {
            out.push_str(&format!(r#"<w:gridCol w:w="{w}"/>"#));
        }
        out.push_str("</w:tblGrid>");
    }
    for row in &t.rows {
        render_row(row, out);
    }
    out.push_str("</w:tbl>");
}

fn render_row(r: &TableRow, out: &mut String) {
    out.push_str("<w:tr>");
    if r.header || r.height != 0 {
        out.push_str("<w:trPr>");
        if r.height != 0 {
            out.push_str(&format!(r#"<w:trHeight w:val="{}"/>"#, r.height));
        }
        if r.header {
            out.push_str("<w:tblHeader/>");
        }
        out.push_str("</w:trPr>");
    }
    for c in &r.cells {
        render_cell(c, out);
    }
    out.push_str("</w:tr>");
}

fn render_cell(c: &TableCell, out: &mut String) {
    out.push_str("<w:tc><w:tcPr>");
    if c.width != 0 {
        out.push_str(&format!(
            r#"<w:tcW w:w="{}" w:type="dxa"/>"#,
            c.width
        ));
    }
    if c.grid_span > 1 {
        out.push_str(&format!(r#"<w:gridSpan w:val="{}"/>"#, c.grid_span));
    }
    match c.v_merge {
        VMerge::None => {}
        VMerge::Restart => out.push_str(r#"<w:vMerge w:val="restart"/>"#),
        VMerge::Continue => out.push_str("<w:vMerge/>"),
    }
    if let Some(sh) = &c.shading {
        out.push_str(&format!(
            r#"<w:shd w:val="clear" w:color="auto" w:fill="{}"/>"#,
            escape_attr(sh)
        ));
    }
    out.push_str("</w:tcPr>");
    if c.content.is_empty() {
        out.push_str("<w:p/>");
    } else {
        for b in &c.content {
            render_block(b, out);
        }
    }
    out.push_str("</w:tc>");
}

fn render_core(md: &DocumentMetadata) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"#,
    );
    if let Some(t) = &md.title {
        xml.push_str(&format!("<dc:title>{}</dc:title>", escape_text(t)));
    }
    if let Some(c) = &md.author {
        xml.push_str(&format!("<dc:creator>{}</dc:creator>", escape_text(c)));
    }
    if let Some(d) = &md.description {
        xml.push_str(&format!(
            "<dc:description>{}</dc:description>",
            escape_text(d)
        ));
    }
    if let Some(k) = &md.keywords {
        xml.push_str(&format!("<cp:keywords>{}</cp:keywords>", escape_text(k)));
    }
    if let Some(l) = &md.last_modified_by {
        xml.push_str(&format!(
            "<cp:lastModifiedBy>{}</cp:lastModifiedBy>",
            escape_text(l)
        ));
    }
    if let Some(c) = &md.created {
        xml.push_str(&format!(
            r#"<dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>"#,
            escape_text(c)
        ));
    }
    if let Some(m) = &md.modified {
        xml.push_str(&format!(
            r#"<dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>"#,
            escape_text(m)
        ));
    }
    xml.push_str("</cp:coreProperties>");
    xml
}
