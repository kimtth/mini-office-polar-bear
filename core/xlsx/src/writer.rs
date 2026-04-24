//! Write a `CalcDocument` to a `.xlsx` byte stream.

use pbo_common::model::calc::{
    CalcDocument, Cell, CellBorder, CellFill, CellFont, CellValue, CellXf, NumFmt, Sheet,
    SheetHyperlink, WorkbookStyles,
};
use pbo_common::PboResult;
use pbo_ooxml::xml::escape_attr;
use pbo_ooxml::{ContentTypes, OpcPackage, Relationships};
use std::collections::HashMap;

const NS_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_OFFICE_DOC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE: &str = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_SHEET: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const REL_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const REL_HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const CT_WORKBOOK: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const CT_SHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const CT_STYLES: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const CT_SHARED_STRINGS: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";

pub fn write_xlsx(doc: &CalcDocument) -> PboResult<Vec<u8>> {
    let mut pkg = OpcPackage::default();
    let mut ct = ContentTypes::default();
    ct.add_default("rels", "application/vnd.openxmlformats-package.relationships+xml");
    ct.add_default("xml", "application/xml");
    ct.add_override("/xl/workbook.xml", CT_WORKBOOK);
    ct.add_override("/xl/styles.xml", CT_STYLES);
    ct.add_override("/xl/sharedStrings.xml", CT_SHARED_STRINGS);
    ct.add_override("/docProps/core.xml", CT_CORE);

    // Root rels
    let mut root_rels = Relationships::default();
    root_rels.add("rId1", REL_OFFICE_DOC, "xl/workbook.xml");
    root_rels.add("rId2", REL_CORE, "docProps/core.xml");
    pkg.set_part("_rels/.rels", root_rels.to_xml().into_bytes());

    pkg.set_part("docProps/core.xml", render_core(&doc.metadata).into_bytes());

    // Shared strings: collect all unique texts.
    let (sst, sst_xml) = build_shared_strings(doc);
    pkg.set_part("xl/sharedStrings.xml", sst_xml.into_bytes());

    // Styles
    pkg.set_part("xl/styles.xml", render_styles(&doc.styles).into_bytes());

    // Workbook rels
    let mut wb_rels = Relationships::default();
    let mut rid = 1u32;
    let mut sheet_rids = Vec::with_capacity(doc.sheets.len());
    for (i, _s) in doc.sheets.iter().enumerate() {
        let rel_id = format!("rId{rid}");
        wb_rels.add(&rel_id, REL_SHEET, &format!("worksheets/sheet{}.xml", i + 1));
        sheet_rids.push(rel_id);
        rid += 1;
    }
    let styles_rid = format!("rId{rid}");
    wb_rels.add(&styles_rid, REL_STYLES, "styles.xml");
    rid += 1;
    let sst_rid = format!("rId{rid}");
    wb_rels.add(&sst_rid, REL_SHARED_STRINGS, "sharedStrings.xml");
    pkg.set_part(
        "xl/_rels/workbook.xml.rels",
        wb_rels.to_xml().into_bytes(),
    );

    // Workbook
    pkg.set_part(
        "xl/workbook.xml",
        render_workbook(doc, &sheet_rids).into_bytes(),
    );

    // Sheets
    for (i, sheet) in doc.sheets.iter().enumerate() {
        let path = format!("xl/worksheets/sheet{}.xml", i + 1);
        ct.add_override(&format!("/{path}"), CT_SHEET);

        // Per-sheet rels for hyperlinks.
        let mut sheet_rels = Relationships::default();
        let mut link_rids: Vec<String> = Vec::new();
        for (j, link) in sheet.hyperlinks.iter().enumerate() {
            let id = format!("rId{}", j + 1);
            sheet_rels.add(&id, REL_HYPERLINK, &link.target);
            link_rids.push(id);
        }
        if !sheet.hyperlinks.is_empty() {
            pkg.set_part(
                format!("xl/worksheets/_rels/sheet{}.xml.rels", i + 1),
                sheet_rels.to_xml().into_bytes(),
            );
        }

        pkg.set_part(path, render_sheet(sheet, &sst, &link_rids).into_bytes());
    }

    pkg.set_part("[Content_Types].xml", ct.to_xml().into_bytes());
    pkg.to_bytes()
}

// ── shared strings ──────────────────────────────────
fn build_shared_strings(doc: &CalcDocument) -> (HashMap<String, u32>, String) {
    let mut map: HashMap<String, u32> = HashMap::new();
    let mut list: Vec<String> = Vec::new();
    for sheet in &doc.sheets {
        for row in sheet.cells.values() {
            for cell in row.values() {
                if let CellValue::Text(t) = &cell.value && !map.contains_key(t) {
                    map.insert(t.clone(), list.len() as u32);
                    list.push(t.clone());
                }
            }
        }
    }
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<sst xmlns="{NS_MAIN}" count="{}" uniqueCount="{}">"#,
        list.len(),
        list.len()
    ));
    for s in &list {
        xml.push_str(&format!(
            r#"<si><t xml:space="preserve">{}</t></si>"#,
            escape_attr(s)
        ));
    }
    xml.push_str("</sst>");
    (map, xml)
}

// ── workbook.xml ────────────────────────────────────
fn render_workbook(doc: &CalcDocument, sheet_rids: &[String]) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_R}"><sheets>"#
    ));
    for (i, sheet) in doc.sheets.iter().enumerate() {
        let rid = sheet_rids.get(i).cloned().unwrap_or_default();
        xml.push_str(&format!(
            r#"<sheet name="{}" sheetId="{}" r:id="{}"/>"#,
            escape_attr(&sheet.name),
            i + 1,
            rid
        ));
    }
    xml.push_str("</sheets>");
    if !doc.defined_names.is_empty() {
        xml.push_str("<definedNames>");
        for dn in &doc.defined_names {
            match dn.local_sheet_id {
                Some(id) => xml.push_str(&format!(
                    r#"<definedName name="{}" localSheetId="{id}">{}</definedName>"#,
                    escape_attr(&dn.name),
                    escape_attr(&dn.formula)
                )),
                None => xml.push_str(&format!(
                    r#"<definedName name="{}">{}</definedName>"#,
                    escape_attr(&dn.name),
                    escape_attr(&dn.formula)
                )),
            }
        }
        xml.push_str("</definedNames>");
    }
    xml.push_str("</workbook>");
    xml
}

// ── styles.xml ──────────────────────────────────────
fn render_styles(styles: &WorkbookStyles) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(r#"<styleSheet xmlns="{NS_MAIN}">"#));

    if !styles.num_fmts.is_empty() {
        xml.push_str(&format!(
            r#"<numFmts count="{}">"#,
            styles.num_fmts.len()
        ));
        for n in &styles.num_fmts {
            xml.push_str(&render_num_fmt(n));
        }
        xml.push_str("</numFmts>");
    }

    let fonts = if styles.fonts.is_empty() {
        vec![CellFont::default()]
    } else {
        styles.fonts.clone()
    };
    xml.push_str(&format!(r#"<fonts count="{}">"#, fonts.len()));
    for f in &fonts {
        xml.push_str(&render_font(f));
    }
    xml.push_str("</fonts>");

    let fills = if styles.fills.is_empty() {
        vec![
            CellFill {
                pattern_type: "none".into(),
                ..Default::default()
            },
            CellFill {
                pattern_type: "gray125".into(),
                ..Default::default()
            },
        ]
    } else {
        styles.fills.clone()
    };
    xml.push_str(&format!(r#"<fills count="{}">"#, fills.len()));
    for f in &fills {
        xml.push_str(&render_fill(f));
    }
    xml.push_str("</fills>");

    let borders = if styles.borders.is_empty() {
        vec![CellBorder::default()]
    } else {
        styles.borders.clone()
    };
    xml.push_str(&format!(r#"<borders count="{}">"#, borders.len()));
    for b in &borders {
        xml.push_str(&render_border(b));
    }
    xml.push_str("</borders>");

    xml.push_str(r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#);

    let xfs = if styles.cell_xfs.is_empty() {
        vec![CellXf::default()]
    } else {
        styles.cell_xfs.clone()
    };
    xml.push_str(&format!(r#"<cellXfs count="{}">"#, xfs.len()));
    for xf in &xfs {
        xml.push_str(&render_xf(xf));
    }
    xml.push_str("</cellXfs>");
    xml.push_str(r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#);
    xml.push_str("</styleSheet>");
    xml
}

fn render_num_fmt(n: &NumFmt) -> String {
    format!(
        r#"<numFmt numFmtId="{}" formatCode="{}"/>"#,
        n.id,
        escape_attr(&n.format_code)
    )
}

fn render_font(f: &CellFont) -> String {
    let mut s = String::from("<font>");
    if let Some(sz) = f.size {
        s.push_str(&format!(r#"<sz val="{sz}"/>"#));
    }
    if let Some(name) = &f.name {
        s.push_str(&format!(r#"<name val="{}"/>"#, escape_attr(name)));
    }
    if f.bold {
        s.push_str("<b/>");
    }
    if f.italic {
        s.push_str("<i/>");
    }
    if f.underline {
        s.push_str("<u/>");
    }
    if f.strike {
        s.push_str("<strike/>");
    }
    if let Some(c) = &f.color {
        s.push_str(&format!(r#"<color rgb="{}"/>"#, escape_attr(c)));
    }
    s.push_str("</font>");
    s
}

fn render_fill(f: &CellFill) -> String {
    let pt = if f.pattern_type.is_empty() {
        "none"
    } else {
        &f.pattern_type
    };
    let mut s = format!(r#"<fill><patternFill patternType="{pt}">"#);
    if let Some(c) = &f.fg_color {
        s.push_str(&format!(r#"<fgColor rgb="{}"/>"#, escape_attr(c)));
    }
    if let Some(c) = &f.bg_color {
        s.push_str(&format!(r#"<bgColor rgb="{}"/>"#, escape_attr(c)));
    }
    s.push_str("</patternFill></fill>");
    s
}

fn render_border(b: &CellBorder) -> String {
    let mut s = String::from("<border>");
    for (tag, edge) in [
        ("left", &b.left),
        ("right", &b.right),
        ("top", &b.top),
        ("bottom", &b.bottom),
    ] {
        match edge {
            Some(e) => {
                if let Some(color) = &e.color {
                    s.push_str(&format!(
                        r#"<{tag} style="{}"><color rgb="{}"/></{tag}>"#,
                        escape_attr(&e.style),
                        escape_attr(color)
                    ));
                } else {
                    s.push_str(&format!(r#"<{tag} style="{}"/>"#, escape_attr(&e.style)));
                }
            }
            None => s.push_str(&format!("<{tag}/>")),
        }
    }
    s.push_str("<diagonal/></border>");
    s
}

fn render_xf(xf: &CellXf) -> String {
    let mut s = String::from("<xf");
    if let Some(v) = xf.num_fmt_id {
        s.push_str(&format!(r#" numFmtId="{v}""#));
    } else {
        s.push_str(r#" numFmtId="0""#);
    }
    s.push_str(&format!(
        r#" fontId="{}""#,
        xf.font_id.unwrap_or(0)
    ));
    s.push_str(&format!(
        r#" fillId="{}""#,
        xf.fill_id.unwrap_or(0)
    ));
    s.push_str(&format!(
        r#" borderId="{}""#,
        xf.border_id.unwrap_or(0)
    ));
    s.push_str(r#" xfId="0""#);
    if xf.apply_number_format {
        s.push_str(r#" applyNumberFormat="1""#);
    }
    if xf.apply_font {
        s.push_str(r#" applyFont="1""#);
    }
    if xf.apply_fill {
        s.push_str(r#" applyFill="1""#);
    }
    if xf.apply_border {
        s.push_str(r#" applyBorder="1""#);
    }
    s.push_str("/>");
    s
}

// ── sheet ───────────────────────────────────────────
fn render_sheet(sheet: &Sheet, sst: &HashMap<String, u32>, link_rids: &[String]) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_R}">"#
    ));

    if let Some(view) = &sheet.view {
        xml.push_str(r#"<sheetViews><sheetView workbookViewId="0""#);
        if !view.show_grid_lines {
            xml.push_str(r#" showGridLines="0""#);
        }
        if view.zoom_scale != 0 && view.zoom_scale != 100 {
            xml.push_str(&format!(r#" zoomScale="{}""#, view.zoom_scale));
        }
        xml.push('>');
        if view.freeze_row > 0 || view.freeze_col > 0 {
            let top_cell = cell_ref(view.freeze_col + 1, view.freeze_row + 1);
            xml.push_str(&format!(
                r#"<pane xSplit="{}" ySplit="{}" topLeftCell="{top_cell}" state="frozen"/>"#,
                view.freeze_col, view.freeze_row
            ));
        }
        xml.push_str("</sheetView></sheetViews>");
    }

    if !sheet.columns.is_empty() {
        xml.push_str("<cols>");
        for c in &sheet.columns {
            xml.push_str(&format!(
                r#"<col min="{}" max="{}" width="{}""#,
                c.min, c.max, c.width
            ));
            if c.custom_width {
                xml.push_str(r#" customWidth="1""#);
            }
            if c.hidden {
                xml.push_str(r#" hidden="1""#);
            }
            xml.push_str("/>");
        }
        xml.push_str("</cols>");
    }

    xml.push_str("<sheetData>");
    for (row_idx, row) in &sheet.cells {
        xml.push_str(&format!(r#"<row r="{row_idx}""#));
        if let Some(h) = sheet.row_heights.get(row_idx) {
            xml.push_str(&format!(r#" ht="{h}" customHeight="1""#));
        }
        xml.push('>');
        for (col_idx, cell) in row {
            xml.push_str(&render_cell(*row_idx, *col_idx, cell, sst));
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData>");

    if !sheet.merges.is_empty() {
        xml.push_str(&format!(
            r#"<mergeCells count="{}">"#,
            sheet.merges.len()
        ));
        for m in &sheet.merges {
            xml.push_str(&format!(
                r#"<mergeCell ref="{}:{}"/>"#,
                cell_ref(m.col_start, m.row_start),
                cell_ref(m.col_end, m.row_end)
            ));
        }
        xml.push_str("</mergeCells>");
    }

    if !sheet.hyperlinks.is_empty() {
        xml.push_str("<hyperlinks>");
        for (i, h) in sheet.hyperlinks.iter().enumerate() {
            let rid = link_rids.get(i).cloned().unwrap_or_default();
            xml.push_str(&render_hyperlink(h, &rid));
        }
        xml.push_str("</hyperlinks>");
    }
    xml.push_str("</worksheet>");
    xml
}

fn render_hyperlink(h: &SheetHyperlink, rid: &str) -> String {
    let mut s = format!(
        r#"<hyperlink ref="{}" r:id="{rid}""#,
        escape_attr(&h.cell_ref)
    );
    if let Some(d) = &h.display {
        s.push_str(&format!(r#" display="{}""#, escape_attr(d)));
    }
    s.push_str("/>");
    s
}

fn render_cell(row: u32, col: u32, cell: &Cell, sst: &HashMap<String, u32>) -> String {
    let r = cell_ref(col, row);
    let mut s = format!(r#"<c r="{r}""#);
    if let Some(sid) = cell.style_id {
        s.push_str(&format!(r#" s="{sid}""#));
    }
    match &cell.value {
        CellValue::Empty => {
            if cell.formula.is_none() && cell.style_id.is_none() {
                return String::new();
            }
            s.push_str("/>");
            return s;
        }
        CellValue::Text(t) => match sst.get(t) {
            Some(idx) => {
                s.push_str(&format!(r#" t="s"><v>{idx}</v>"#));
            }
            None => {
                s.push_str(&format!(
                    r#" t="inlineStr"><is><t xml:space="preserve">{}</t></is>"#,
                    escape_attr(t)
                ));
            }
        },
        CellValue::Number(n) => {
            s.push_str(&format!("><v>{n}</v>"));
        }
        CellValue::Boolean(b) => {
            s.push_str(&format!(
                r#" t="b"><v>{}</v>"#,
                if *b { 1 } else { 0 }
            ));
        }
        CellValue::Date(d) => {
            s.push_str(&format!(
                r#" t="str"><v>{}</v>"#,
                escape_attr(d)
            ));
        }
        CellValue::Error(err) => {
            s.push_str(&format!(
                r#" t="e"><v>{}</v>"#,
                escape_attr(err)
            ));
        }
    }
    if let Some(f) = &cell.formula {
        s.push_str(&format!("<f>{}</f>", escape_attr(f)));
    }
    s.push_str("</c>");
    s
}

fn cell_ref(col: u32, row: u32) -> String {
    let mut letters = String::new();
    let mut n = col;
    while n > 0 {
        let rem = (n - 1) % 26;
        letters.insert(0, (b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    format!("{letters}{row}")
}

// ── core.xml ────────────────────────────────────────
fn render_core(md: &pbo_common::model::calc::SpreadsheetMetadata) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"#,
    );
    if let Some(t) = &md.title {
        xml.push_str(&format!("<dc:title>{}</dc:title>", escape_attr(t)));
    }
    if let Some(c) = &md.author {
        xml.push_str(&format!("<dc:creator>{}</dc:creator>", escape_attr(c)));
    }
    if let Some(l) = &md.last_modified_by {
        xml.push_str(&format!(
            "<cp:lastModifiedBy>{}</cp:lastModifiedBy>",
            escape_attr(l)
        ));
    }
    if let Some(c) = &md.created {
        xml.push_str(&format!(
            r#"<dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>"#,
            escape_attr(c)
        ));
    }
    if let Some(m) = &md.modified {
        xml.push_str(&format!(
            r#"<dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>"#,
            escape_attr(m)
        ));
    }
    xml.push_str("</cp:coreProperties>");
    xml
}