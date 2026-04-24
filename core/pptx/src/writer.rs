// PPTX writer: emits a SlidesDocument as a .pptx OPC package.

use pbo_common::b64;
use pbo_common::model::slides::{
    Connector, ConnectorType, GroupShape, Picture, PresentationMetadata, Shape, ShapeType, Slide,
    SlideElement, SlidesDocument, SlidesImage, TextAlign, TextBox,
};
use pbo_common::{PboError, PboResult};
use pbo_ooxml::xml::{escape_attr, escape_text};
use pbo_ooxml::{ContentTypes, OpcPackage, Relationship, Relationships};

const NS_P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const REL_OFFICE_DOC: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE: &str = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_SLIDE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const REL_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_NOTES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const REL_NOTES_SLIDE_BACK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

const CT_PRES: &str = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
const CT_SLIDE: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_NOTES: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml";
const CT_CORE: &str = "application/vnd.openxmlformats-package.core-properties+xml";

pub fn write_pptx(doc: &SlidesDocument) -> PboResult<Vec<u8>> {
    let mut pkg = OpcPackage::default();
    let mut ct = ContentTypes::default();
    ct.add_default("rels", "application/vnd.openxmlformats-package.relationships+xml");
    ct.add_default("xml", "application/xml");
    ct.add_override("/ppt/presentation.xml", CT_PRES);
    ct.add_override("/docProps/core.xml", CT_CORE);

    // Root rels
    let mut root_rels = Relationships::default();
    root_rels.add("rId1", REL_OFFICE_DOC, "ppt/presentation.xml");
    root_rels.add("rId2", REL_CORE, "docProps/core.xml");
    pkg.set_part("_rels/.rels", root_rels.to_xml().into_bytes());
    pkg.set_part("docProps/core.xml", render_core(&doc.metadata).into_bytes());

    // Media — write images once and register extensions in Content Types.
    for (idx, img) in doc.images.iter().enumerate() {
        let ext = if let Some(sub) = img.mime.strip_prefix("image/") {
            match sub {
                "jpeg" => "jpg",
                "svg+xml" => "svg",
                other => other,
            }
        } else {
            "bin"
        };
        ct.add_default(ext, &img.mime);
        let bytes = b64::decode(&img.data_base64)
            .map_err(|e| PboError::InvalidStructure(format!("bad image b64: {e}")))?;
        let file_name = format!("image{}.{ext}", idx + 1);
        pkg.set_part(format!("ppt/media/{file_name}"), bytes);
    }

    // presentation rels: one slide per deck + notesSize handled elsewhere.
    let mut pres_rels = Relationships::default();
    for (idx, _) in doc.slides.iter().enumerate() {
        pres_rels.add(
            &format!("rId{}", idx + 1),
            REL_SLIDE,
            &format!("slides/slide{}.xml", idx + 1),
        );
    }
    pkg.set_part(
        "ppt/_rels/presentation.xml.rels",
        pres_rels.to_xml().into_bytes(),
    );
    pkg.set_part(
        "ppt/presentation.xml",
        render_presentation(doc).into_bytes(),
    );

    // Slides
    for (idx, slide) in doc.slides.iter().enumerate() {
        let n = idx + 1;
        let slide_path = format!("ppt/slides/slide{n}.xml");
        let slide_rels_path = format!("ppt/slides/_rels/slide{n}.xml.rels");
        ct.add_override(&format!("/{}", slide_path), CT_SLIDE);

        let mut slide_rels = Relationships::default();
        // Image rels — one per unique image id used in this slide (simple: link all).
        for (img_idx, img) in doc.images.iter().enumerate() {
            let ext = ext_for(&img.mime);
            slide_rels.add(
                &img_id_rid(img_idx),
                REL_IMAGE,
                &format!("../media/image{}.{ext}", img_idx + 1),
            );
        }

        if slide.notes.is_some() {
            slide_rels.add(
                &format!("rIdNotes{n}"),
                REL_NOTES,
                &format!("../notesSlides/notesSlide{n}.xml"),
            );
            // Notes slide
            let notes_path = format!("ppt/notesSlides/notesSlide{n}.xml");
            let notes_rels_path = format!("ppt/notesSlides/_rels/notesSlide{n}.xml.rels");
            ct.add_override(&format!("/{}", notes_path), CT_NOTES);
            let mut notes_rels = Relationships::default();
            notes_rels.0.push(Relationship {
                id: "rIdSlide".into(),
                rel_type: REL_NOTES_SLIDE_BACK.into(),
                target: format!("../slides/slide{n}.xml"),
                target_mode: None,
            });
            pkg.set_part(notes_rels_path, notes_rels.to_xml().into_bytes());
            pkg.set_part(
                notes_path,
                render_notes_slide(slide.notes.as_deref().unwrap_or("")).into_bytes(),
            );
        }

        pkg.set_part(slide_rels_path, slide_rels.to_xml().into_bytes());
        pkg.set_part(slide_path, render_slide(slide, &doc.images).into_bytes());
    }

    pkg.set_part("[Content_Types].xml", ct.to_xml().into_bytes());
    pkg.to_bytes()
}

fn ext_for(mime: &str) -> &'static str {
    match mime {
        "image/png" => "png",
        "image/jpeg" => "jpg",
        "image/gif" => "gif",
        "image/bmp" => "bmp",
        "image/svg+xml" => "svg",
        _ => "bin",
    }
}

fn img_id_rid(idx: usize) -> String {
    format!("rIdImg{}", idx + 1)
}

fn render_presentation(doc: &SlidesDocument) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    let w = if doc.slide_width > 0.0 { doc.slide_width as i64 } else { 9144000 };
    let h = if doc.slide_height > 0.0 { doc.slide_height as i64 } else { 6858000 };
    xml.push_str(&format!(
        r#"<p:presentation xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"><p:sldIdLst>"#
    ));
    for (idx, _) in doc.slides.iter().enumerate() {
        xml.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            256 + idx,
            idx + 1
        ));
    }
    xml.push_str(&format!(
        r#"</p:sldIdLst><p:sldSz cx="{w}" cy="{h}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    ));
    xml
}

fn render_slide(slide: &Slide, images: &[SlidesImage]) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<p:sld xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"><p:cSld><p:spTree>"#
    ));
    xml.push_str(
        r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    let mut id_counter = 2u32;
    for el in &slide.elements {
        render_element(el, images, &mut xml, &mut id_counter);
    }
    xml.push_str("</p:spTree></p:cSld></p:sld>");
    xml
}

fn render_element(
    el: &SlideElement,
    images: &[SlidesImage],
    out: &mut String,
    id: &mut u32,
) {
    match el {
        SlideElement::TextBox(tb) => render_textbox(tb, out, id),
        SlideElement::Shape(s) => render_shape(s, out, id),
        SlideElement::Picture(p) => render_picture(p, images, out, id),
        SlideElement::Connector(c) => render_connector(c, out, id),
        SlideElement::Group(g) => render_group(g, images, out, id),
        SlideElement::Table(_) => {}
    }
}

fn next_id(id: &mut u32) -> u32 {
    let v = *id;
    *id += 1;
    v
}

fn render_textbox(tb: &TextBox, out: &mut String, id: &mut u32) {
    let n = next_id(id);
    out.push_str(&format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{n}" name="TextBox {n}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#
    ));
    render_sp_pr(tb.x, tb.y, tb.width, tb.height, tb.rotation, "rect", out);
    render_tx_body(&tb.text, tb.bold, tb.italic, tb.font_size, tb.color.as_deref(), tb.align, out);
    out.push_str("</p:sp>");
}

fn render_shape(s: &Shape, out: &mut String, id: &mut u32) {
    let n = next_id(id);
    let prst = prst_for(s.shape_type);
    out.push_str(&format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{n}" name="Shape {n}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#
    ));
    render_sp_pr(s.x, s.y, s.width, s.height, s.rotation, prst, out);
    if let Some(t) = &s.text {
        render_tx_body(&t.text, t.bold, t.italic, t.font_size, t.color.as_deref(), t.align, out);
    } else {
        out.push_str(r#"<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>"#);
    }
    out.push_str("</p:sp>");
}

fn render_picture(p: &Picture, images: &[SlidesImage], out: &mut String, id: &mut u32) {
    let idx = images.iter().position(|i| i.id == p.image_id).unwrap_or(0);
    let rid = img_id_rid(idx);
    let n = next_id(id);
    out.push_str(&format!(
        r#"<p:pic><p:nvPicPr><p:cNvPr id="{n}" name="{}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>"#,
        escape_attr(p.name.as_deref().unwrap_or("Picture"))
    ));
    out.push_str(&format!(
        r#"<p:blipFill><a:blip r:embed="{}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>"#,
        escape_attr(&rid)
    ));
    render_sp_pr(p.x, p.y, p.width, p.height, p.rotation, "rect", out);
    out.push_str("</p:pic>");
}

fn render_connector(c: &Connector, out: &mut String, id: &mut u32) {
    let n = next_id(id);
    let prst = match c.connector_type {
        ConnectorType::Straight => "line",
        ConnectorType::Bent => "bentConnector3",
        ConnectorType::Curved => "curvedConnector3",
    };
    out.push_str(&format!(
        r#"<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{n}" name="Connector {n}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>"#
    ));
    out.push_str("<p:spPr><a:xfrm");
    if c.flip_h {
        out.push_str(r#" flipH="1""#);
    }
    if c.flip_v {
        out.push_str(r#" flipV="1""#);
    }
    out.push('>');
    out.push_str(&format!(
        r#"<a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/>"#,
        c.x as i64, c.y as i64, c.width as i64, c.height as i64
    ));
    out.push_str(&format!(
        r#"</a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>"#
    ));
}

fn render_group(g: &GroupShape, images: &[SlidesImage], out: &mut String, id: &mut u32) {
    let n = next_id(id);
    out.push_str(&format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{n}" name="Group {n}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#
    ));
    out.push_str(&format!(
        r#"<p:grpSpPr><a:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/><a:chOff x="{}" y="{}"/><a:chExt cx="{}" cy="{}"/></a:xfrm></p:grpSpPr>"#,
        g.x as i64,
        g.y as i64,
        g.width as i64,
        g.height as i64,
        g.x as i64,
        g.y as i64,
        g.width as i64,
        g.height as i64
    ));
    for el in &g.elements {
        render_element(el, images, out, id);
    }
    out.push_str("</p:grpSp>");
}

fn render_sp_pr(x: f64, y: f64, w: f64, h: f64, rot: i32, prst: &str, out: &mut String) {
    out.push_str("<p:spPr><a:xfrm");
    if rot != 0 {
        out.push_str(&format!(r#" rot="{rot}""#));
    }
    out.push('>');
    out.push_str(&format!(
        r#"<a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/>"#,
        x as i64, y as i64, w as i64, h as i64
    ));
    out.push_str(&format!(
        r#"</a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom></p:spPr>"#
    ));
}

fn render_tx_body(
    text: &str,
    bold: bool,
    italic: bool,
    size: Option<f32>,
    color: Option<&str>,
    align: Option<TextAlign>,
    out: &mut String,
) {
    out.push_str(r#"<p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>"#);
    let paras: Vec<&str> = text.split('\n').collect();
    if paras.is_empty() {
        out.push_str("<a:p/>");
    }
    for para in paras {
        out.push_str("<a:p>");
        if let Some(a) = align {
            let v = match a {
                TextAlign::Left => "l",
                TextAlign::Center => "ctr",
                TextAlign::Right => "r",
                TextAlign::Justify => "just",
            };
            out.push_str(&format!(r#"<a:pPr algn="{v}"/>"#));
        }
        if para.is_empty() {
            out.push_str("<a:endParaRPr/>");
        } else {
            out.push_str("<a:r><a:rPr lang=\"en-US\"");
            if bold {
                out.push_str(r#" b="1""#);
            }
            if italic {
                out.push_str(r#" i="1""#);
            }
            if let Some(sz) = size {
                let hundredths = (sz * 100.0).round() as i32;
                out.push_str(&format!(r#" sz="{hundredths}""#));
            }
            out.push('>');
            if let Some(c) = color {
                out.push_str(&format!(
                    r#"<a:solidFill><a:srgbClr val="{}"/></a:solidFill>"#,
                    escape_attr(c)
                ));
            }
            out.push_str("</a:rPr>");
            out.push_str(&format!("<a:t>{}</a:t>", escape_text(para)));
            out.push_str("</a:r>");
        }
        out.push_str("</a:p>");
    }
    out.push_str("</p:txBody>");
}

fn prst_for(s: ShapeType) -> &'static str {
    match s {
        ShapeType::Rectangle => "rect",
        ShapeType::RoundedRectangle => "roundRect",
        ShapeType::Ellipse => "ellipse",
        ShapeType::Triangle => "triangle",
        ShapeType::RightTriangle => "rtTriangle",
        ShapeType::Diamond => "diamond",
        ShapeType::Parallelogram => "parallelogram",
        ShapeType::Trapezoid => "trapezoid",
        ShapeType::Pentagon => "pentagon",
        ShapeType::Hexagon => "hexagon",
        ShapeType::Octagon => "octagon",
        ShapeType::Star5 => "star5",
        ShapeType::Arrow => "rightArrow",
        ShapeType::Line => "line",
        ShapeType::Callout => "wedgeRectCallout",
        ShapeType::Other => "rect",
    }
}

fn render_notes_slide(text: &str) -> String {
    let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<p:notes xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"><p:cSld><p:spTree>"#
    ));
    xml.push_str(
        r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    );
    xml.push_str(r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:spPr/>"#);
    xml.push_str(r#"<p:txBody><a:bodyPr/><a:lstStyle/>"#);
    for para in text.split('\n') {
        xml.push_str("<a:p>");
        if !para.is_empty() {
            xml.push_str(&format!(
                r#"<a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r>"#,
                escape_text(para)
            ));
        }
        xml.push_str("</a:p>");
    }
    xml.push_str("</p:txBody></p:sp></p:spTree></p:cSld></p:notes>");
    xml
}

fn render_core(md: &PresentationMetadata) -> String {
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
