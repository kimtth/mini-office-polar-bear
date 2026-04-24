//! OPC relationship (`.rels`) parsing and generation.

use pbo_common::{PboError, PboResult};
use quick_xml::Reader;
use quick_xml::events::Event;

const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Relationships(pub Vec<Relationship>);

impl Relationships {
    pub fn parse(xml: &str) -> PboResult<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);
        let mut out = Vec::new();
        let mut buf = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buf)
                .map_err(|e| PboError::Xml(e.to_string()))?
            {
                Event::Empty(e) | Event::Start(e) if e.name().as_ref() == b"Relationship" => {
                    let mut id = String::new();
                    let mut rt = String::new();
                    let mut tgt = String::new();
                    let mut mode = None;
                    for attr in e.attributes().flatten() {
                        let v = attr
                            .decode_and_unescape_value(reader.decoder())
                            .map_err(|e| PboError::Xml(e.to_string()))?
                            .into_owned();
                        match attr.key.as_ref() {
                            b"Id" => id = v,
                            b"Type" => rt = v,
                            b"Target" => tgt = v,
                            b"TargetMode" => mode = Some(v),
                            _ => {}
                        }
                    }
                    if !id.is_empty() {
                        out.push(Relationship {
                            id,
                            rel_type: rt,
                            target: tgt,
                            target_mode: mode,
                        });
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(Relationships(out))
    }

    pub fn to_xml(&self) -> String {
        let mut s = String::with_capacity(256);
        s.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        s.push_str(&format!("<Relationships xmlns=\"{REL_NS}\">"));
        for r in &self.0 {
            s.push_str(&format!(
                "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"",
                crate::xml::escape_attr(&r.id),
                crate::xml::escape_attr(&r.rel_type),
                crate::xml::escape_attr(&r.target)
            ));
            if let Some(m) = &r.target_mode {
                s.push_str(&format!(" TargetMode=\"{}\"", crate::xml::escape_attr(m)));
            }
            s.push_str("/>");
        }
        s.push_str("</Relationships>");
        s
    }

    pub fn add(&mut self, id: &str, rel_type: &str, target: &str) {
        self.0.push(Relationship {
            id: id.to_string(),
            rel_type: rel_type.to_string(),
            target: target.to_string(),
            target_mode: None,
        });
    }

    pub fn find_by_type(&self, rel_type: &str) -> Option<&Relationship> {
        self.0.iter().find(|r| r.rel_type == rel_type)
    }

    pub fn find_by_id(&self, id: &str) -> Option<&Relationship> {
        self.0.iter().find(|r| r.id == id)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_rels() {
        let mut rels = Relationships::default();
        rels.add(
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "word/document.xml",
        );
        let xml = rels.to_xml();
        let parsed = Relationships::parse(&xml).unwrap();
        assert_eq!(rels, parsed);
    }
}
