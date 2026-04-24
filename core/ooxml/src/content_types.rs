//! `[Content_Types].xml` parsing and generation.

use pbo_common::{PboError, PboResult};
use quick_xml::Reader;
use quick_xml::events::Event;

const CT_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefaultType {
    pub extension: String,
    pub content_type: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OverrideType {
    pub part_name: String,
    pub content_type: String,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ContentTypes {
    pub defaults: Vec<DefaultType>,
    pub overrides: Vec<OverrideType>,
}

impl ContentTypes {
    pub fn parse(xml: &str) -> PboResult<Self> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut out = ContentTypes::default();
        let mut buf = Vec::new();
        loop {
            match reader
                .read_event_into(&mut buf)
                .map_err(|e| PboError::Xml(e.to_string()))?
            {
                Event::Empty(e) | Event::Start(e) => {
                    let local = e.name().as_ref().to_vec();
                    if local == b"Default" {
                        let mut ext = String::new();
                        let mut ct = String::new();
                        for attr in e.attributes().flatten() {
                            let v = attr
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(|e| PboError::Xml(e.to_string()))?
                                .into_owned();
                            match attr.key.as_ref() {
                                b"Extension" => ext = v,
                                b"ContentType" => ct = v,
                                _ => {}
                            }
                        }
                        if !ext.is_empty() {
                            out.defaults.push(DefaultType {
                                extension: ext,
                                content_type: ct,
                            });
                        }
                    } else if local == b"Override" {
                        let mut pn = String::new();
                        let mut ct = String::new();
                        for attr in e.attributes().flatten() {
                            let v = attr
                                .decode_and_unescape_value(reader.decoder())
                                .map_err(|e| PboError::Xml(e.to_string()))?
                                .into_owned();
                            match attr.key.as_ref() {
                                b"PartName" => pn = v,
                                b"ContentType" => ct = v,
                                _ => {}
                            }
                        }
                        if !pn.is_empty() {
                            out.overrides.push(OverrideType {
                                part_name: pn,
                                content_type: ct,
                            });
                        }
                    }
                }
                Event::Eof => break,
                _ => {}
            }
            buf.clear();
        }
        Ok(out)
    }

    pub fn to_xml(&self) -> String {
        let mut s = String::with_capacity(256);
        s.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        s.push_str(&format!("<Types xmlns=\"{CT_NS}\">"));
        for d in &self.defaults {
            s.push_str(&format!(
                "<Default Extension=\"{}\" ContentType=\"{}\"/>",
                crate::xml::escape_attr(&d.extension),
                crate::xml::escape_attr(&d.content_type)
            ));
        }
        for o in &self.overrides {
            s.push_str(&format!(
                "<Override PartName=\"{}\" ContentType=\"{}\"/>",
                crate::xml::escape_attr(&o.part_name),
                crate::xml::escape_attr(&o.content_type)
            ));
        }
        s.push_str("</Types>");
        s
    }

    pub fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults.push(DefaultType {
            extension: extension.to_string(),
            content_type: content_type.to_string(),
        });
    }

    pub fn add_override(&mut self, part_name: &str, content_type: &str) {
        self.overrides.push(OverrideType {
            part_name: part_name.to_string(),
            content_type: content_type.to_string(),
        });
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_basic() {
        let mut ct = ContentTypes::default();
        ct.add_default("rels", "application/vnd.openxmlformats-package.relationships+xml");
        ct.add_default("xml", "application/xml");
        ct.add_override(
            "/word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        );
        let xml = ct.to_xml();
        let parsed = ContentTypes::parse(&xml).unwrap();
        assert_eq!(ct, parsed);
    }
}
