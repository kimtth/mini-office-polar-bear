//! XML reading/writing helpers shared across format crates.

use pbo_common::{PboError, PboResult};
use quick_xml::Reader;
use quick_xml::events::BytesStart;

/// Strip a namespace prefix (`w:p` → `p`).
pub fn local_name(qname: &[u8]) -> &[u8] {
    match qname.iter().position(|&b| b == b':') {
        Some(i) => &qname[i + 1..],
        None => qname,
    }
}

/// Look up an attribute by local name (prefix ignored).
pub fn attr_local(
    e: &BytesStart<'_>,
    key: &[u8],
    reader: &Reader<&[u8]>,
) -> PboResult<Option<String>> {
    for attr in e.attributes().flatten() {
        if local_name(attr.key.as_ref()) == key {
            return Ok(Some(
                attr.decode_and_unescape_value(reader.decoder())
                    .map_err(|err| PboError::Xml(err.to_string()))?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

/// Look up an attribute by full qualified name (e.g. `r:id`).
pub fn attr_qualified(
    e: &BytesStart<'_>,
    key: &[u8],
    reader: &Reader<&[u8]>,
) -> PboResult<Option<String>> {
    for attr in e.attributes().flatten() {
        if attr.key.as_ref() == key {
            return Ok(Some(
                attr.decode_and_unescape_value(reader.decoder())
                    .map_err(|err| PboError::Xml(err.to_string()))?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

/// Parse an integer attribute, returning `None` if absent or malformed.
pub fn attr_i32(
    e: &BytesStart<'_>,
    key: &[u8],
    reader: &Reader<&[u8]>,
) -> PboResult<Option<i32>> {
    Ok(attr_local(e, key, reader)?.and_then(|v| v.parse::<i32>().ok()))
}

pub fn attr_i64(
    e: &BytesStart<'_>,
    key: &[u8],
    reader: &Reader<&[u8]>,
) -> PboResult<Option<i64>> {
    Ok(attr_local(e, key, reader)?.and_then(|v| v.parse::<i64>().ok()))
}

pub fn attr_u32(
    e: &BytesStart<'_>,
    key: &[u8],
    reader: &Reader<&[u8]>,
) -> PboResult<Option<u32>> {
    Ok(attr_local(e, key, reader)?.and_then(|v| v.parse::<u32>().ok()))
}

pub fn attr_f64(
    e: &BytesStart<'_>,
    key: &[u8],
    reader: &Reader<&[u8]>,
) -> PboResult<Option<f64>> {
    Ok(attr_local(e, key, reader)?.and_then(|v| v.parse::<f64>().ok()))
}

/// Interpret an attribute with DOCX-style on/off semantics.
///
/// In WordprocessingML, `<w:b/>` means "on" and `<w:b w:val="0"/>` or `<w:b w:val="false"/>`
/// means "off". Any other val (or missing val) means "on".
pub fn on_off(e: &BytesStart<'_>, reader: &Reader<&[u8]>) -> PboResult<bool> {
    match attr_local(e, b"val", reader)? {
        None => Ok(true),
        Some(v) => Ok(!matches!(v.as_str(), "0" | "false" | "off")),
    }
}
