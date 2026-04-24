//! Read and write OPC (ZIP-based) packages.

use pbo_common::{PboError, PboResult};
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};

/// An in-memory representation of an OPC package (ZIP of XML parts).
#[derive(Debug, Clone, Default)]
pub struct OpcPackage {
    pub parts: HashMap<String, Vec<u8>>,
}

impl OpcPackage {
    /// Open a package from raw bytes.
    pub fn from_bytes(data: &[u8]) -> PboResult<Self> {
        let reader = Cursor::new(data);
        let mut archive =
            zip::ZipArchive::new(reader).map_err(|e| PboError::Zip(e.to_string()))?;

        let mut parts = HashMap::new();
        for i in 0..archive.len() {
            let mut file = archive
                .by_index(i)
                .map_err(|e| PboError::Zip(e.to_string()))?;
            let name = file.name().to_string();
            let mut buf = Vec::new();
            file.read_to_end(&mut buf)?;
            parts.insert(name, buf);
        }

        Ok(Self { parts })
    }

    /// Serialize the package to bytes.
    pub fn to_bytes(&self) -> PboResult<Vec<u8>> {
        let buf = Vec::new();
        let cursor = Cursor::new(buf);
        let mut writer = zip::ZipWriter::new(cursor);

        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated);

        for (name, data) in &self.parts {
            writer
                .start_file(name, options)
                .map_err(|e| PboError::Zip(e.to_string()))?;
            writer.write_all(data)?;
        }

        let cursor = writer
            .finish()
            .map_err(|e| PboError::Zip(e.to_string()))?;
        Ok(cursor.into_inner())
    }

    /// Get a part's content as a UTF-8 string.
    pub fn get_part_str(&self, name: &str) -> PboResult<&str> {
        let data = self
            .parts
            .get(name)
            .ok_or_else(|| PboError::InvalidStructure(format!("Missing part: {name}")))?;
        std::str::from_utf8(data)
            .map_err(|e| PboError::InvalidStructure(format!("Invalid UTF-8 in {name}: {e}")))
    }

    /// Insert or replace a part.
    pub fn set_part(&mut self, name: impl Into<String>, data: Vec<u8>) {
        self.parts.insert(name.into(), data);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trip_empty_package() {
        let mut pkg = OpcPackage::default();
        pkg.set_part("hello.xml", b"<hello/>".to_vec());

        let bytes = pkg.to_bytes().unwrap();
        let pkg2 = OpcPackage::from_bytes(&bytes).unwrap();

        assert_eq!(pkg2.get_part_str("hello.xml").unwrap(), "<hello/>");
    }
}
