//! Integration tests that exercise real fixture OOXML files.

use std::path::{Path, PathBuf};

fn fixtures_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("fixtures")
}

#[test]
fn read_docx_fixtures() {
    let dir = fixtures_dir().join("docx");
    let mut count = 0;
    for entry in std::fs::read_dir(&dir).unwrap() {
        let path = entry.unwrap().path();
        if path.extension().and_then(|s| s.to_str()) != Some("docx") {
            continue;
        }
        let data = std::fs::read(&path).unwrap();
        let doc = pbo_docx::read_docx(&data)
            .unwrap_or_else(|e| panic!("failed to read {}: {e}", path.display()));
        // Round-trip: write should succeed.
        let _bytes = pbo_docx::write_docx(&doc).unwrap();
        count += 1;
    }
    assert!(count > 0, "no docx fixtures found in {}", dir.display());
}
