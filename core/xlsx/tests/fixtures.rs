//! Integration tests that exercise real fixture OOXML files.

use std::path::{Path, PathBuf};

fn fixtures_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("..")
        .join("..")
        .join("fixtures")
}

#[test]
fn read_xlsx_fixtures() {
    let dir = fixtures_dir().join("xlsx");
    let mut count = 0;
    for entry in std::fs::read_dir(&dir).unwrap() {
        let path = entry.unwrap().path();
        if path.extension().and_then(|s| s.to_str()) != Some("xlsx") {
            continue;
        }
        let data = std::fs::read(&path).unwrap();
        let doc = pbo_xlsx::read_xlsx(&data)
            .unwrap_or_else(|e| panic!("failed to read {}: {e}", path.display()));
        let _bytes = pbo_xlsx::write_xlsx(&doc).unwrap();
        count += 1;
    }
    assert!(count > 0, "no xlsx fixtures found in {}", dir.display());
}
