//! End-to-end tests for the PolarBearOffice HTTP API.
//!
//! Drives the axum router in-process (no network) using `tower::ServiceExt::oneshot`,
//! pumping every fixture under `fixtures/{docx,xlsx,pptx}` through the
//! `parse → write` round trip and asserting that the returned bytes are a
//! valid OPC package the core can read again.

use std::path::{Path, PathBuf};

use axum::{
    body::{Body, Bytes},
    http::{Request, StatusCode, header},
};
use http_body_util::BodyExt;
use tower::ServiceExt;

use pbo_server::{AppState, router};

fn workspace_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .expect("server crate has parent")
        .to_path_buf()
}

fn fixtures_dir(kind: &str) -> PathBuf {
    workspace_root().join("fixtures").join(kind)
}

fn list_fixture_files(kind: &str) -> Vec<PathBuf> {
    let dir = fixtures_dir(kind);
    let Ok(entries) = std::fs::read_dir(&dir) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    let suffix = format!(".{kind}");
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_file()
            && path
                .file_name()
                .and_then(|s| s.to_str())
                .map(|n| n.to_ascii_lowercase().ends_with(&suffix))
                .unwrap_or(false)
        {
            out.push(path);
        }
    }
    out.sort();
    out
}

async fn collect_body(body: Body) -> Bytes {
    body.collect().await.expect("body collect").to_bytes()
}

fn app() -> axum::Router {
    router(AppState::with_fixtures_root(
        workspace_root().join("fixtures"),
    ))
}

async fn post_bytes(path: &str, mime: &str, bytes: Vec<u8>) -> (StatusCode, Bytes) {
    let req = Request::builder()
        .method("POST")
        .uri(path)
        .header(header::CONTENT_TYPE, mime)
        .body(Body::from(bytes))
        .expect("build request");
    let resp = app().oneshot(req).await.expect("oneshot");
    let status = resp.status();
    let bytes = collect_body(resp.into_body()).await;
    (status, bytes)
}

async fn get_path(path: &str) -> (StatusCode, Bytes) {
    let req = Request::builder()
        .method("GET")
        .uri(path)
        .body(Body::empty())
        .expect("build request");
    let resp = app().oneshot(req).await.expect("oneshot");
    let status = resp.status();
    let bytes = collect_body(resp.into_body()).await;
    (status, bytes)
}

fn looks_like_zip(bytes: &[u8]) -> bool {
    bytes.len() >= 4 && &bytes[..2] == b"PK"
}

fn assert_roundtrip(kind: &str, original: &Path, emitted: &[u8]) {
    assert!(
        looks_like_zip(emitted),
        "{}: emitted bytes are not a ZIP/OPC package ({} bytes)",
        original.display(),
        emitted.len()
    );
    // Re-parse the emitted bytes through the core to confirm the writer
    // produced something the reader still accepts.
    match kind {
        "docx" => {
            pbo_docx::read_docx(emitted)
                .unwrap_or_else(|e| panic!("re-read docx failed for {}: {e}", original.display()));
        }
        "xlsx" => {
            pbo_xlsx::read_xlsx(emitted)
                .unwrap_or_else(|e| panic!("re-read xlsx failed for {}: {e}", original.display()));
        }
        "pptx" => {
            pbo_pptx::read_pptx(emitted)
                .unwrap_or_else(|e| panic!("re-read pptx failed for {}: {e}", original.display()));
        }
        _ => unreachable!(),
    }
}

#[tokio::test]
async fn health_endpoint_responds() {
    let (status, body) = get_path("/api/health").await;
    assert_eq!(status, StatusCode::OK);
    let v: serde_json::Value = serde_json::from_slice(&body).expect("json");
    assert_eq!(v["status"], "ok");
}

#[tokio::test]
async fn list_fixtures_includes_all_kinds() {
    let (status, body) = get_path("/api/fixtures").await;
    assert_eq!(status, StatusCode::OK);
    let v: serde_json::Value = serde_json::from_slice(&body).expect("json");
    let items = v["items"].as_array().expect("items array");
    let kinds: std::collections::HashSet<&str> =
        items.iter().filter_map(|i| i["kind"].as_str()).collect();
    // Each kind should appear if the fixtures directory has files for it.
    for kind in ["docx", "xlsx", "pptx"] {
        if !list_fixture_files(kind).is_empty() {
            assert!(
                kinds.contains(kind),
                "missing fixture kind '{kind}' in listing"
            );
        }
    }
}

#[tokio::test]
async fn rejects_path_traversal_in_fixture_name() {
    let (status, _) = get_path("/api/fixtures/docx/..%2Fevil.txt").await;
    assert_eq!(status, StatusCode::BAD_REQUEST);
}

async fn run_format_e2e(kind: &str, parse_path: &str, roundtrip_path: &str, mime: &str) {
    let files = list_fixture_files(kind);
    if files.is_empty() {
        eprintln!("[skip] no fixtures present for {kind}");
        return;
    }
    for fixture in files {
        let bytes =
            std::fs::read(&fixture).unwrap_or_else(|e| panic!("read {}: {e}", fixture.display()));

        // 1. parse fixture via server.
        let (status, body) = post_bytes(parse_path, mime, bytes.clone()).await;
        assert_eq!(
            status,
            StatusCode::OK,
            "{}: parse failed; body={}",
            fixture.display(),
            String::from_utf8_lossy(&body)
        );
        let _model: serde_json::Value = serde_json::from_slice(&body).expect("parse returned JSON");

        // 2. roundtrip via server.
        let (status, body) = post_bytes(roundtrip_path, mime, bytes).await;
        assert_eq!(
            status,
            StatusCode::OK,
            "{}: roundtrip failed; body={}",
            fixture.display(),
            String::from_utf8_lossy(&body)
        );
        assert_roundtrip(kind, &fixture, &body);
    }
}

#[tokio::test]
async fn docx_fixtures_round_trip() {
    run_format_e2e(
        "docx",
        "/api/docx/parse",
        "/api/docx/roundtrip",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    .await;
}

#[tokio::test]
async fn xlsx_fixtures_round_trip() {
    run_format_e2e(
        "xlsx",
        "/api/xlsx/parse",
        "/api/xlsx/roundtrip",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    .await;
}

#[tokio::test]
async fn pptx_fixtures_round_trip() {
    run_format_e2e(
        "pptx",
        "/api/pptx/parse",
        "/api/pptx/roundtrip",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    .await;
}

#[tokio::test]
async fn write_endpoint_accepts_parsed_model() {
    // Sanity: parse → JSON → /write should produce ZIP bytes.
    let docs = list_fixture_files("docx");
    let Some(fixture) = docs.into_iter().next() else {
        eprintln!("[skip] no docx fixtures");
        return;
    };
    let bytes = std::fs::read(&fixture).expect("read");
    let (status, body) = post_bytes("/api/docx/parse", "application/octet-stream", bytes).await;
    assert_eq!(status, StatusCode::OK);

    let (status, body) = post_bytes("/api/docx/write", "application/json", body.to_vec()).await;
    assert_eq!(
        status,
        StatusCode::OK,
        "write failed: {}",
        String::from_utf8_lossy(&body)
    );
    assert!(looks_like_zip(&body), "write did not return a ZIP");
}

#[tokio::test]
async fn advanced_session_round_trip_serves_uploaded_bytes() {
    let payload = b"advanced-editor-doc".to_vec();
    let req = Request::builder()
        .method("POST")
        .uri("/api/advanced-session/docx?title=demo.docx")
        .header(
            header::CONTENT_TYPE,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        .body(Body::from(payload.clone()))
        .expect("build request");
    let resp = app().oneshot(req).await.expect("oneshot");
    assert_eq!(resp.status(), StatusCode::OK);
    let body = collect_body(resp.into_body()).await;
    let json: serde_json::Value = serde_json::from_slice(&body).expect("json body");
    let url = json["url"].as_str().expect("url in response");

    let (status, served) = get_path(url).await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(served, Bytes::from(payload));
}
