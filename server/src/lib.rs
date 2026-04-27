//! PolarBearOffice HTTP API.
//!
//! The UI talks to the core engine through this server. Each format has
//! `parse` (bytes → JSON model), `write` (JSON model → bytes), and
//! `roundtrip` (bytes → bytes via the model) endpoints.
//!
//! Fixture endpoints expose `fixtures/` for end-to-end testing.

use std::{
    path::PathBuf,
    time::{SystemTime, UNIX_EPOCH},
};

use axum::{
    Json, Router,
    body::Bytes,
    extract::{DefaultBodyLimit, Path, Query, State},
    http::{HeaderMap, StatusCode, header},
    response::{IntoResponse, Response},
    routing::{get, post},
};
use serde::Serialize;
use tower_http::cors::CorsLayer;

use pbo_common::model::calc::CalcDocument;
use pbo_common::model::slides::SlidesDocument;
use pbo_common::model::writer::WriterDocument;

#[derive(Clone)]
pub struct AppState {
    pub fixtures_root: PathBuf,
    pub advanced_sessions_root: PathBuf,
}

impl AppState {
    pub fn from_workspace_root() -> Self {
        let crate_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        // <workspace>/server/Cargo.toml → workspace = parent of server.
        let workspace = crate_dir
            .parent()
            .map(PathBuf::from)
            .unwrap_or(crate_dir.clone());
        Self {
            fixtures_root: workspace.join("fixtures"),
            advanced_sessions_root: workspace.join("tmp").join("advanced-sessions"),
        }
    }

    pub fn with_fixtures_root(root: impl Into<PathBuf>) -> Self {
        let mut state = Self::from_workspace_root();
        state.fixtures_root = root.into();
        state
    }

    fn ensure_advanced_root(&self) -> Result<(), ApiError> {
        std::fs::create_dir_all(&self.advanced_sessions_root).map_err(|e| {
            ApiError::server(format!(
                "failed to create advanced editor storage at {}: {e}",
                self.advanced_sessions_root.display()
            ))
        })
    }

    fn advanced_session_path(&self, kind: &str, id: &str) -> PathBuf {
        self.advanced_sessions_root.join(format!("{id}.{kind}"))
    }
}

fn file_name_for_title(title: &str, kind: &str) -> String {
    let trimmed = title.trim();
    let base = if trimmed.is_empty() {
        format!("untitled.{kind}")
    } else if trimmed.to_ascii_lowercase().ends_with(&format!(".{kind}")) {
        trimmed.to_string()
    } else {
        format!("{trimmed}.{kind}")
    };
    sanitize_file_name(&base)
}

fn sanitize_file_name(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        let safe = ch.is_ascii_alphanumeric() || matches!(ch, '.' | '-' | '_' | ' ' | '(' | ')');
        out.push(if safe { ch } else { '_' });
    }
    let trimmed = out.trim().trim_matches('.');
    if trimmed.is_empty() {
        "untitled".to_string()
    } else {
        trimmed.to_string()
    }
}

fn advanced_session_id() -> String {
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default();
    format!("{:x}{:x}", std::process::id(), now.as_nanos())
}

fn validate_kind(kind: &str) -> Result<&'static str, ApiError> {
    match kind {
        "docx" => Ok(MIME_DOCX),
        "xlsx" => Ok(MIME_XLSX),
        "pptx" => Ok(MIME_PPTX),
        _ => Err(ApiError::bad_request(format!(
            "unknown document kind '{kind}'"
        ))),
    }
}

fn validate_session_id(id: &str) -> Result<(), ApiError> {
    if id.is_empty() || !id.chars().all(|ch| ch.is_ascii_hexdigit()) {
        return Err(ApiError::bad_request("invalid session id"));
    }
    Ok(())
}

fn content_disposition_inline(filename: &str) -> Result<header::HeaderValue, ApiError> {
    let value = format!("inline; filename=\"{}\"", filename.replace('"', "_"));
    header::HeaderValue::from_str(&value)
        .map_err(|e| ApiError::server(format!("invalid content disposition header: {e}")))
}

#[derive(Serialize)]
struct AdvancedSessionBody {
    id: String,
    kind: String,
    title: String,
    url: String,
}

#[derive(serde::Deserialize)]
struct AdvancedCreateQuery {
    title: Option<String>,
}

#[derive(serde::Deserialize)]
struct AdvancedFilePath {
    kind: String,
    id: String,
    name: String,
}

pub fn router(state: AppState) -> Router {
    Router::new()
        .route("/api/health", get(health))
        // DOCX
        .route("/api/docx/parse", post(parse_docx))
        .route("/api/docx/write", post(write_docx))
        .route("/api/docx/roundtrip", post(roundtrip_docx))
        // XLSX
        .route("/api/xlsx/parse", post(parse_xlsx))
        .route("/api/xlsx/write", post(write_xlsx))
        .route("/api/xlsx/roundtrip", post(roundtrip_xlsx))
        // PPTX
        .route("/api/pptx/parse", post(parse_pptx))
        .route("/api/pptx/render", post(render_pptx))
        .route("/api/pptx/write", post(write_pptx))
        .route("/api/pptx/roundtrip", post(roundtrip_pptx))
        // Advanced editor handoff
        .route("/api/advanced-session/{kind}", post(create_advanced_session))
        .route(
            "/api/advanced-session/{kind}/{id}/{name}",
            get(get_advanced_session_file),
        )
        // Fixtures
        .route("/api/fixtures", get(list_fixtures))
        .route("/api/fixtures/{kind}/{name}", get(get_fixture))
        .with_state(state)
        // Office documents can be large; lift the default 2 MiB body limit
        // up to 256 MiB for the upload endpoints.
        .layer(DefaultBodyLimit::max(256 * 1024 * 1024))
        .layer(CorsLayer::permissive())
}

// ── error type ────────────────────────────────────────

pub struct ApiError {
    status: StatusCode,
    message: String,
}

impl ApiError {
    fn bad_request(msg: impl Into<String>) -> Self {
        Self {
            status: StatusCode::BAD_REQUEST,
            message: msg.into(),
        }
    }
    fn not_found(msg: impl Into<String>) -> Self {
        Self {
            status: StatusCode::NOT_FOUND,
            message: msg.into(),
        }
    }
    fn server(msg: impl Into<String>) -> Self {
        Self {
            status: StatusCode::INTERNAL_SERVER_ERROR,
            message: msg.into(),
        }
    }
}

impl IntoResponse for ApiError {
    fn into_response(self) -> Response {
        let body = serde_json::json!({ "error": self.message });
        (self.status, Json(body)).into_response()
    }
}

impl From<pbo_common::PboError> for ApiError {
    fn from(e: pbo_common::PboError) -> Self {
        ApiError::bad_request(e.to_string())
    }
}

// ── handlers ──────────────────────────────────────────

#[derive(Serialize)]
struct HealthBody {
    status: &'static str,
    app: &'static str,
    formats: [&'static str; 3],
}

async fn health() -> Json<HealthBody> {
    Json(HealthBody {
        status: "ok",
        app: "PolarBearOffice",
        formats: ["docx", "xlsx", "pptx"],
    })
}

// DOCX
async fn parse_docx(body: Bytes) -> Result<Json<WriterDocument>, ApiError> {
    Ok(Json(pbo_docx::read_docx(&body)?))
}

async fn write_docx(Json(doc): Json<WriterDocument>) -> Result<Response, ApiError> {
    let bytes = pbo_docx::write_docx(&doc)?;
    Ok(binary_response(bytes, MIME_DOCX))
}

async fn roundtrip_docx(body: Bytes) -> Result<Response, ApiError> {
    let doc = pbo_docx::read_docx(&body)?;
    let bytes = pbo_docx::write_docx(&doc)?;
    Ok(binary_response(bytes, MIME_DOCX))
}

// XLSX
async fn parse_xlsx(body: Bytes) -> Result<Json<CalcDocument>, ApiError> {
    Ok(Json(pbo_xlsx::read_xlsx(&body)?))
}

async fn write_xlsx(Json(doc): Json<CalcDocument>) -> Result<Response, ApiError> {
    let bytes = pbo_xlsx::write_xlsx(&doc)?;
    Ok(binary_response(bytes, MIME_XLSX))
}

async fn roundtrip_xlsx(body: Bytes) -> Result<Response, ApiError> {
    let doc = pbo_xlsx::read_xlsx(&body)?;
    let bytes = pbo_xlsx::write_xlsx(&doc)?;
    Ok(binary_response(bytes, MIME_XLSX))
}

// PPTX
async fn parse_pptx(body: Bytes) -> Result<Json<SlidesDocument>, ApiError> {
    Ok(Json(pbo_pptx::read_pptx(&body)?))
}

async fn render_pptx(body: Bytes) -> Result<Json<SlidesDocument>, ApiError> {
    Ok(Json(pbo_pptx::read_pptx_rendered(&body)?))
}

async fn write_pptx(Json(doc): Json<SlidesDocument>) -> Result<Response, ApiError> {
    let bytes = pbo_pptx::write_pptx(&doc)?;
    Ok(binary_response(bytes, MIME_PPTX))
}

async fn roundtrip_pptx(body: Bytes) -> Result<Response, ApiError> {
    let doc = pbo_pptx::read_pptx(&body)?;
    let bytes = pbo_pptx::write_pptx(&doc)?;
    Ok(binary_response(bytes, MIME_PPTX))
}

// ── advanced editor handoff ─────────────────────────

async fn create_advanced_session(
    State(state): State<AppState>,
    Path(kind): Path<String>,
    Query(query): Query<AdvancedCreateQuery>,
    body: Bytes,
) -> Result<Json<AdvancedSessionBody>, ApiError> {
    validate_kind(&kind)?;
    state.ensure_advanced_root()?;

    let id = advanced_session_id();
    let title = file_name_for_title(query.title.as_deref().unwrap_or("untitled"), &kind);
    let path = state.advanced_session_path(&kind, &id);
    std::fs::write(&path, &body)
        .map_err(|e| ApiError::server(format!("failed to persist advanced editor session: {e}")))?;

    Ok(Json(AdvancedSessionBody {
        id: id.clone(),
        kind: kind.clone(),
        title: title.clone(),
        url: format!("/api/advanced-session/{kind}/{id}/{title}"),
    }))
}

async fn get_advanced_session_file(
    State(state): State<AppState>,
    Path(AdvancedFilePath { kind, id, name }): Path<AdvancedFilePath>,
) -> Result<Response, ApiError> {
    let mime = validate_kind(&kind)?;
    validate_session_id(&id)?;
    let filename = file_name_for_title(&name, &kind);
    let path = state.advanced_session_path(&kind, &id);
    let bytes = std::fs::read(&path).map_err(|e| {
        ApiError::not_found(format!(
            "advanced editor document not found: {}: {e}",
            path.display()
        ))
    })?;

    let mut headers = HeaderMap::new();
    headers.insert(header::CONTENT_TYPE, mime.parse().unwrap());
    headers.insert(
        header::CONTENT_DISPOSITION,
        content_disposition_inline(&filename)?,
    );
    Ok((StatusCode::OK, headers, bytes).into_response())
}

// ── fixtures ──────────────────────────────────────────

#[derive(Serialize)]
struct FixtureEntry {
    kind: &'static str,
    name: String,
    bytes: u64,
}

#[derive(Serialize)]
struct FixturesBody {
    items: Vec<FixtureEntry>,
}

async fn list_fixtures(State(state): State<AppState>) -> Result<Json<FixturesBody>, ApiError> {
    let mut items = Vec::new();
    for kind in ["docx", "xlsx", "pptx"] {
        let dir = state.fixtures_root.join(kind);
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if !path.is_file() {
                continue;
            }
            let name = match path.file_name().and_then(|s| s.to_str()) {
                Some(n) => n.to_string(),
                None => continue,
            };
            // Only expose files matching the kind extension.
            if !name.to_ascii_lowercase().ends_with(&format!(".{kind}")) {
                continue;
            }
            let bytes = entry.metadata().map(|m| m.len()).unwrap_or(0);
            items.push(FixtureEntry { kind, name, bytes });
        }
    }
    items.sort_by(|a, b| (a.kind, &a.name).cmp(&(b.kind, &b.name)));
    Ok(Json(FixturesBody { items }))
}

async fn get_fixture(
    State(state): State<AppState>,
    Path((kind, name)): Path<(String, String)>,
) -> Result<Response, ApiError> {
    let kind_ok = matches!(kind.as_str(), "docx" | "xlsx" | "pptx");
    if !kind_ok {
        return Err(ApiError::bad_request(format!(
            "unknown fixture kind '{kind}'"
        )));
    }
    if name.contains('/') || name.contains('\\') || name.contains("..") {
        return Err(ApiError::bad_request("invalid fixture name"));
    }
    let path = state.fixtures_root.join(&kind).join(&name);
    let bytes = std::fs::read(&path)
        .map_err(|e| ApiError::not_found(format!("fixture not found: {}: {e}", path.display())))?;
    let mime = match kind.as_str() {
        "docx" => MIME_DOCX,
        "xlsx" => MIME_XLSX,
        "pptx" => MIME_PPTX,
        _ => "application/octet-stream",
    };
    Ok(binary_response(bytes, mime))
}

// ── helpers ───────────────────────────────────────────

const MIME_DOCX: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const MIME_XLSX: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const MIME_PPTX: &str = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

fn binary_response(bytes: Vec<u8>, mime: &str) -> Response {
    let mut headers = HeaderMap::new();
    headers.insert(header::CONTENT_TYPE, mime.parse().unwrap());
    (StatusCode::OK, headers, bytes).into_response()
}

// Re-export an init helper so the test harness can find it.
impl Default for AppState {
    fn default() -> Self {
        Self::from_workspace_root()
    }
}
