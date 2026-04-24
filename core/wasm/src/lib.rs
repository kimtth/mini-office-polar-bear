//! WebAssembly entry point for PolarBearOffice.
//!
//! Exposes document operations to JavaScript via wasm-bindgen.
//! Feature-gated so apps can load only the format they need.

use wasm_bindgen::prelude::*;

/// Initialize the WASM module (set up panic hook, logging, etc.)
#[wasm_bindgen(start)]
pub fn init() {
    #[cfg(feature = "console_error_panic_hook")]
    console_error_panic_hook::set_once();
}

// ── DOCX ──────────────────────────────────────────────

#[cfg(feature = "docx")]
#[wasm_bindgen]
pub fn parse_docx(data: &[u8]) -> Result<String, JsValue> {
    let doc = pbo_docx::read_docx(data).map_err(|e| JsValue::from_str(&e.to_string()))?;
    serde_json::to_string(&doc).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[cfg(feature = "docx")]
#[wasm_bindgen]
pub fn create_docx(json: &str) -> Result<Vec<u8>, JsValue> {
    let doc: pbo_common::model::writer::WriterDocument =
        serde_json::from_str(json).map_err(|e| JsValue::from_str(&e.to_string()))?;
    pbo_docx::write_docx(&doc).map_err(|e| JsValue::from_str(&e.to_string()))
}

// ── PPTX ──────────────────────────────────────────────

#[cfg(feature = "pptx")]
#[wasm_bindgen]
pub fn parse_pptx(data: &[u8]) -> Result<String, JsValue> {
    let doc = pbo_pptx::read_pptx(data).map_err(|e| JsValue::from_str(&e.to_string()))?;
    serde_json::to_string(&doc).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[cfg(feature = "pptx")]
#[wasm_bindgen]
pub fn create_pptx(json: &str) -> Result<Vec<u8>, JsValue> {
    let doc: pbo_common::model::slides::SlidesDocument =
        serde_json::from_str(json).map_err(|e| JsValue::from_str(&e.to_string()))?;
    pbo_pptx::write_pptx(&doc).map_err(|e| JsValue::from_str(&e.to_string()))
}

// ── XLSX ──────────────────────────────────────────────

#[cfg(feature = "xlsx")]
#[wasm_bindgen]
pub fn parse_xlsx(data: &[u8]) -> Result<String, JsValue> {
    let doc = pbo_xlsx::read_xlsx(data).map_err(|e| JsValue::from_str(&e.to_string()))?;
    serde_json::to_string(&doc).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[cfg(feature = "xlsx")]
#[wasm_bindgen]
pub fn create_xlsx(json: &str) -> Result<Vec<u8>, JsValue> {
    let doc: pbo_common::model::calc::CalcDocument =
        serde_json::from_str(json).map_err(|e| JsValue::from_str(&e.to_string()))?;
    pbo_xlsx::write_xlsx(&doc).map_err(|e| JsValue::from_str(&e.to_string()))
}
