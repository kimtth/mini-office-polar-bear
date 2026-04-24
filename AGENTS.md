# PolarBearOffice Bootstrap

- Goal: locally open and edit `.docx`, `.xlsx`, and `.pptx` with a Rust core and WASM UI.
- Refactor behavior from `references/` into Rust under `core/`; do not extend the reference core as the product engine.
- Keep the canonical editing model in Rust; OOXML crates are import/export adapters.
- Build the web app in `ui/`; keep WASM interop in `ui/src/wasm/`.
- Server scope is local development and local file support only; no public cloud assumptions.
- Architecture and core decisions live in `docs/ARCHITECTURE.md`.

