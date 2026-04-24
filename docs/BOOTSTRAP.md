# PolarBearOffice After Initialization

## Product Goal

PolarBearOffice is an open-source Microsoft Office 365 style application built with Rust, a local Axum server, and a React UI.

The primary goal is to open and edit these OOXML formats:
- `.docx`
- `.xlsx`
- `.pptx`

The application should feel like a modern browser office suite, but the core document engine must be implemented in Rust and exposed to the web UI through a narrow integration layer.

## Working Order

1. Read `docs/ARCHITECTURE.md` before making architecture or model changes.
2. Use material under `references/` as reference input, not as product code.
3. If a reference subtree contains a `graphify-out/` directory or generated graphify artifacts, consult that summary first and then verify against the raw reference sources.
4. If `graphify-out/` is missing for a reference subtree, fall back to the raw reference files and keep notes local to the current task.

## Source Strategy

### Core Engine
- Treat `references/core` and related OnlyOffice material under `references/` as behavior and architecture reference material.
- Refactor the office core behavior into native Rust crates under `core/`.
- Do not extend the JavaScript core from the reference implementation as the long-term engine.
- Keep the live editing model in Rust, not in the reference JavaScript code.
- Use the existing crate split as the target architecture:
  - `core/common`: shared types, errors, and canonical internal document models
  - `core/ooxml`: ZIP package, relationships, content types, and XML helpers
  - `core/docx`: DOCX import and export
  - `core/pptx`: PPTX import and export
  - `core/xlsx`: XLSX import and export
  - `core/wasm`: optional thin wasm-bindgen bridge for browser consumption

### UI and Web App
- Build the user-facing web application in `ui/`.
- Use `references/office-website` as a product, UX, information architecture, and interaction reference.
- For each reference subtree, prefer its `graphify-out/` summary when present, then confirm behavior in the source files.
- Preserve PolarBearOffice architecture and naming rather than cloning the reference app structure mechanically.
- Prefer local-first document handling via the local server wherever feasible.
- Keep transport and integration code isolated inside `ui/src/wasm/` and keep React components focused on presentation and editor workflow.

## Implementation Priorities

1. Open existing `.docx`, `.xlsx`, and `.pptx` files reliably.
2. Preserve document structure when reading into internal models.
3. Support editing against the internal Rust models.
4. Write modified documents back to valid OOXML packages.
5. Keep round-trip fidelity high: read -> edit -> write -> read should remain equivalent for supported features.

## Editing Priorities By App

### Writer (`.docx`)
- Prioritize paragraphs, runs, text formatting, headings, lists, and tables.
- Add images, sections, and richer layout features only after basic text editing is stable.

### Calc (`.xlsx`)
- Prioritize worksheets, cells, shared strings, numeric and text values, and basic formulas.
- Add formatting, merges, and richer workbook features after core grid editing is stable.

### Slides (`.pptx`)
- Prioritize slides, text boxes, basic shapes, positioning, and slide order.
- Add themes, images, and richer drawing behavior after base slide editing is stable.

## Engineering Standards

### Rust
- Use `PboError` and `PboResult<T>` from `core/common` for fallible operations.
- Avoid `unwrap()` and `expect()` in library code.
- Prefer borrowing over cloning.
- Keep parsing and serialization logic format-specific, but keep shared ZIP and XML logic in `core/ooxml`.
- Add tests for each public function and for each supported document round-trip path.

### TypeScript and React
- Use strict TypeScript.
- Keep components functional and hook-based.
- Do not let UI components parse OOXML directly.
- Route all document IO and transformation logic through the Rust layer, currently via `pbo-server`.

## Reference Usage Rules

- Use material under `references/` to understand workflows, feature coverage, and UI patterns.
- Re-implement required behavior in PolarBearOffice code instead of copying large blocks of reference code.
- Keep PolarBearOffice-specific abstractions, naming, and module boundaries clear.

## Definition Of Success

PolarBearOffice succeeds when a user can locally open, edit, and save `.docx`, `.xlsx`, and `.pptx` documents through the web app, with the editing engine implemented in Rust and integrated through a clean local UI/API boundary.