# PolarBearOffice Architecture

## Overview

PolarBearOffice is an open-source office suite inspired by Microsoft Office 365, built with **Rust**, a local **Axum** server, and a **React** UI. It supports `.docx`, `.pptx`, and `.xlsx` formats via the OOXML standard.

This document is the authority for core decisions, architecture boundaries, and reference interpretation.

## Layered Architecture

```
┌─────────────────────────────────────────────────┐
│                    ui/ (React)                  │
│         Writer  │  Slides  │  Calc              │
├─────────────────────────────────────────────────┤
│ server/ (Axum local HTTP API + session storage) │
├──────────┬──────────┬──────────┬────────────────┤
│ pbo-docx │ pbo-pptx │ pbo-xlsx │    core/wasm   │
│  (OOXML  │  (OOXML  │  (OOXML  │   (Axum API)   │
│  import/ │  import/ │  import/ │                │
│  export) │  export) │  export) │                │
├──────────┴──────────┴──────────┴────────────────┤
│                   pbo-ooxml                     │
│         OPC package handling (ZIP + XML)        │
├─────────────────────────────────────────────────┤
│                   pbo-common                    │
│     Internal document models + shared types     │
└─────────────────────────────────────────────────┘
```

## Key Design Decisions

### Source Of Truth
- The Rust internal document models are the source of truth for editing.
- OOXML is an import and export format, not the live editing representation.
- Architecture decisions should preserve this separation even when the reference implementation organizes behavior differently.

### Internal Document Models (pbo-common)
The OOXML crates are **import/export adapters**, not the live editing format. Each app has its own canonical model:
- `WriterDocument` — paragraphs, runs, tables
- `SlidesDocument` — slides, text boxes, shapes, images
- `CalcDocument` — sheets, cells, formulas

This separation keeps the editor independent from file-format quirks and enables undo/redo, collaboration, and cross-format conversion.

### OPC Package Layer (pbo-ooxml)
All OOXML formats are ZIP archives of XML parts. `pbo-ooxml` handles the shared ZIP/XML/relationship/content-type logic so format crates stay focused on their specific XML schemas.

### WASM Bindings (pbo-wasm)
`core/wasm` remains an optional integration target, but the current UI does not load it at runtime. The active UI path is `ui` -> `pbo-server` -> core crates.

If the WASM path is revived, the boundary should stay narrow:
- export task-oriented functions rather than rich object graphs
- prefer JSON-compatible payloads or byte buffers across the boundary
- keep UI code from depending directly on OOXML parsing logic

### Server (pbo-server)
Axum-based API for local-development support, file handling, and future local workflows.

It is not a public-cloud target at this stage. Do not introduce cloud-only assumptions, multi-tenant behavior, or deployment-driven architecture.

Current server responsibilities include:
- parsing and writing OOXML payloads for `.docx`, `.pptx`, and `.xlsx`
- returning render-ready PPTX documents through `POST /api/pptx/render`

The server remains a local orchestration layer. Rendering rules, document structure, and format behavior still belong in Rust crates under `core/`.

## Runtime Flows

### Slides Runtime Flow
The slides app now has two PPTX load paths:
- preferred path: `ui/src/apps/slides/Slides.tsx` -> `ui/src/wasm/loader.ts::renderPptx()` -> `POST /api/pptx/render` -> `pbo_pptx::read_pptx_rendered()`
- fallback path: `ui/src/apps/slides/Slides.tsx` -> `parsePptx()` -> `POST /api/pptx/parse` -> `pbo_pptx::read_pptx()`

This split is intentional:
- `read_pptx()` builds the canonical `SlidesDocument`
- `render_document()` materializes layout/master inheritance and placeholder geometry for display
- `read_pptx_rendered()` wraps both steps so the UI does not have to re-implement PresentationML inheritance rules

The UI should prefer the render-ready path for viewing and editing, while keeping the canonical Rust slide model as the source of truth.

## UI/Server Payload Examples

The UI/server boundary currently uses two transport shapes:
- raw OOXML bytes for `parse`, `render`, and `roundtrip` routes
- JSON for canonical document models returned by `parse` and `render`, and posted to `write`

### 1. PPTX Parse Request
UI call site:
- `ui/src/wasm/loader.ts::parsePptx(bytes)`

HTTP shape:

```http
POST /api/pptx/parse
Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation

<raw .pptx bytes>
```

Response shape:

```json
{
      "slides": [
            {
                  "elements": [
                        {
                              "TextBox": {
                                    "text": "Hello",
                                    "x": 914400,
                                    "y": 685800,
                                    "width": 3657600,
                                    "height": 914400,
                                    "font_size": 24.0,
                                    "bold": false,
                                    "italic": false,
                                    "color": "1F1F1F",
                                    "align": "Left",
                                    "rotation": 0,
                                    "placeholder": "title",
                                    "placeholder_idx": null,
                                    "paragraphs": [],
                                    "fill": null,
                                    "outline": null,
                                    "body": {
                                          "inset_left": null,
                                          "inset_top": null,
                                          "inset_right": null,
                                          "inset_bottom": null,
                                          "anchor": null,
                                          "rotation": null,
                                          "vertical": null,
                                          "wrap": null,
                                          "vertical_overflow": null,
                                          "horizontal_overflow": null
                                    },
                                    "flip_h": false,
                                    "flip_v": false
                              }
                        }
                  ],
                  "notes": null,
                  "layout_idx": 0,
                  "background": null
            }
      ],
      "slide_width": 12192000,
      "slide_height": 6858000,
      "layouts": [],
      "masters": [],
      "images": [],
      "metadata": null
}
```

This payload is a `SlidesDocument` JSON projection of the canonical Rust model.

### 2. PPTX Render Request
UI call site:
- `ui/src/wasm/loader.ts::renderPptx(bytes)`

HTTP shape:

```http
POST /api/pptx/render
Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation

<raw .pptx bytes>
```

Response shape:
- JSON `SlidesDocument`, same envelope as `parse`
- slide elements are already materialized for display by `pbo_pptx::read_pptx_rendered()`

Practical meaning for the UI:
- inherited placeholder geometry is already resolved
- layout/master chrome is already merged into slide-local render data
- React renders a display-ready document without re-implementing OOXML inheritance

### 3. PPTX Write Request
UI call site:
- `ui/src/wasm/loader.ts::createPptx(doc)`

HTTP shape:

```http
POST /api/pptx/write
Content-Type: application/json

{
      "slides": [
            {
                  "elements": [
                        {
                              "Shape": {
                                    "shape_type": "Rectangle",
                                    "preset": null,
                                    "adjustments": [],
                                    "x": 1143000,
                                    "y": 1905000,
                                    "width": 2095500,
                                    "height": 1047750,
                                    "rotation": 0,
                                    "fill": { "Solid": "2b5ce6" },
                                    "outline": null,
                                    "text": null,
                                    "custom_geometry": null,
                                    "flip_h": false,
                                    "flip_v": false,
                                    "effects": {
                                          "shadow": null,
                                          "glow": null,
                                          "soft_edge_radius": null
                                    }
                              }
                        }
                  ],
                  "notes": "Speaker notes",
                  "layout_idx": null,
                  "background": null
            }
      ],
      "slide_width": 12192000,
      "slide_height": 6858000,
      "layouts": [],
      "masters": [],
      "images": [],
      "metadata": null
}
```

Response shape:

```http
200 OK
Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation

<raw .pptx bytes>
```

### Contract Rules
- UI-to-server binary uploads must use the OOXML MIME type for the selected kind.
- Canonical document writes use JSON, not multipart form uploads.
- The UI should treat response JSON as the Rust-owned model contract defined in `pbo-common` and mirrored in `ui/src/wasm/types.ts`.
- Render-ready responses are for display/workflow convenience and must not become an independent client-side document model.

## Reference Interpretation

### Reference Role
- Material under `references/` is behavior, UX, and workflow reference material.
- PolarBearOffice should re-implement required behavior in Rust and TypeScript rather than extending the reference codebase directly.
- Preserve PolarBearOffice naming and module boundaries instead of mirroring the reference tree mechanically.

## Core Decisions

### Core Refactor Direction
- Refactor office behavior from `references/core` into native Rust crates under `core/`.
- Do not use the reference JavaScript core as the long-term engine.
- Keep format-specific parsing and serialization in `pbo-docx`, `pbo-pptx`, and `pbo-xlsx`.
- Keep shared ZIP and XML concerns in `pbo-ooxml`.

### UI Direction
- Build the product web app in `ui/`.
- Keep transport and integration concerns isolated in `ui/src/wasm/`.
- Keep React components focused on presentation and editor workflow, not OOXML parsing.
- Keep slide rendering heuristics in the product renderer, but keep OOXML parsing, inheritance resolution, and canonical document semantics in Rust.

### PPTX Direction
- `pbo-pptx` owns PPTX parse, write, and render-materialization behavior.
- The render-materialization layer is a projection for display, not a second document model.
- Master/layout inheritance, placeholder resolution, theme/color mapping, and related OOXML semantics should be resolved in Rust rather than duplicated in React.
- The slides UI may apply presentation-layer sizing and text-fit behavior, but should not become the source of OOXML truth.

### Delivery Scope
- Current success means reliable local open, edit, and save for `.docx`, `.xlsx`, and `.pptx`.
- Public cloud deployment is out of scope for current architecture decisions.
- Prefer local-first workflows wherever feasible.

## Crate Graph

```
pbo-common  ←  pbo-ooxml  ←  pbo-docx
                          ←  pbo-pptx
                          ←  pbo-xlsx
                                 ↑
                        pbo-server (native target used by UI; parse/render/write APIs)
                        pbo-wasm (optional alternate target)
```

## Directory Structure

```
polar-bear-office/
├── Cargo.toml          # Workspace root
├── core/
│   ├── common/         # Shared types & internal models
│   ├── ooxml/          # OPC package handling
│   ├── docx/           # DOCX import/export
│   ├── pptx/           # PPTX import/export
│   ├── xlsx/           # XLSX import/export
│   └── wasm/           # Optional WASM bindings
├── server/             # Axum API server
├── ui/                 # React + TypeScript + Vite
│   ├── src/apps/       # Writer, Slides, Calc views
│   └── src/wasm/       # Transport/loading boundary used by the UI
├── fixtures/           # Test files (OOXML samples)
└── docs/               # Documentation
```
