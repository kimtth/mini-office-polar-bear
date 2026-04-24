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
│           server/ (Axum local HTTP API)         │
├──────────┬──────────┬──────────┬────────────────┤
│ pbo-docx │ pbo-pptx │ pbo-xlsx │    core/wasm    │
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

## Reference Interpretation

### Reference Role
- Material under `references/` is behavior, UX, and workflow reference material.
- PolarBearOffice should re-implement required behavior in Rust and TypeScript rather than extending the reference codebase directly.
- Preserve PolarBearOffice naming and module boundaries instead of mirroring the reference tree mechanically.

### Graphify Workflow
- For each reference subtree under `references/`, look for a local `graphify-out/` directory or graphify-generated summary artifacts.
- When present, use `graphify-out/` as the first-pass map of concepts, files, and workflows.
- Verify important behavior against the original reference sources before porting or redesigning it.
- If a `graphify-out/` directory is absent, inspect the reference sources directly and avoid inventing synthetic summaries.

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
                      pbo-server (native target used by UI)
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
│   └── src/apps/       # Writer, Slides, Calc views
├── fixtures/           # Test files (OOXML samples)
└── docs/               # Documentation
```
