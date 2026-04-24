# PolarBearOffice UI

React + TypeScript + Vite frontend for PolarBearOffice. The current UI talks to
the Rust engine through `pbo-server` over HTTP.

## Prerequisites

- Node.js 18+
- Rust toolchain for running `pbo-server`

## Start the UI

From the `ui/` directory:

```sh
npm install
npm run dev
```

In a second terminal, start the local API server from the workspace root:

```sh
cargo run -p pbo-server
```

## Run the dev server

```sh
npm run dev
```

Open http://localhost:5173 and navigate to Writer, Calc, or Slides.

## Layout

- `src/apps/writer/` — DOCX app (paragraphs, runs, headings, alignment)
- `src/apps/calc/` — XLSX app (sheets, cells, shared strings)
- `src/apps/slides/` — PPTX app (text boxes, basic shapes)
- `src/wasm/` — transport client and typed façade for the local API.
- `src/components/` — Shared UI pieces (e.g. `FileToolbar`)
