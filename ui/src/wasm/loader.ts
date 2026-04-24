/**
 * HTTP client for the PolarBearOffice server (`pbo-server`).
 *
 * The UI talks to the Rust core via the server, so the browser does not need
 * to load any WASM. Vite proxies `/api` to `http://localhost:3001` in dev.
 *
 * If you ever want to swap to in-browser parsing, replace this module with
 * a wasm-bindgen wrapper exposing the same façade.
 */

import type { CalcDocument, SlidesDocument, WriterDocument } from "./types";

const API_BASE = (import.meta.env.VITE_PBO_API as string | undefined) ?? "/api";

async function postBytes(path: string, mime: string, bytes: Uint8Array): Promise<Response> {
  // Copy into a fresh ArrayBuffer so fetch typings accept the BodyInit
  // regardless of whether Uint8Array is backed by SharedArrayBuffer.
  const ab = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(ab).set(bytes);
  const resp = await fetch(`${API_BASE}${path}`, {
    method: "POST",
    headers: { "Content-Type": mime },
    body: ab,
  });
  if (!resp.ok) {
    throw new Error(await formatError(resp));
  }
  return resp;
}

async function postJson(path: string, body: unknown): Promise<Response> {
  const resp = await fetch(`${API_BASE}${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!resp.ok) {
    throw new Error(await formatError(resp));
  }
  return resp;
}

async function formatError(resp: Response): Promise<string> {
  const text = await resp.text();
  try {
    const j = JSON.parse(text) as { error?: string };
    if (j.error) return `${resp.status} ${resp.statusText}: ${j.error}`;
  } catch {
    /* not JSON */
  }
  return `${resp.status} ${resp.statusText}: ${text || "(empty body)"}`;
}

// ── Typed façade ────────────────────────────────────

export async function parseDocx(bytes: Uint8Array): Promise<WriterDocument> {
  const resp = await postBytes("/docx/parse", MIME.docx, bytes);
  return (await resp.json()) as WriterDocument;
}

export async function createDocx(doc: WriterDocument): Promise<Uint8Array> {
  const resp = await postJson("/docx/write", doc);
  return new Uint8Array(await resp.arrayBuffer());
}

export async function parseXlsx(bytes: Uint8Array): Promise<CalcDocument> {
  const resp = await postBytes("/xlsx/parse", MIME.xlsx, bytes);
  return (await resp.json()) as CalcDocument;
}

export async function createXlsx(doc: CalcDocument): Promise<Uint8Array> {
  const resp = await postJson("/xlsx/write", doc);
  return new Uint8Array(await resp.arrayBuffer());
}

export async function parsePptx(bytes: Uint8Array): Promise<SlidesDocument> {
  const resp = await postBytes("/pptx/parse", MIME.pptx, bytes);
  return (await resp.json()) as SlidesDocument;
}

export async function createPptx(doc: SlidesDocument): Promise<Uint8Array> {
  const resp = await postJson("/pptx/write", doc);
  return new Uint8Array(await resp.arrayBuffer());
}

// ── Fixtures (developer/demo aid) ───────────────────

export interface FixtureEntry {
  kind: "docx" | "xlsx" | "pptx";
  name: string;
  bytes: number;
}

export async function listFixtures(): Promise<FixtureEntry[]> {
  const resp = await fetch(`${API_BASE}/fixtures`);
  if (!resp.ok) throw new Error(await formatError(resp));
  const j = (await resp.json()) as { items: FixtureEntry[] };
  return j.items;
}

export async function fetchFixture(kind: "docx" | "xlsx" | "pptx", name: string): Promise<Uint8Array> {
  const resp = await fetch(`${API_BASE}/fixtures/${kind}/${encodeURIComponent(name)}`);
  if (!resp.ok) throw new Error(await formatError(resp));
  return new Uint8Array(await resp.arrayBuffer());
}

// ── File helpers ────────────────────────────────────

export async function readFileBytes(file: File): Promise<Uint8Array> {
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

export function downloadBytes(
  bytes: Uint8Array,
  filename: string,
  mime: string,
): void {
  const ab = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(ab).set(bytes);
  const blob = new Blob([ab], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export const MIME = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
} as const;
