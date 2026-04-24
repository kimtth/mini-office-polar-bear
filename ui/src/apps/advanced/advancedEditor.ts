import type { NavigateFunction } from "react-router-dom";
import { MIME } from "../../wasm/loader";

const API_BASE = (import.meta.env.VITE_PBO_API as string | undefined) ?? "/api";
const ONLYOFFICE_BASE = ((import.meta.env.VITE_ONLYOFFICE_BASE_URL as string | undefined) ?? "http://localhost:8080").replace(/\/$/, "");

export type AdvancedEditorKind = "docx" | "xlsx" | "pptx";

interface AdvancedSessionResponse {
  id: string;
  kind: AdvancedEditorKind;
  title: string;
  url: string;
}

interface LaunchAdvancedEditorOptions {
  bytes: Uint8Array;
  filename: string;
  kind: AdvancedEditorKind;
  navigate: NavigateFunction;
}

function ensureExtension(filename: string, kind: AdvancedEditorKind): string {
  const trimmed = filename.trim();
  if (!trimmed) return `untitled.${kind}`;
  return trimmed.toLowerCase().endsWith(`.${kind}`) ? trimmed : `${trimmed}.${kind}`;
}

function toBody(bytes: Uint8Array): ArrayBuffer {
  const ab = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(ab).set(bytes);
  return ab;
}

export function onlyOfficeDocsApiUrl(): string {
  return `${ONLYOFFICE_BASE}/web-apps/apps/api/documents/api.js`;
}

export function onlyOfficeBaseUrl(): string {
  return ONLYOFFICE_BASE;
}

export function documentTypeFor(kind: AdvancedEditorKind): "word" | "cell" | "slide" {
  if (kind === "docx") return "word";
  if (kind === "xlsx") return "cell";
  return "slide";
}

export async function createAdvancedSession(
  kind: AdvancedEditorKind,
  filename: string,
  bytes: Uint8Array,
): Promise<AdvancedSessionResponse> {
  const title = ensureExtension(filename, kind);
  const resp = await fetch(`${API_BASE}/advanced-session/${kind}?title=${encodeURIComponent(title)}`, {
    method: "POST",
    headers: {
      "Content-Type": MIME[kind],
    },
    body: toBody(bytes),
  });
  if (!resp.ok) {
    const message = await resp.text();
    throw new Error(message || "Failed to create advanced editor session.");
  }
  return (await resp.json()) as AdvancedSessionResponse;
}

export async function launchAdvancedEditor({
  bytes,
  filename,
  kind,
  navigate,
}: LaunchAdvancedEditorOptions): Promise<void> {
  const session = await createAdvancedSession(kind, filename, bytes);
  const params = new URLSearchParams({
    kind: session.kind,
    title: session.title,
    file: session.url,
  });
  navigate(`/advanced-editor?${params.toString()}`);
}