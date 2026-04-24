// Shared rendering helpers across Writer/Calc/Slides apps.
//
// Keeps CSS-mapping logic out of each app's JSX so changes to OOXML
// interpretation only touch one file.

import type {
  AbstractNum,
  BorderLine,
  CellBorder,
  CellBorders,
  ListFormat,
  ListRef,
  NumLevel,
  NumberingTable,
  ParagraphStyle,
  RunStyle,
  StyleCatalog,
  StyleDefinition,
} from "./types";
import type { CSSProperties } from "react";

// ── DOCX: named style inheritance ──────────────────

/** Walk the `based_on` chain to resolve a paragraph's effective style. */
export function resolveParagraphStyle(
  catalog: StyleCatalog,
  styleId: string | null | undefined,
): { paragraph: ParagraphStyle | null; run: RunStyle | null } {
  if (!styleId) return { paragraph: null, run: null };
  const chain: StyleDefinition[] = [];
  let id: string | null | undefined = styleId;
  const seen = new Set<string>();
  while (id && !seen.has(id)) {
    seen.add(id);
    const def = catalog.styles.find((s) => s.style_id === id);
    if (!def) break;
    chain.unshift(def);
    id = def.based_on;
  }
  let paragraph: ParagraphStyle | null = null;
  let run: RunStyle | null = null;
  for (const def of chain) {
    if (def.paragraph) paragraph = mergeParagraph(paragraph, def.paragraph);
    if (def.run) run = mergeRun(run, def.run);
  }
  return { paragraph, run };
}

function mergeParagraph(
  a: ParagraphStyle | null,
  b: ParagraphStyle,
): ParagraphStyle {
  if (!a) return { ...b };
  return {
    alignment: b.alignment ?? a.alignment,
    heading_level: b.heading_level ?? a.heading_level,
    style_id: b.style_id ?? a.style_id,
    indent: b.indent ?? a.indent,
    spacing: b.spacing ?? a.spacing,
    list: b.list ?? a.list,
  };
}

function mergeRun(a: RunStyle | null, b: RunStyle): RunStyle {
  if (!a) return { ...b };
  return {
    bold: b.bold || a.bold,
    italic: b.italic || a.italic,
    underline: b.underline || a.underline,
    strike: b.strike || a.strike,
    font_size: b.font_size ?? a.font_size,
    font_family: b.font_family ?? a.font_family,
    color: b.color ?? a.color,
    highlight: b.highlight ?? a.highlight,
    vertical_align: b.vertical_align ?? a.vertical_align,
  };
}

/** Merge two optional run styles; `override` wins. */
export function overlayRun(
  base: RunStyle | null,
  override: RunStyle | null,
): RunStyle | null {
  if (!base) return override;
  if (!override) return base;
  return mergeRun(base, override);
}

// ── DOCX: list marker resolution ───────────────────

/** Resolve a `{num_id, level}` ref to its bullet/number level definition. */
export function resolveListLevel(
  numbering: NumberingTable,
  ref: ListRef | null | undefined,
): NumLevel | null {
  if (!ref) return null;
  const inst = numbering.instances.find((i) => i.num_id === ref.num_id);
  if (!inst) return null;
  const abs: AbstractNum | undefined = numbering.abstract_nums.find(
    (a) => a.abstract_num_id === inst.abstract_num_id,
  );
  if (!abs) return null;
  return abs.levels.find((l) => l.ilvl === ref.level) ?? null;
}

/** Bullet glyph or numeric prefix to display before a list paragraph. */
export function listMarker(
  level: NumLevel,
  counter: number,
): string {
  const fmt: ListFormat = level.format;
  switch (fmt) {
    case "Bullet":
      return bulletChar(level.text);
    case "Decimal":
      return `${counter}.`;
    case "LowerLetter":
      return `${toLetters(counter, false)}.`;
    case "UpperLetter":
      return `${toLetters(counter, true)}.`;
    case "LowerRoman":
      return `${toRoman(counter).toLowerCase()}.`;
    case "UpperRoman":
      return `${toRoman(counter)}.`;
    case "Ordinal":
      return `${counter}`;
    default:
      return level.text || "•";
  }
}

function bulletChar(hint: string): string {
  if (!hint) return "•";
  // Common OOXML bullet glyphs map to nearer unicode equivalents.
  const first = hint.charAt(0);
  if (first === "o") return "○";
  if (first === "§") return "▪";
  if (first === "·" || first === "•") return "•";
  if (first === "-") return "–";
  return first;
}

function toLetters(n: number, upper: boolean): string {
  let s = "";
  let c = n;
  while (c > 0) {
    const rem = (c - 1) % 26;
    s = String.fromCharCode((upper ? 65 : 97) + rem) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s || (upper ? "A" : "a");
}

function toRoman(n: number): string {
  if (n <= 0) return "";
  const map: Array<[number, string]> = [
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"],
    [100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
    [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
  ];
  let s = "";
  let v = n;
  for (const [k, sym] of map) {
    while (v >= k) { s += sym; v -= k; }
  }
  return s;
}

// ── DOCX: CSS mappers ─────────────────────────────

/** Map a OOXML `w:tblBorders`/`w:tcBorders` value to a CSS border string. */
export function borderLineToCss(b: BorderLine | null | undefined): string | undefined {
  if (!b) return undefined;
  const width = Math.max(0.5, (b.size ?? 4) / 8); // OOXML size = eighths of a point
  const style = borderStyleToCss(b.style);
  const color = b.color ? `#${b.color}` : "#888";
  return `${width}pt ${style} ${color}`;
}

function borderStyleToCss(style: string): string {
  switch ((style || "").toLowerCase()) {
    case "none":
    case "nil":
      return "none";
    case "dashed":
    case "dashsmallgap":
      return "dashed";
    case "dotted":
      return "dotted";
    case "double":
      return "double";
    case "thick":
      return "solid";
    default:
      return "solid";
  }
}

export function cellBordersToCss(
  b: CellBorders | null | undefined,
): CSSProperties {
  if (!b) return {};
  return {
    borderTop: borderLineToCss(b.top),
    borderBottom: borderLineToCss(b.bottom),
    borderLeft: borderLineToCss(b.left),
    borderRight: borderLineToCss(b.right),
  };
}

// ── XLSX: CSS mappers ─────────────────────────────

export function xlsxBorderToCss(
  b: CellBorder | null | undefined,
): CSSProperties {
  if (!b) return {};
  const edge = (e: { style: string; color: string | null } | null) => {
    if (!e || !e.style || e.style === "none") return undefined;
    const color = e.color ? `#${e.color}` : "#000";
    const width = e.style === "thick" || e.style === "double" ? "1.5pt" : "1pt";
    const style = ["dashed", "dotted", "double"].includes(e.style)
      ? e.style
      : "solid";
    return `${width} ${style} ${color}`;
  };
  return {
    borderTop: edge(b.top),
    borderBottom: edge(b.bottom),
    borderLeft: edge(b.left),
    borderRight: edge(b.right),
  };
}

// ── Image resolution (base-64 data URI) ────────────

export function imageDataUri(
  mime: string,
  data_base64: string,
): string {
  return `data:${mime};base64,${data_base64}`;
}
