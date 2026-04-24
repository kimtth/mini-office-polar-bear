import React, {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import { useNavigate } from "react-router-dom";
import MenuBar, { type Menu } from "../../components/MenuBar";
import FixturePicker from "../../components/FixturePicker";
import { useHiddenFilePicker } from "../../components/FilePicker";
import { launchAdvancedEditor } from "../advanced/advancedEditor";
import {
  createXlsx,
  downloadBytes,
  fetchFixture,
  MIME,
  parseXlsx,
  readFileBytes,
} from "../../wasm/loader";
import type {
  CalcDocument,
  Cell,
  CellAlignment,
  CellFill,
  CellFont,
  CellValue,
  CellXf,
  HAlign,
  Sheet,
  SheetHyperlink,
  VAlign,
} from "../../wasm/types";
import { formatCellValue } from "../../wasm/numFormat";
import { xlsxBorderToCss } from "../../wasm/renderHelpers";
import "./calc.css";

// ── Built-in Excel number-format table (ECMA-376 §18.8.30) ─────────────────────
const BUILTIN_FMTS: Record<number, string> = {
  0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00",
  9: "0%", 10: "0.00%", 11: "0.00E+00",
  12: "# ?/?", 13: "# ??/??",
  14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
  18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
  22: "m/d/yy h:mm",
  37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)",
  45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
};

// ── Helpers ───────────────────────────────────────────

function emptySheet(name: string): Sheet {
  return {
    name,
    cells: {},
    columns: [],
    row_heights: {},
    column_widths: {},
    merges: [],
    hyperlinks: [],
    view: null,
    tab_color: null,
  };
}
function emptyDoc(): CalcDocument {
  return {
    metadata: {
      title: null,
      author: null,
      last_modified_by: null,
      created: null,
      modified: null,
    },
    sheets: [emptySheet("Sheet1")],
    styles: {
      num_fmts: [],
      fonts: [],
      fills: [],
      borders: [],
      cell_xfs: [],
    },
    defined_names: [],
  };
}

interface EffectiveStyle {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  font_size: number | null;
  font_family: string | null;
  text_color: string | null;
  background_color: string | null;
  number_format: string | null;
  alignment: CellAlignment | null;
  borderCss: React.CSSProperties;
}

function numberFormatFor(doc: CalcDocument, xf: CellXf): string | null {
  if (xf.num_fmt_id == null) return null;
  const id = xf.num_fmt_id;
  const builtin = BUILTIN_FMTS[id];
  if (builtin) return builtin;
  const user = doc.styles.num_fmts.find((n) => n.id === id);
  return user?.format_code ?? null;
}

function resolveStyle(doc: CalcDocument, cell: Cell | undefined): EffectiveStyle {
  const out: EffectiveStyle = {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    font_size: null,
    font_family: null,
    text_color: null,
    background_color: null,
    number_format: null,
    alignment: null,
    borderCss: {},
  };
  if (!cell || cell.style_id == null) return out;
  const xf = doc.styles.cell_xfs[cell.style_id];
  if (!xf) return out;
  if (xf.font_id != null) {
    const f = doc.styles.fonts[xf.font_id];
    if (f) {
      out.bold = f.bold;
      out.italic = f.italic;
      out.underline = f.underline;
      out.strike = f.strike;
      out.font_size = f.size;
      out.font_family = f.name;
      out.text_color = f.color;
    }
  }
  if (xf.fill_id != null) {
    const fl = doc.styles.fills[xf.fill_id];
    if (fl && fl.pattern_type && fl.pattern_type !== "none") {
      out.background_color = fl.fg_color ?? fl.bg_color;
    }
  }
  if (xf.border_id != null) {
    const b = doc.styles.borders[xf.border_id];
    if (b) out.borderCss = xlsxBorderToCss(b);
  }
  if (xf.alignment) out.alignment = xf.alignment;
  out.number_format = numberFormatFor(doc, xf);
  return out;
}

function cellValueToJs(v: CellValue): unknown {
  if (v === "Empty") return null;
  if ("Text" in v) return v.Text;
  if ("Number" in v) return v.Number;
  if ("Boolean" in v) return v.Boolean;
  if ("Date" in v) return v.Date;
  if ("Error" in v) return v.Error;
  return null;
}

function formattedCellDisplay(cell: Cell | undefined, fmt: string | null): string {
  if (!cell) return "";
  const raw = cellValueToJs(cell.value);
  if (raw == null) return "";
  if (cell.formula && (cell.value === "Empty")) return ""; // unevaluated
  const r = formatCellValue(raw, fmt);
  return r.text;
}

function hAlignCss(a: HAlign | null | undefined): React.CSSProperties["textAlign"] {
  switch (a) {
    case "Center": return "center";
    case "Right": return "right";
    case "Justify": return "justify";
    case "Left": return "left";
    default: return undefined;
  }
}

function vAlignCss(a: VAlign | null | undefined): React.CSSProperties["alignItems"] {
  switch (a) {
    case "Top": return "flex-start";
    case "Bottom": return "flex-end";
    case "Middle": return "center";
    default: return undefined;
  }
}

// Default alignment when no explicit alignment present:
// numbers right, booleans/errors center, text left.
function implicitHAlign(cell: Cell | undefined): React.CSSProperties["textAlign"] {
  if (!cell) return undefined;
  const v = cell.value;
  if (v === "Empty") return undefined;
  if ("Number" in v) return "right";
  if ("Boolean" in v || "Error" in v) return "center";
  return undefined;
}

/**
 * Allocate a new style (font/fill/xf) that inherits from an existing style_id
 * and applies the patch. Returns a new doc + the new style index.
 */
function allocateStyle(
  doc: CalcDocument,
  baseStyleId: number | null,
  patch: Partial<EffectiveStyle>,
): { doc: CalcDocument; style_id: number } {
  const baseXf: CellXf =
    (baseStyleId != null ? doc.styles.cell_xfs[baseStyleId] : undefined) ?? {
      num_fmt_id: null,
      font_id: null,
      fill_id: null,
      border_id: null,
      alignment: null,
      apply_number_format: false,
      apply_font: false,
      apply_fill: false,
      apply_border: false,
      apply_alignment: false,
    };
  const baseFont: CellFont =
    (baseXf.font_id != null ? doc.styles.fonts[baseXf.font_id] : undefined) ?? {
      size: null,
      name: null,
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      color: null,
    };
  const baseFill: CellFill =
    (baseXf.fill_id != null ? doc.styles.fills[baseXf.fill_id] : undefined) ?? {
      pattern_type: "none",
      fg_color: null,
      bg_color: null,
    };

  const newFont: CellFont = {
    ...baseFont,
    bold: patch.bold ?? baseFont.bold,
    italic: patch.italic ?? baseFont.italic,
    size: patch.font_size !== undefined ? patch.font_size : baseFont.size,
    color: patch.text_color !== undefined ? patch.text_color : baseFont.color,
  };
  const newFill: CellFill =
    patch.background_color !== undefined
      ? {
          pattern_type: patch.background_color ? "solid" : "none",
          fg_color: patch.background_color,
          bg_color: patch.background_color,
        }
      : baseFill;

  const fonts = [...doc.styles.fonts, newFont];
  const fills = [...doc.styles.fills, newFill];
  const newXf: CellXf = {
    ...baseXf,
    font_id: fonts.length - 1,
    fill_id: fills.length - 1,
    apply_font: true,
    apply_fill: patch.background_color !== undefined || baseXf.apply_fill,
  };
  const cell_xfs = [...doc.styles.cell_xfs, newXf];
  const newStyles = { ...doc.styles, fonts, fills, cell_xfs };
  return { doc: { ...doc, styles: newStyles }, style_id: cell_xfs.length - 1 };
}

function cellDisplay(v: CellValue): string {
  if (v === "Empty") return "";
  if ("Text" in v) return v.Text;
  if ("Number" in v) return String(v.Number);
  if ("Boolean" in v) return v.Boolean ? "TRUE" : "FALSE";
  if ("Date" in v) return v.Date;
  if ("Error" in v) return v.Error;
  return "";
}

function cellRaw(cell: Cell | undefined): string {
  if (!cell) return "";
  if (cell.formula) return `=${cell.formula}`;
  return cellDisplay(cell.value);
}

// ── CalcGrid: virtualized grid with merges/widths/heights/hyperlinks ──
interface CalcGridProps {
  sheet: Sheet;
  doc: CalcDocument;
  maxRow: number;
  maxCol: number;
  sel: { row: number; col: number };
  editing: { row: number; col: number; text: string } | null;
  onSelect: (row: number, col: number) => void;
  onBeginEdit: (row: number, col: number, initial?: string) => void;
  onChangeEditing: (text: string) => void;
  onCommitEdit: (advance: "row" | "col" | "none") => void;
  onCancelEdit: () => void;
}

function CalcGrid(props: CalcGridProps): React.ReactElement {
  const { sheet, doc, maxRow, maxCol, sel, editing } = props;
  const scrollRef = React.useRef<HTMLDivElement>(null);
  const [scrollTop, setScrollTop] = React.useState(0);
  const [viewportH, setViewportH] = React.useState(480);

  React.useEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    setViewportH(el.clientHeight);
    const onResize = () => setViewportH(el.clientHeight);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  const rowHeights = React.useMemo(() => {
    const arr: number[] = new Array(maxRow + 1).fill(22);
    for (let r = 1; r <= maxRow; r++) {
      arr[r] = resolveRowHeight(sheet, r) ?? 22;
    }
    return arr;
  }, [sheet, maxRow]);

  // Build prefix-sum of row top positions for virtualization.
  const rowTops = React.useMemo(() => {
    const tops: number[] = new Array(maxRow + 2).fill(0);
    for (let r = 1; r <= maxRow; r++) tops[r + 1] = tops[r] + rowHeights[r];
    return tops;
  }, [rowHeights, maxRow]);

  const totalHeight = rowTops[maxRow + 1] ?? 0;

  // Window of rows to render (vertical virtualization only).
  const OVERSCAN = 12;
  const firstVisible = React.useMemo(() => {
    let lo = 1, hi = maxRow;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (rowTops[mid + 1] <= scrollTop) lo = mid + 1;
      else hi = mid;
    }
    return Math.max(1, lo - OVERSCAN);
  }, [scrollTop, rowTops, maxRow]);

  const lastVisible = React.useMemo(() => {
    let r = firstVisible;
    while (r <= maxRow && rowTops[r] < scrollTop + viewportH) r++;
    return Math.min(maxRow, r + OVERSCAN);
  }, [firstVisible, rowTops, scrollTop, viewportH, maxRow]);

  const { anchor, covered } = React.useMemo(
    () => buildMergeMap(sheet),
    [sheet],
  );

  const colWidths = React.useMemo(() => {
    const arr: number[] = new Array(maxCol + 1).fill(96);
    for (let c = 1; c <= maxCol; c++) {
      arr[c] = resolveColumnWidth(sheet, c) ?? 96;
    }
    return arr;
  }, [sheet, maxCol]);

  return (
    <div
      className="pbo-calc-grid"
      ref={scrollRef}
      onScroll={(e) => setScrollTop((e.target as HTMLDivElement).scrollTop)}
    >
      <table style={{ height: totalHeight }}>
        <colgroup>
          <col style={{ width: 40 }} />
          {colWidths.slice(1, maxCol + 1).map((w, i) => (
            <col key={i} style={{ width: w }} />
          ))}
        </colgroup>
        <thead>
          <tr>
            <th className="pbo-corner" />
            {Array.from({ length: maxCol }, (_, i) => (
              <th key={i}>{columnLetters(i + 1)}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {firstVisible > 1 && (
            <tr style={{ height: rowTops[firstVisible] }}>
              <th />
              {Array.from({ length: maxCol }, (_, i) => (
                <td key={i} />
              ))}
            </tr>
          )}
          {Array.from({ length: lastVisible - firstVisible + 1 }, (_, idx) => {
            const row = firstVisible + idx;
            return (
              <tr key={row} style={{ height: rowHeights[row] }}>
                <th>{row}</th>
                {Array.from({ length: maxCol }, (_, c) => {
                  const col = c + 1;
                  const key = `${row}:${col}`;
                  if (covered.has(key)) return null;
                  const span = anchor.get(key);
                  const cell = sheet.cells[String(row)]?.[String(col)];
                  const selected = sel.row === row && sel.col === col;
                  const isEditing =
                    editing && editing.row === row && editing.col === col;
                  const style = resolveStyle(doc, cell);
                  const hl = hyperlinkFor(sheet, row, col);
                  const formatted = formattedCellDisplay(cell, style.number_format);
                  const cssBase = effectiveStyleToCSS(style, cell);
                  const vAlign = vAlignCss(style.alignment?.vertical ?? null);
                  return (
                    <td
                      key={col}
                      className={selected ? "pbo-selected" : undefined}
                      colSpan={span?.colSpan}
                      rowSpan={span?.rowSpan}
                      style={{
                        background: cssBase.background,
                        borderTop: cssBase.borderTop,
                        borderBottom: cssBase.borderBottom,
                        borderLeft: cssBase.borderLeft,
                        borderRight: cssBase.borderRight,
                      }}
                      onClick={() => props.onSelect(row, col)}
                      onDoubleClick={() => props.onBeginEdit(row, col)}
                    >
                      {isEditing ? (
                        <input
                          autoFocus
                          value={editing.text}
                          onChange={(e) => props.onChangeEditing(e.target.value)}
                          onBlur={() => props.onCommitEdit("none")}
                          onKeyDown={(e) => {
                            if (e.key === "Enter") {
                              e.preventDefault();
                              props.onCommitEdit("row");
                            } else if (e.key === "Tab") {
                              e.preventDefault();
                              props.onCommitEdit("col");
                            } else if (e.key === "Escape") {
                              props.onCancelEdit();
                            }
                          }}
                        />
                      ) : (
                        <div
                          className="pbo-cell-view"
                          style={{
                            ...cssBase,
                            background: undefined,
                            borderTop: undefined,
                            borderBottom: undefined,
                            borderLeft: undefined,
                            borderRight: undefined,
                            display: "flex",
                            alignItems: vAlign ?? "center",
                            justifyContent:
                              (cssBase.textAlign as string) === "right"
                                ? "flex-end"
                                : (cssBase.textAlign as string) === "center"
                                  ? "center"
                                  : "flex-start",
                          }}
                        >
                          {hl ? (
                            <a
                              href={hl.target}
                              target="_blank"
                              rel="noreferrer noopener"
                              onClick={(e) => e.stopPropagation()}
                              style={{ color: "inherit", textDecoration: "underline" }}
                            >
                              {hl.display ?? formatted}
                            </a>
                          ) : (
                            formatted
                          )}
                        </div>
                      )}
                    </td>
                  );
                })}
              </tr>
            );
          })}
          {lastVisible < maxRow && (
            <tr
              style={{ height: (rowTops[maxRow + 1] ?? 0) - (rowTops[lastVisible + 1] ?? 0) }}
            >
              <th />
              {Array.from({ length: maxCol }, (_, i) => (
                <td key={i} />
              ))}
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

function parseCellInput(text: string): {
  value: CellValue;
  formula: string | null;
} {
  if (text === "") return { value: "Empty", formula: null };
  if (text.startsWith("=")) {
    return { value: "Empty", formula: text.slice(1) };
  }
  const trimmed = text.trim();
  if (trimmed === "TRUE") return { value: { Boolean: true }, formula: null };
  if (trimmed === "FALSE") return { value: { Boolean: false }, formula: null };
  const num = Number(trimmed);
  if (trimmed !== "" && !Number.isNaN(num)) {
    return { value: { Number: num }, formula: null };
  }
  return { value: { Text: text }, formula: null };
}

function columnLetters(col: number): string {
  let s = "";
  let c = col;
  while (c > 0) {
    const rem = (c - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

function boundsOf(sheet: Sheet, minRows: number, minCols: number) {
  let maxRow = minRows;
  let maxCol = minCols;
  for (const [r, row] of Object.entries(sheet.cells)) {
    const rn = Number(r);
    if (rn > maxRow) maxRow = rn;
    for (const c of Object.keys(row)) {
      const cn = Number(c);
      if (cn > maxCol) maxCol = cn;
    }
  }
  return { maxRow, maxCol };
}

function getCell(sheet: Sheet, row: number, col: number): Cell | undefined {
  return sheet.cells[String(row)]?.[String(col)];
}

function writeCell(sheet: Sheet, row: number, col: number, text: string): Sheet {
  const cells = { ...sheet.cells };
  const rk = String(row);
  const rowMap = { ...(cells[rk] ?? {}) };
  if (text === "") {
    delete rowMap[String(col)];
  } else {
    const parsed = parseCellInput(text);
    const existing = rowMap[String(col)];
    rowMap[String(col)] = {
      value: parsed.value,
      formula: parsed.formula,
      style_id: existing?.style_id ?? null,
    };
  }
  if (Object.keys(rowMap).length === 0) delete cells[rk];
  else cells[rk] = rowMap;
  return { ...sheet, cells };
}

function setCellStyleId(sheet: Sheet, row: number, col: number, style_id: number): Sheet {
  const cells = { ...sheet.cells };
  const rk = String(row);
  const rowMap = { ...(cells[rk] ?? {}) };
  const existing = rowMap[String(col)];
  rowMap[String(col)] = {
    value: existing?.value ?? "Empty",
    formula: existing?.formula ?? null,
    style_id,
  };
  cells[rk] = rowMap;
  return { ...sheet, cells };
}

function effectiveStyleToCSS(s: EffectiveStyle, cell: Cell | undefined): React.CSSProperties {
  return {
    fontWeight: s.bold ? 700 : undefined,
    fontStyle: s.italic ? "italic" : undefined,
    textDecoration: [
      s.underline ? "underline" : null,
      s.strike ? "line-through" : null,
    ].filter(Boolean).join(" ") || undefined,
    fontSize: s.font_size ? `${s.font_size}pt` : undefined,
    fontFamily: s.font_family ?? undefined,
    background: s.background_color ? `#${s.background_color}` : undefined,
    color: s.text_color ? `#${s.text_color}` : undefined,
    textAlign:
      hAlignCss(s.alignment?.horizontal ?? null) ?? implicitHAlign(cell),
    whiteSpace: s.alignment?.wrap_text ? "normal" : "nowrap",
    ...s.borderCss,
  };
}

// ── Merge handling ────────────────────────────────
function buildMergeMap(sheet: Sheet): {
  anchor: Map<string, { rowSpan: number; colSpan: number }>;
  covered: Set<string>;
} {
  const anchor = new Map<string, { rowSpan: number; colSpan: number }>();
  const covered = new Set<string>();
  for (const m of sheet.merges) {
    const rows = m.row_end - m.row_start + 1;
    const cols = m.col_end - m.col_start + 1;
    anchor.set(`${m.row_start}:${m.col_start}`, { rowSpan: rows, colSpan: cols });
    for (let r = m.row_start; r <= m.row_end; r++) {
      for (let c = m.col_start; c <= m.col_end; c++) {
        if (r === m.row_start && c === m.col_start) continue;
        covered.add(`${r}:${c}`);
      }
    }
  }
  return { anchor, covered };
}

// ── Column width → px (1 char ≈ 7px, OOXML default char width ≈ 8.43) ──────────
function resolveColumnWidth(sheet: Sheet, col: number): number | undefined {
  for (const c of sheet.columns) {
    if (col >= c.min && col <= c.max) {
      if (c.hidden) return 0;
      if (c.width > 0) return Math.round(c.width * 7);
    }
  }
  const legacy = sheet.column_widths[String(col)];
  if (legacy) return Math.round(legacy * 7);
  return undefined;
}

function resolveRowHeight(sheet: Sheet, row: number): number | undefined {
  const h = sheet.row_heights[String(row)];
  return h ? Math.round(h * 1.333) : undefined; // 1 pt ≈ 1.333 px
}

function hyperlinkFor(
  sheet: Sheet,
  row: number,
  col: number,
): SheetHyperlink | undefined {
  const ref = `${columnLetters(col)}${row}`;
  for (const h of sheet.hyperlinks) {
    if (h.cell_ref === ref) return h;
    // Range form: A1:B3
    const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(h.cell_ref);
    if (m) {
      const c1 = parseColLetters(m[1]);
      const r1 = Number(m[2]);
      const c2 = parseColLetters(m[3]);
      const r2 = Number(m[4]);
      if (row >= r1 && row <= r2 && col >= c1 && col <= c2) return h;
    }
  }
  return undefined;
}

function parseColLetters(s: string): number {
  let n = 0;
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

// ── Component ─────────────────────────────────────────

export default function Calc() {
  const nav = useNavigate();
  const [doc, setDoc] = useState<CalcDocument>(emptyDoc());
  const [filename, setFilename] = useState<string>("Untitled.xlsx");
  const [activeIdx, setActiveIdx] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [extraRows, setExtraRows] = useState(0);
  const [extraCols, setExtraCols] = useState(0);
  const [sel, setSel] = useState<{ row: number; col: number }>({
    row: 1,
    col: 1,
  });
  const [editing, setEditing] = useState<{
    row: number;
    col: number;
    text: string;
  } | null>(null);
  const formulaRef = useRef<HTMLInputElement>(null);

  const sheet = doc.sheets[activeIdx] ?? doc.sheets[0];
  const { maxRow, maxCol } = useMemo(
    () => boundsOf(sheet, 20 + extraRows, 8 + extraCols),
    [sheet, extraRows, extraCols],
  );

  const setSheet = useCallback(
    (next: Sheet) => {
      setDoc((d) => ({
        ...d,
        sheets: d.sheets.map((s, i) => (i === activeIdx ? next : s)),
      }));
    },
    [activeIdx],
  );

  // ── File actions ───────────────────────────────
  const openFromFile = useCallback(async (file: File) => {
    setError(null);
    setBusy(true);
    try {
      const bytes = await readFileBytes(file);
      const parsed = await parseXlsx(bytes);
      setDoc(parsed.sheets.length ? parsed : emptyDoc());
      setActiveIdx(0);
      setFilename(file.name);
      setSel({ row: 1, col: 1 });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }, []);

  const { el: hiddenInput, open: openPicker } = useHiddenFilePicker(
    ".xlsx",
    openFromFile,
  );

  const handleSave = useCallback(async () => {
    setError(null);
    setBusy(true);
    try {
      const bytes = await createXlsx(doc);
      downloadBytes(bytes, filename, MIME.xlsx);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }, [doc, filename]);

  const handleNew = useCallback(() => {
    setDoc(emptyDoc());
    setFilename("Untitled.xlsx");
    setActiveIdx(0);
    setExtraRows(0);
    setExtraCols(0);
    setSel({ row: 1, col: 1 });
  }, []);

  // ── Edit/commit ────────────────────────────────
  const beginEdit = (row: number, col: number, initial?: string) => {
    const text = initial ?? cellRaw(getCell(sheet, row, col));
    setEditing({ row, col, text });
  };

  const commitEdit = useCallback(
    (text?: string) => {
      if (!editing) return;
      setSheet(writeCell(sheet, editing.row, editing.col, text ?? editing.text));
      setEditing(null);
    },
    [editing, sheet, setSheet],
  );

  const cancelEdit = () => setEditing(null);

  // ── Formatting ─────────────────────────────────
  const patchSelectionStyle = (patch: Partial<EffectiveStyle>) => {
    const current = getCell(sheet, sel.row, sel.col);
    const { doc: nextDoc, style_id } = allocateStyle(
      doc,
      current?.style_id ?? null,
      patch,
    );
    const nextSheet = setCellStyleId(
      nextDoc.sheets[activeIdx],
      sel.row,
      sel.col,
      style_id,
    );
    setDoc({
      ...nextDoc,
      sheets: nextDoc.sheets.map((s, i) => (i === activeIdx ? nextSheet : s)),
    });
  };
  const toggleBold = () => {
    const eff = resolveStyle(doc, getCell(sheet, sel.row, sel.col));
    patchSelectionStyle({ bold: !eff.bold });
  };
  const toggleItalic = () => {
    const eff = resolveStyle(doc, getCell(sheet, sel.row, sel.col));
    patchSelectionStyle({ italic: !eff.italic });
  };
  const setBg = (hex: string) =>
    patchSelectionStyle({ background_color: hex.replace("#", "") });
  const setFg = (hex: string) =>
    patchSelectionStyle({ text_color: hex.replace("#", "") });

  // ── Sheet management ───────────────────────────
  const addSheet = () => {
    const names = new Set(doc.sheets.map((s) => s.name));
    let n = doc.sheets.length + 1;
    let name = `Sheet${n}`;
    while (names.has(name)) {
      n += 1;
      name = `Sheet${n}`;
    }
    setDoc({ ...doc, sheets: [...doc.sheets, emptySheet(name)] });
    setActiveIdx(doc.sheets.length);
  };
  const renameSheet = () => {
    const name = prompt("Sheet name:", sheet.name);
    if (!name) return;
    setSheet({ ...sheet, name });
  };
  const deleteSheet = () => {
    if (doc.sheets.length <= 1) {
      alert("A workbook must contain at least one sheet.");
      return;
    }
    if (!confirm(`Delete sheet "${sheet.name}"?`)) return;
    const sheets = doc.sheets.filter((_, i) => i !== activeIdx);
    setDoc({ ...doc, sheets });
    setActiveIdx(Math.max(0, activeIdx - 1));
  };

  const addRow = () => setExtraRows((n) => n + 1);
  const addCol = () => setExtraCols((n) => n + 1);

  // ── Keyboard navigation ────────────────────────
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
        e.preventDefault();
        handleSave();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "o") {
        e.preventDefault();
        openPicker();
        return;
      }
      if (editing) return;
      const target = e.target as HTMLElement | null;
      if (target && target.tagName === "INPUT") return;
      let handled = false;
      const move = (dr: number, dc: number) => {
        setSel((s) => ({
          row: Math.max(1, Math.min(maxRow, s.row + dr)),
          col: Math.max(1, Math.min(maxCol, s.col + dc)),
        }));
        handled = true;
      };
      if (e.key === "ArrowUp") move(-1, 0);
      else if (e.key === "ArrowDown" || e.key === "Enter") move(1, 0);
      else if (e.key === "ArrowLeft") move(0, -1);
      else if (e.key === "ArrowRight" || e.key === "Tab") move(0, 1);
      else if (e.key === "Delete" || e.key === "Backspace") {
        setSheet(writeCell(sheet, sel.row, sel.col, ""));
        handled = true;
      } else if (e.key === "F2") {
        beginEdit(sel.row, sel.col);
        handled = true;
      } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        beginEdit(sel.row, sel.col, e.key);
        handled = true;
      }
      if (handled) e.preventDefault();
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sel, maxRow, maxCol, editing, sheet, handleSave, openPicker]);

  // ── Menus ──────────────────────────────────────
  const menus = useMemo<Menu[]>(
    () => [
      {
        label: "File",
        actions: [
          { label: "New", onClick: handleNew, kbd: "Ctrl+N" },
          { label: "Open…", onClick: openPicker, kbd: "Ctrl+O" },
          { label: "Open sample…", onClick: () => setPickerOpen(true) },
          { separator: true, label: "" },
          { label: "Save (.xlsx)", onClick: handleSave, kbd: "Ctrl+S" },
          { separator: true, label: "" },
          { label: "Back to home", onClick: () => nav("/") },
        ],
      },
      {
        label: "Sheet",
        actions: [
          { label: "New sheet", onClick: addSheet },
          { label: "Rename current sheet…", onClick: renameSheet },
          { label: "Delete current sheet", onClick: deleteSheet },
          { separator: true, label: "" },
          { label: "Insert row below", onClick: addRow },
          { label: "Insert column right", onClick: addCol },
        ],
      },
      {
        label: "Format",
        actions: [
          { label: "Bold", onClick: toggleBold, kbd: "Ctrl+B" },
          { label: "Italic", onClick: toggleItalic, kbd: "Ctrl+I" },
          {
            label: "Clear formatting",
            onClick: () => {
              const nextSheet = setCellStyleId(sheet, sel.row, sel.col, -1);
              // Setting to -1 is out of range; instead clear by setting to null via a fresh write.
              const cells = { ...nextSheet.cells };
              const rk = String(sel.row);
              const rowMap = { ...(cells[rk] ?? {}) };
              const existing = rowMap[String(sel.col)];
              if (existing) {
                rowMap[String(sel.col)] = {
                  value: existing.value,
                  formula: existing.formula,
                  style_id: null,
                };
                cells[rk] = rowMap;
                setSheet({ ...sheet, cells });
              }
            },
          },
        ],
      },
    ],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [handleNew, openPicker, handleSave, nav, doc, sheet, sel],
  );

  const selectedCell = getCell(sheet, sel.row, sel.col);
  const selectedRef = `${columnLetters(sel.col)}${sel.row}`;

  const openAdvancedEditor = useCallback(async () => {
    setBusy(true);
    try {
      const bytes = await createXlsx(doc);
      await launchAdvancedEditor({
        bytes,
        filename,
        kind: "xlsx",
        navigate: nav,
      });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }, [doc, filename, nav]);

  return (
    <div className="pbo-app">
      {hiddenInput}
      <FixturePicker
        kind="xlsx"
        open={pickerOpen}
        onClose={() => setPickerOpen(false)}
        onPick={async (f) => {
          setPickerOpen(false);
          setBusy(true);
          try {
            const bytes = await fetchFixture(f.kind, f.name);
            const parsed = await parseXlsx(bytes);
            setDoc(parsed.sheets.length ? parsed : emptyDoc());
            setActiveIdx(0);
            setFilename(f.name);
            setSel({ row: 1, col: 1 });
          } catch (e) {
            setError((e as Error).message);
          } finally {
            setBusy(false);
          }
        }}
      />

      <MenuBar
        menus={menus}
        appName="Calc"
        accent="var(--pbo-accent-calc)"
        rightSlot={
          <span
            style={{
              color: "var(--pbo-text-muted)",
              fontSize: "0.8rem",
              padding: "0 0.75rem",
            }}
          >
            {filename}
            {busy && " · working…"}
          </span>
        }
      />

      <div className="pbo-ribbon">
        <div className="pbo-group">
          <button
            className="pbo-btn"
            onClick={openPicker}
            title="Open (Ctrl+O)"
          >
            📂 Open
          </button>
          <button
            className="pbo-btn pbo-primary"
            onClick={handleSave}
            disabled={busy}
            title="Save (Ctrl+S)"
          >
            💾 Save
          </button>
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={toggleBold} title="Bold">
            <b>B</b>
          </button>
          <button className="pbo-btn" onClick={toggleItalic} title="Italic">
            <i>I</i>
          </button>
          <input
            type="color"
            className="pbo-color"
            defaultValue="#1d1d1d"
            onChange={(e) => setFg(e.target.value)}
            title="Text color"
          />
          <input
            type="color"
            className="pbo-color"
            defaultValue="#ffff00"
            onChange={(e) => setBg(e.target.value)}
            title="Fill color"
          />
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={addRow} title="Add row">
            ＋ Row
          </button>
          <button className="pbo-btn" onClick={addCol} title="Add column">
            ＋ Col
          </button>
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={addSheet} title="New sheet">
            ＋ Sheet
          </button>
          <button
            className="pbo-btn"
            onClick={renameSheet}
            title="Rename sheet"
          >
            ✎ Rename
          </button>
          <button
            className="pbo-btn"
            onClick={deleteSheet}
            title="Delete sheet"
          >
            🗑 Delete
          </button>
        </div>
        <div className="pbo-group pbo-group-note-stack">
          <button className="pbo-btn pbo-btn-advanced" onClick={openAdvancedEditor} disabled={busy} title="Open the full OnlyOffice editor">
            ↗ Advanced Editor
          </button>
        </div>
      </div>

      <div className="pbo-calc-formula">
        <div className="pbo-name-box">{selectedRef}</div>
        <span className="pbo-fx">ƒx</span>
        <input
          ref={formulaRef}
          className="pbo-formula-input"
          value={
            editing && editing.row === sel.row && editing.col === sel.col
              ? editing.text
              : cellRaw(selectedCell)
          }
          onChange={(e) => {
            if (!editing || editing.row !== sel.row || editing.col !== sel.col) {
              setEditing({ row: sel.row, col: sel.col, text: e.target.value });
            } else {
              setEditing({ ...editing, text: e.target.value });
            }
          }}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              commitEdit();
              setSel((s) => ({ ...s, row: s.row + 1 }));
            } else if (e.key === "Escape") {
              cancelEdit();
            }
          }}
          onBlur={() => {
            if (editing && editing.row === sel.row && editing.col === sel.col) {
              commitEdit();
            }
          }}
        />
      </div>

      {error && <div className="pbo-error">{error}</div>}

      <CalcGrid
        sheet={sheet}
        doc={doc}
        maxRow={maxRow}
        maxCol={maxCol}
        sel={sel}
        editing={editing}
        onSelect={(r, c) => setSel({ row: r, col: c })}
        onBeginEdit={beginEdit}
        onChangeEditing={(t) =>
          editing ? setEditing({ ...editing, text: t }) : undefined
        }
        onCommitEdit={(advance) => {
          if (!editing) return;
          const { row, col } = editing;
          commitEdit();
          if (advance === "row") setSel({ row: row + 1, col });
          else if (advance === "col") setSel({ row, col: col + 1 });
        }}
        onCancelEdit={cancelEdit}
      />

      <div className="pbo-calc-tabs">
        {doc.sheets.map((s, i) => (
          <button
            key={i}
            className={i === activeIdx ? "pbo-active-sheet" : ""}
            onClick={() => setActiveIdx(i)}
          >
            {s.name || `Sheet${i + 1}`}
          </button>
        ))}
        <button className="pbo-add-sheet" onClick={addSheet} title="New sheet">
          ＋
        </button>
      </div>

      <div className="pbo-status">
        <span>Cell: {selectedRef}</span>
        <span>
          Sheet: {sheet.name} ({activeIdx + 1} / {doc.sheets.length})
        </span>
        <div className="pbo-status-spacer" />
        <span>XLSX</span>
      </div>
    </div>
  );
}
