import {
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
  createDocx,
  downloadBytes,
  fetchFixture,
  MIME,
  parseDocx,
  readFileBytes,
} from "../../wasm/loader";
import type {
  Alignment,
  Block,
  DocumentMetadata,
  EmbeddedImage,
  Hyperlink,
  InlineImage,
  Paragraph,
  ParagraphStyle,
  Run,
  RunElement,
  RunStyle,
  Table,
  TableCell,
  WriterDocument,
} from "../../wasm/types";
import {
  cellBordersToCss,
  imageDataUri,
  listMarker,
  overlayRun,
  resolveListLevel,
  resolveParagraphStyle,
} from "../../wasm/renderHelpers";
import "./writer.css";

function emptyMetadata(): DocumentMetadata {
  return {
    title: null,
    author: null,
    last_modified_by: null,
    description: null,
    subject: null,
    keywords: null,
    created: null,
    modified: null,
  };
}

function emptyParagraphStyle(): ParagraphStyle {
  return {
    alignment: null,
    heading_level: null,
    style_id: null,
    indent: null,
    spacing: null,
    list: null,
  };
}

function emptyDoc(): WriterDocument {
  return {
    metadata: emptyMetadata(),
    body: [{ Paragraph: { runs: [], style: emptyParagraphStyle(), bookmarks: [] } }],
    section: null,
    resources: { images: [], hyperlinks: [] },
    numbering: { instances: [], abstract_nums: [] },
    styles: { styles: [] },
  };
}

function runStyleToCSS(style: RunStyle | null | undefined): React.CSSProperties {
  if (!style) return {};
  return {
    fontWeight: style.bold ? 700 : undefined,
    fontStyle: style.italic ? "italic" : undefined,
    textDecoration:
      [style.underline ? "underline" : null, style.strike ? "line-through" : null]
        .filter(Boolean)
        .join(" ") || undefined,
    fontSize: style.font_size ? `${style.font_size}pt` : undefined,
    fontFamily: style.font_family ?? undefined,
    color: style.color ? `#${style.color}` : undefined,
    backgroundColor: style.highlight ? `#${style.highlight}` : undefined,
  };
}

function alignToCSS(a: Alignment | null | undefined): React.CSSProperties["textAlign"] {
  if (a === "Center") return "center";
  if (a === "Right") return "right";
  if (a === "Justify") return "justify";
  return undefined;
}



function domToDoc(root: HTMLElement, prior: WriterDocument): WriterDocument {
  const body: Block[] = [];
  for (const node of Array.from(root.childNodes)) {
    const block = nodeToBlock(node);
    if (block) body.push(block);
  }
  if (body.length === 0) {
    body.push({ Paragraph: { runs: [], style: emptyParagraphStyle(), bookmarks: [] } });
  }
  return { ...prior, body };
}

function nodeToBlock(node: Node): Block | null {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? "";
    if (!text.trim()) return null;
    return {
      Paragraph: {
        runs: [{ Run: { text, style: null } }],
        style: emptyParagraphStyle(),
        bookmarks: [],
      },
    };
  }
  if (!(node instanceof HTMLElement)) return null;
  const tag = node.tagName.toLowerCase();
  if (tag === "hr") return "PageBreak";
  if (tag === "br") return null;

  const heading = /^h([1-6])$/.exec(tag);
  const rawAlign = node.style.textAlign.toLowerCase();
  const alignment: Alignment | null =
    rawAlign === "center"
      ? "Center"
      : rawAlign === "right"
      ? "Right"
      : rawAlign === "justify"
      ? "Justify"
      : null;

  const runs: Run[] = [];
  collectRuns(node, {}, runs);

  const style = emptyParagraphStyle();
  style.alignment = alignment;
  style.heading_level = heading ? Number(heading[1]) : null;

  const para: Paragraph = {
    runs: runs.map((r) => ({ Run: r })),
    style,
    bookmarks: [],
  };
  return { Paragraph: para };
}

function collectRuns(
  el: Node,
  inherited: Partial<RunStyle>,
  out: Run[],
): void {
  if (el.nodeType === Node.TEXT_NODE) {
    const text = el.textContent ?? "";
    if (!text) return;
    out.push({ text, style: pickStyle(inherited) });
    return;
  }
  if (!(el instanceof HTMLElement)) return;
  if (el.tagName === "BR") {
    out.push({ text: "\n", style: pickStyle(inherited) });
    return;
  }
  const style = { ...inherited };
  const tag = el.tagName.toLowerCase();
  if (
    tag === "b" ||
    tag === "strong" ||
    el.style.fontWeight === "bold" ||
    Number(el.style.fontWeight) >= 600
  ) {
    style.bold = true;
  }
  if (tag === "i" || tag === "em" || el.style.fontStyle === "italic") {
    style.italic = true;
  }
  if (tag === "u" || el.style.textDecoration?.includes("underline")) {
    style.underline = true;
  }
  if (el.style.fontFamily) {
    style.font_family = el.style.fontFamily.replace(/["']/g, "");
  }
  if (el.style.fontSize) {
    const m = /([0-9.]+)\s*pt/.exec(el.style.fontSize);
    if (m) style.font_size = Number(m[1]);
  }
  if (el.style.color) {
    const hex = rgbToHex(el.style.color);
    if (hex) style.color = hex;
  }
  for (const child of Array.from(el.childNodes)) {
    collectRuns(child, style, out);
  }
}

function pickStyle(s: Partial<RunStyle>): RunStyle | null {
  if (
    !s.bold &&
    !s.italic &&
    !s.underline &&
    !s.strike &&
    s.font_size == null &&
    s.font_family == null &&
    s.color == null &&
    s.highlight == null
  ) {
    return null;
  }
  return {
    bold: !!s.bold,
    italic: !!s.italic,
    underline: !!s.underline,
    strike: !!s.strike,
    font_size: s.font_size ?? null,
    font_family: s.font_family ?? null,
    color: s.color ?? null,
    highlight: s.highlight ?? null,
    vertical_align: s.vertical_align ?? null,
  };
}

function rgbToHex(rgb: string): string | null {
  const m = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/.exec(rgb);
  if (!m) return rgb.startsWith("#") ? rgb.slice(1) : null;
  const hex = (n: string) => Number(n).toString(16).padStart(2, "0");
  return `${hex(m[1])}${hex(m[2])}${hex(m[3])}`;
}

// ── Rich block rendering helpers ─────────────────────

/** Map an image-id to a resolved data URI (or null if not embedded). */
function resolveImage(
  doc: WriterDocument,
  id: string,
): EmbeddedImage | undefined {
  return doc.resources.images.find((i) => i.id === id);
}

/** Resolve a hyperlink rel-id to a URL. */
function resolveHyperlinkUrl(
  doc: WriterDocument,
  rel_id: string | null,
  anchor: string | null,
): string | undefined {
  if (anchor) return `#${anchor}`;
  if (!rel_id) return undefined;
  return doc.resources.hyperlinks.find((h) => h.id === rel_id)?.url;
}

function renderInlineElement(
  doc: WriterDocument,
  el: RunElement,
  key: number,
  inheritedRun: RunStyle | null,
): React.ReactNode {
  if ("Run" in el) {
    const style = overlayRun(inheritedRun, el.Run.style);
    return (
      <span key={key} style={runStyleToCSS(style)}>
        {el.Run.text}
      </span>
    );
  }
  if ("Hyperlink" in el) {
    const hl: Hyperlink = el.Hyperlink;
    const url = resolveHyperlinkUrl(doc, hl.rel_id, hl.anchor);
    const inner = hl.runs.map((r, i) => (
      <span key={i} style={runStyleToCSS(overlayRun(inheritedRun, r.style))}>
        {r.text}
      </span>
    ));
    return url ? (
      <a
        key={key}
        href={url}
        target="_blank"
        rel="noreferrer noopener"
        className="pbo-hyperlink"
      >
        {inner}
      </a>
    ) : (
      <span key={key} className="pbo-hyperlink">
        {inner}
      </span>
    );
  }
  if ("InlineImage" in el) {
    const img: InlineImage = el.InlineImage;
    const src = resolveImage(doc, img.image_id);
    if (!src) {
      return (
        <span key={key} className="pbo-image-missing">
          [image: {img.alt ?? img.image_id}]
        </span>
      );
    }
    const widthPx = Math.max(1, img.width_emu / 9525);
    const heightPx = Math.max(1, img.height_emu / 9525);
    return (
      <img
        key={key}
        src={imageDataUri(src.mime, src.data_base64)}
        alt={img.alt ?? ""}
        style={{
          width: widthPx,
          height: heightPx,
          maxWidth: "100%",
          verticalAlign: "middle",
        }}
      />
    );
  }
  return null;
}

function paragraphInlineStyle(
  style: ParagraphStyle | null,
): React.CSSProperties {
  if (!style) return {};
  const css: React.CSSProperties = {
    textAlign: alignToCSS(style.alignment ?? null),
  };
  if (style.indent) {
    const tw = (n: number) => (n * 1.333) / 20; // twips → px (1 pt ≈ 1.333 px)
    const left = Math.max(0, style.indent.left - style.indent.hanging);
    if (left > 0) css.marginLeft = tw(left);
    if (style.indent.right > 0) css.marginRight = tw(style.indent.right);
    if (style.indent.first_line > 0) css.textIndent = tw(style.indent.first_line);
  }
  if (style.spacing) {
    if (style.spacing.before > 0) css.marginTop = (style.spacing.before * 1.333) / 20;
    if (style.spacing.after > 0) css.marginBottom = (style.spacing.after * 1.333) / 20;
    if (style.spacing.line > 0) css.lineHeight = style.spacing.line / 240;
  }
  return css;
}

/** Counter state keyed by `${num_id}:${level}`. */
type ListCounters = Map<string, number>;

function resolveEffectiveStyle(
  doc: WriterDocument,
  p: Paragraph,
): { paragraph: ParagraphStyle | null; run: RunStyle | null } {
  const fromCat = resolveParagraphStyle(doc.styles, p.style?.style_id ?? null);
  const merged: ParagraphStyle | null = p.style
    ? { ...(fromCat.paragraph ?? {} as ParagraphStyle), ...cleanParagraph(p.style) }
    : fromCat.paragraph;
  return { paragraph: merged, run: fromCat.run };
}

function cleanParagraph(p: ParagraphStyle): ParagraphStyle {
  // Drop `null` keys so catalogue defaults survive Object spread.
  const out: Record<string, unknown> = { ...(p as unknown as Record<string, unknown>) };
  for (const k of Object.keys(out)) {
    if (out[k] == null) delete out[k];
  }
  return out as unknown as ParagraphStyle;
}

function renderTable(doc: WriterDocument, table: Table, key: number) {
  const tableBorders = table.style?.borders ?? null;
  return (
    <table
      key={key}
      className="pbo-writer-table"
      style={{
        ...cellBordersToCss(tableBorders),
      }}
    >
      <tbody>
        {table.rows.map((row, ri) => (
          <tr key={ri}>
            {row.cells.map((cell, ci) => renderTableCell(doc, cell, ci))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function renderTableCell(doc: WriterDocument, cell: TableCell, key: number) {
  const borders = cellBordersToCss(cell.borders);
  const style: React.CSSProperties = {
    ...borders,
    background: cell.shading ? `#${cell.shading}` : undefined,
    verticalAlign: "top",
    padding: "0.35em 0.5em",
  };
  // Vertical merge: skip `Continue` rows (model already has merges; we collapse via rowSpan on Restart).
  const skipRender = cell.v_merge === "Continue";
  return (
    <td
      key={key}
      colSpan={cell.grid_span > 1 ? cell.grid_span : undefined}
      style={style}
      hidden={skipRender ? true : undefined}
    >
      {cell.content.map((b, bi) => renderBlock(doc, b, bi, new Map()))}
    </td>
  );
}

function renderBlock(
  doc: WriterDocument,
  block: Block,
  i: number,
  counters: ListCounters,
): React.ReactNode {
  if (block === "PageBreak") return <hr key={i} className="pbo-page-break" />;
  if ("Table" in block) return renderTable(doc, block.Table, i);
  if (!("Paragraph" in block)) return null;

  const p = block.Paragraph;
  const { paragraph: effPara, run: effRun } = resolveEffectiveStyle(doc, p);
  const heading = effPara?.heading_level ?? null;

  // List marker resolution
  const listRef = effPara?.list;
  let marker: React.ReactNode = null;
  if (listRef) {
    const level = resolveListLevel(doc.numbering, listRef);
    if (level) {
      const key = `${listRef.num_id}:${listRef.level}`;
      const current = (counters.get(key) ?? 0) + 1;
      counters.set(key, current);
      marker = (
        <span className="pbo-list-marker" aria-hidden>
          {listMarker(level, current)}
        </span>
      );
    }
  }

  const paragraphCss: React.CSSProperties = {
    ...paragraphInlineStyle(effPara),
    ...runStyleToCSS(effRun),
  };
  const inlineNodes = p.runs.map((el, j) =>
    renderInlineElement(doc, el, j, effRun),
  );

  const content = marker ? (
    <>
      {marker}
      {inlineNodes.length ? inlineNodes : "\u00A0"}
    </>
  ) : inlineNodes.length ? (
    inlineNodes
  ) : (
    <br />
  );

  const bookmarkAnchors = p.bookmarks.map((bm) => (
    <a key={`bm-${bm.id}`} id={bm.name} aria-hidden />
  ));

  if (heading && heading >= 1 && heading <= 6) {
    const Tag = `h${heading}` as "h1" | "h2" | "h3" | "h4" | "h5" | "h6";
    return (
      <Tag key={i} style={paragraphCss} data-style-id={effPara?.style_id ?? undefined}>
        {bookmarkAnchors}
        {content}
      </Tag>
    );
  }
  return (
    <p
      key={i}
      style={paragraphCss}
      data-style-id={effPara?.style_id ?? undefined}
      className={listRef ? "pbo-list-para" : undefined}
    >
      {bookmarkAnchors}
      {content}
    </p>
  );
}

function WriterBody({ doc }: { doc: WriterDocument }) {
  const counters: ListCounters = new Map();
  return <>{doc.body.map((b, i) => renderBlock(doc, b, i, counters))}</>;
}

export default function Writer() {
  const nav = useNavigate();
  const [doc, setDoc] = useState<WriterDocument>(emptyDoc());
  const [filename, setFilename] = useState<string>("Untitled.docx");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [docVersion, setDocVersion] = useState(0);
  const [stats, setStats] = useState({ words: 0, chars: 0 });
  const editorRef = useRef<HTMLDivElement>(null);

  const openFromFile = useCallback(async (file: File) => {
    setError(null);
    setBusy(true);
    try {
      const bytes = await readFileBytes(file);
      const parsed = await parseDocx(bytes);
      setDoc(parsed.body.length ? parsed : emptyDoc());
      setFilename(file.name);
      setDocVersion((v) => v + 1);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }, []);

  const { el: hiddenInput, open: openPicker } = useHiddenFilePicker(
    ".docx",
    openFromFile,
  );

  const syncFromDom = useCallback(() => {
    if (!editorRef.current) return;
    setDoc((prior) => domToDoc(editorRef.current!, prior));
    const text = editorRef.current.innerText;
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    setStats({ words, chars: text.length });
  }, []);

  useEffect(() => {
    if (!editorRef.current) return;
    const text = editorRef.current.innerText;
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    setStats({ words, chars: text.length });
  }, [docVersion]);

  const handleSave = useCallback(async () => {
    setError(null);
    setBusy(true);
    try {
      const latest = editorRef.current ? domToDoc(editorRef.current, doc) : doc;
      const bytes = await createDocx(latest);
      downloadBytes(bytes, filename, MIME.docx);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }, [doc, filename]);

  const handleNew = useCallback(() => {
    setDoc(emptyDoc());
    setFilename("Untitled.docx");
    setDocVersion((v) => v + 1);
    setError(null);
  }, []);

  const handleRename = useCallback(() => {
    const name = prompt("File name:", filename);
    if (name) setFilename(name.endsWith(".docx") ? name : `${name}.docx`);
  }, [filename]);

  const exec = useCallback(
    (cmd: string, val?: string) => {
      editorRef.current?.focus();
      document.execCommand(cmd, false, val);
      syncFromDom();
    },
    [syncFromDom],
  );
  const bold = useCallback(() => exec("bold"), [exec]);
  const italic = useCallback(() => exec("italic"), [exec]);
  const underline = useCallback(() => exec("underline"), [exec]);
  const align = useCallback(
    (how: "Left" | "Center" | "Right" | "Justify") =>
      exec(
        how === "Left"
          ? "justifyLeft"
          : how === "Center"
          ? "justifyCenter"
          : how === "Right"
          ? "justifyRight"
          : "justifyFull",
      ),
    [exec],
  );
  const heading = useCallback(
    (level: number) => {
      editorRef.current?.focus();
      document.execCommand("formatBlock", false, level === 0 ? "P" : `H${level}`);
      syncFromDom();
    },
    [syncFromDom],
  );
  const setFont = useCallback((f: string) => exec("fontName", f), [exec]);
  const setSize = useCallback(
    (s: string) => {
      editorRef.current?.focus();
      const sel = window.getSelection();
      if (!sel || sel.rangeCount === 0 || sel.isCollapsed) {
        syncFromDom();
        return;
      }
      const range = sel.getRangeAt(0);
      const span = document.createElement("span");
      span.style.fontSize = `${s}pt`;
      span.appendChild(range.extractContents());
      range.insertNode(span);
      syncFromDom();
    },
    [syncFromDom],
  );
  const setColor = useCallback((c: string) => exec("foreColor", c), [exec]);
  const insertPageBreak = useCallback(
    () => exec("insertHorizontalRule"),
    [exec],
  );
  const undo = useCallback(() => exec("undo"), [exec]);
  const redo = useCallback(() => exec("redo"), [exec]);

  const menus = useMemo<Menu[]>(
    () => [
      {
        label: "File",
        actions: [
          { label: "New", onClick: handleNew, kbd: "Ctrl+N" },
          { label: "Open…", onClick: openPicker, kbd: "Ctrl+O" },
          { label: "Open sample…", onClick: () => setPickerOpen(true) },
          { separator: true, label: "" },
          { label: "Rename…", onClick: handleRename },
          { label: "Save (.docx)", onClick: handleSave, kbd: "Ctrl+S" },
          { separator: true, label: "" },
          { label: "Back to home", onClick: () => nav("/") },
        ],
      },
      {
        label: "Edit",
        actions: [
          { label: "Undo", onClick: undo, kbd: "Ctrl+Z" },
          { label: "Redo", onClick: redo, kbd: "Ctrl+Y" },
          { separator: true, label: "" },
          { label: "Cut", onClick: () => exec("cut"), kbd: "Ctrl+X" },
          { label: "Copy", onClick: () => exec("copy"), kbd: "Ctrl+C" },
          { label: "Paste", onClick: () => exec("paste"), kbd: "Ctrl+V" },
          { label: "Select all", onClick: () => exec("selectAll"), kbd: "Ctrl+A" },
        ],
      },
      {
        label: "Insert",
        actions: [
          { label: "Page break", onClick: insertPageBreak },
          { label: "Bulleted list", onClick: () => exec("insertUnorderedList") },
          { label: "Numbered list", onClick: () => exec("insertOrderedList") },
        ],
      },
      {
        label: "Format",
        actions: [
          { label: "Bold", onClick: bold, kbd: "Ctrl+B" },
          { label: "Italic", onClick: italic, kbd: "Ctrl+I" },
          { label: "Underline", onClick: underline, kbd: "Ctrl+U" },
          { separator: true, label: "" },
          { label: "Heading 1", onClick: () => heading(1) },
          { label: "Heading 2", onClick: () => heading(2) },
          { label: "Heading 3", onClick: () => heading(3) },
          { label: "Normal text", onClick: () => heading(0) },
        ],
      },
    ],
    [
      handleNew, openPicker, handleRename, handleSave, nav,
      undo, redo, exec, insertPageBreak, bold, italic, underline, heading,
    ],
  );

  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (!(e.ctrlKey || e.metaKey)) return;
      const k = e.key.toLowerCase();
      if (k === "s") { e.preventDefault(); handleSave(); }
      else if (k === "o") { e.preventDefault(); openPicker(); }
      else if (k === "n") { e.preventDefault(); handleNew(); }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [handleSave, openPicker, handleNew]);

  const openAdvancedEditor = useCallback(async () => {
    setBusy(true);
    try {
      const bytes = await createDocx(doc);
      await launchAdvancedEditor({
        bytes,
        filename,
        kind: "docx",
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
        kind="docx"
        open={pickerOpen}
        onClose={() => setPickerOpen(false)}
        onPick={async (f) => {
          setPickerOpen(false);
          setBusy(true);
          try {
            const bytes = await fetchFixture(f.kind, f.name);
            const parsed = await parseDocx(bytes);
            setDoc(parsed.body.length ? parsed : emptyDoc());
            setFilename(f.name);
            setDocVersion((v) => v + 1);
          } catch (e) {
            setError((e as Error).message);
          } finally {
            setBusy(false);
          }
        }}
      />

      <MenuBar
        menus={menus}
        appName="Writer"
        accent="var(--pbo-accent-writer)"
        rightSlot={
          <span style={{ color: "var(--pbo-text-muted)", fontSize: "0.8rem", padding: "0 0.75rem" }}>
            {filename}{busy && " · working…"}
          </span>
        }
      />

      <div className="pbo-ribbon">
        <div className="pbo-group">
          <button className="pbo-btn" onClick={openPicker} title="Open (Ctrl+O)">📂 Open</button>
          <button className="pbo-btn pbo-primary" onClick={handleSave} disabled={busy} title="Save (Ctrl+S)">💾 Save</button>
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={undo} title="Undo">↶</button>
          <button className="pbo-btn" onClick={redo} title="Redo">↷</button>
        </div>
        <div className="pbo-group">
          <select className="pbo-select" defaultValue="" onChange={(e) => {
            const v = e.target.value;
            if (v === "p") heading(0); else if (v) heading(Number(v));
            e.target.value = "";
          }} aria-label="Paragraph style">
            <option value="">Paragraph…</option>
            <option value="1">Heading 1</option>
            <option value="2">Heading 2</option>
            <option value="3">Heading 3</option>
            <option value="4">Heading 4</option>
            <option value="p">Normal</option>
          </select>
          <select className="pbo-select" defaultValue="" onChange={(e) => {
            if (e.target.value) setFont(e.target.value);
            e.target.value = "";
          }} aria-label="Font family">
            <option value="">Font…</option>
            {["Calibri","Cambria","Arial","Times New Roman","Georgia","Consolas"].map((f) => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
          <select className="pbo-select" defaultValue="" onChange={(e) => {
            if (e.target.value) setSize(e.target.value);
            e.target.value = "";
          }} aria-label="Font size">
            <option value="">Size…</option>
            {[8,9,10,11,12,14,16,18,20,24,28,32,36,48].map((n) => (
              <option key={n} value={String(n)}>{n}</option>
            ))}
          </select>
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={bold} title="Bold (Ctrl+B)"><b>B</b></button>
          <button className="pbo-btn" onClick={italic} title="Italic (Ctrl+I)"><i>I</i></button>
          <button className="pbo-btn" onClick={underline} title="Underline (Ctrl+U)"><u>U</u></button>
          <input type="color" className="pbo-color" defaultValue="#1d1d1d"
            onChange={(e) => setColor(e.target.value)} title="Text color" />
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={() => align("Left")} title="Align left">⯇</button>
          <button className="pbo-btn" onClick={() => align("Center")} title="Center">≡</button>
          <button className="pbo-btn" onClick={() => align("Right")} title="Align right">⯈</button>
          <button className="pbo-btn" onClick={() => align("Justify")} title="Justify">☰</button>
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={() => exec("insertUnorderedList")} title="Bulleted list">• List</button>
          <button className="pbo-btn" onClick={() => exec("insertOrderedList")} title="Numbered list">1. List</button>
          <button className="pbo-btn" onClick={insertPageBreak} title="Page break">⎯</button>
        </div>
        <div className="pbo-group pbo-group-note-stack">
          <button className="pbo-btn pbo-btn-advanced" onClick={openAdvancedEditor} disabled={busy} title="Open the full OnlyOffice editor">
            ↗ Advanced Editor
          </button>
        </div>
      </div>

      {error && <div className="pbo-error">{error}</div>}

      <div className="pbo-work">
        <div
          key={docVersion}
          ref={editorRef}
          className="pbo-writer-page"
          contentEditable
          suppressContentEditableWarning
          spellCheck
          onInput={syncFromDom}
          onBlur={syncFromDom}
        >
          <WriterBody doc={doc} />
        </div>
      </div>

      <div className="pbo-status">
        <span>Words: {stats.words}</span>
        <span>Characters: {stats.chars}</span>
        <div className="pbo-status-spacer" />
        <span>DOCX</span>
      </div>
    </div>
  );
}