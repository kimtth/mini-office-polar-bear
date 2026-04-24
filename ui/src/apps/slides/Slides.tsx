import { useCallback, useEffect, useMemo, useState } from "react";
import { useNavigate } from "react-router-dom";
import { launchAdvancedEditor } from "../advanced/advancedEditor";
import MenuBar, { type Menu } from "../../components/MenuBar";
import FixturePicker from "../../components/FixturePicker";
import { useHiddenFilePicker } from "../../components/FilePicker";
import {
  createPptx,
  downloadBytes,
  fetchFixture,
  MIME,
  parsePptx,
  readFileBytes,
} from "../../wasm/loader";
import type {
  Fill,
  GroupShape,
  Shape,
  ShapeType,
  Slide,
  SlideElement,
  SlidesDocument,
  SlidesImage,
  TextBox,
  TextParagraph,
} from "../../wasm/types";
import "./slides.css";

const EMU_PER_PX = 9525;
const emu = (px: number) => Math.round(px * EMU_PER_PX);
const px = (e: number) => e / EMU_PER_PX;

function fillToCss(fill: Fill | null | undefined): string | undefined {
  if (!fill) return undefined;
  if (fill === "None") return "transparent";
  if ("Solid" in fill) return `#${fill.Solid}`;
  if ("Gradient" in fill && fill.Gradient.length) {
    const stops = fill.Gradient.map(
      (g) => `#${g.color} ${(g.position / 1000).toFixed(1)}%`,
    ).join(", ");
    return `linear-gradient(135deg, ${stops})`;
  }
  return undefined;
}

function textAlignCss(a: TextBox["align"]): React.CSSProperties["textAlign"] {
  if (a === "Center") return "center";
  if (a === "Right") return "right";
  if (a === "Justify") return "justify";
  if (a === "Left") return "left";
  return undefined;
}

function findImage(doc: SlidesDocument, id: string): SlidesImage | undefined {
  return doc.images.find((i) => i.id === id);
}

function emptySlide(): Slide {
  return { elements: [], notes: null, layout_idx: null, background: null };
}
function emptyDoc(): SlidesDocument {
  return {
    metadata: {
      title: null,
      author: null,
      last_modified_by: null,
      created: null,
      modified: null,
    },
    slides: [emptySlide()],
    slide_width: emu(960),
    slide_height: emu(540),
    layouts: [],
    masters: [],
    theme: null,
    images: [],
  };
}

function newTextBox(text = "Click to edit"): SlideElement {
  return {
    TextBox: {
      text,
      x: emu(80),
      y: emu(80),
      width: emu(420),
      height: emu(60),
      font_size: 24,
      bold: false,
      italic: false,
      color: null,
      align: null,
      rotation: 0,
      placeholder: null,
      placeholder_idx: null,
      paragraphs: [],
      fill: null,
      outline: null,
    },
  };
}

function newShape(kind: ShapeType, color: string): SlideElement {
  return {
    Shape: {
      shape_type: kind,
      x: emu(120),
      y: emu(200),
      width: emu(220),
      height: emu(140),
      rotation: 0,
      fill: { Solid: color },
      outline: null,
      text: null,
    },
  };
}

function shapeBorderRadius(t: ShapeType): string | number {
  if (t === "Ellipse") return "50%";
  if (t === "RoundedRectangle") return 10;
  return 4;
}

/** Return an SVG path for a preset shape geometry, relative to a 100x100 box. */
function presetShapePath(t: ShapeType): string | null {
  switch (t) {
    case "Triangle":
      return "M 50 0 L 100 100 L 0 100 Z";
    case "RightTriangle":
      return "M 0 0 L 100 100 L 0 100 Z";
    case "Diamond":
      return "M 50 0 L 100 50 L 50 100 L 0 50 Z";
    case "Parallelogram":
      return "M 20 0 L 100 0 L 80 100 L 0 100 Z";
    case "Trapezoid":
      return "M 20 0 L 80 0 L 100 100 L 0 100 Z";
    case "Pentagon":
      return "M 50 0 L 100 38 L 80 100 L 20 100 L 0 38 Z";
    case "Hexagon":
      return "M 25 0 L 75 0 L 100 50 L 75 100 L 25 100 L 0 50 Z";
    case "Octagon":
      return "M 30 0 L 70 0 L 100 30 L 100 70 L 70 100 L 30 100 L 0 70 L 0 30 Z";
    case "Star5":
      return "M 50 0 L 61 35 L 98 35 L 68 57 L 79 91 L 50 70 L 21 91 L 32 57 L 2 35 L 39 35 Z";
    case "Arrow":
      return "M 0 35 L 60 35 L 60 10 L 100 50 L 60 90 L 60 65 L 0 65 Z";
    case "Callout":
      return "M 0 0 L 100 0 L 100 70 L 60 70 L 50 100 L 40 70 L 0 70 Z";
    case "Line":
      return null;
    default:
      return null;
  }
}

// Build a composite key matching OnlyOffice SlideLayout::GetLevelUp:
// match by (type, idx) pair; normalize ctrTitle → title; default type = "body"
function phKey(type: string | null | undefined, idx: string | null | undefined): string {
  const t = (type === 'ctrTitle' ? 'title' : type) ?? 'body';
  const i = idx ?? '';
  return i ? `${t}|${i}` : t;
}

function placeholderKey(el: SlideElement): string | null {
  if ('TextBox' in el) {
    const ph = el.TextBox.placeholder;
    if (!ph) return null;
    return phKey(ph, el.TextBox.placeholder_idx);
  }
  if ('Shape' in el) {
    const ph = el.Shape.text?.placeholder;
    if (!ph) return null;
    return phKey(ph, el.Shape.text?.placeholder_idx);
  }
  return null;
}

function filterInheritedElements(elements: SlideElement[]): SlideElement[] {
  // Layout/master placeholders are edit-time prompts ("Click to add title",
  // "マスター テキストの書式設定") — skip them and only inherit decorative
  // chrome (backgrounds, logos, non-placeholder shapes/pictures/connectors).
  return elements.filter((el) => placeholderKey(el) === null);
}

function findPlaceholderGeometry(
  elements: SlideElement[],
  key: string,
): { x: number; y: number; width: number; height: number } | null {
  for (const el of elements) {
    const k = placeholderKey(el);
    if (k !== key) continue;
    if ('TextBox' in el) {
      const t = el.TextBox;
      if (t.width > 0 && t.height > 0) return { x: t.x, y: t.y, width: t.width, height: t.height };
    }
    if ('Shape' in el) {
      const s = el.Shape;
      if (s.width > 0 && s.height > 0) return { x: s.x, y: s.y, width: s.width, height: s.height };
    }
  }
  return null;
}

function inheritPlaceholderGeometry(
  elements: SlideElement[],
  layoutElements: SlideElement[],
  masterElements: SlideElement[],
): SlideElement[] {
  return elements.map((el) => {
    if ('TextBox' in el) {
      const t = el.TextBox;
      if (t.placeholder && (t.width === 0 || t.height === 0)) {
        const key = phKey(t.placeholder, t.placeholder_idx);
        const geom =
          findPlaceholderGeometry(layoutElements, key) ??
          findPlaceholderGeometry(masterElements, key);
        if (geom) return { TextBox: { ...t, ...geom } };
      }
    }
    if ('Shape' in el) {
      const s = el.Shape;
      const ph = s.text?.placeholder;
      if (ph && (s.width === 0 || s.height === 0)) {
        const key = phKey(ph, s.text?.placeholder_idx);
        const geom =
          findPlaceholderGeometry(layoutElements, key) ??
          findPlaceholderGeometry(masterElements, key);
        if (geom) return { Shape: { ...s, ...geom } };
      }
    }
    return el;
  });
}

function effectiveSlideElements(doc: SlidesDocument, slide: Slide): SlideElement[] {
  const layout =
    slide.layout_idx !== null && slide.layout_idx !== undefined
      ? doc.layouts[slide.layout_idx]
      : undefined;
  const master =
    layout?.master_idx !== null && layout?.master_idx !== undefined
      ? doc.masters[layout.master_idx]
      : undefined;
  const masterElements = filterInheritedElements(master?.elements ?? []);
  const layoutElements = filterInheritedElements(layout?.elements ?? []);
  const slideElements = inheritPlaceholderGeometry(
    slide.elements,
    layout?.elements ?? [],
    master?.elements ?? [],
  );
  return [...masterElements, ...layoutElements, ...slideElements];
}

/** Transform a child element's coordinates from group child space to slide space. */
function transformGroupChild(el: SlideElement, group: GroupShape): SlideElement {
  const hasTransform = group.ch_ext_cx > 0 && group.ch_ext_cy > 0;
  if (!hasTransform) return el;
  const sx = group.width / group.ch_ext_cx;
  const sy = group.height / group.ch_ext_cy;
  const tx = (x: number) => group.x + (x - group.ch_off_x) * sx;
  const ty = (y: number) => group.y + (y - group.ch_off_y) * sy;
  const tw = (w: number) => w * sx;
  const th = (h: number) => h * sy;
  if ('TextBox' in el) {
    const t = el.TextBox;
    return { TextBox: { ...t, x: tx(t.x), y: ty(t.y), width: tw(t.width), height: th(t.height) } };
  }
  if ('Shape' in el) {
    const s = el.Shape;
    return { Shape: { ...s, x: tx(s.x), y: ty(s.y), width: tw(s.width), height: th(s.height) } };
  }
  if ('Picture' in el) {
    const p = el.Picture;
    return { Picture: { ...p, x: tx(p.x), y: ty(p.y), width: tw(p.width), height: th(p.height) } };
  }
  if ('Connector' in el) {
    const c = el.Connector;
    return { Connector: { ...c, x: tx(c.x), y: ty(c.y), width: tw(c.width), height: th(c.height) } };
  }
  if ('Group' in el) {
    const g = el.Group;
    return { Group: { ...g, x: tx(g.x), y: ty(g.y), width: tw(g.width), height: th(g.height) } };
  }
  if ('Table' in el) {
    const t = el.Table;
    return { Table: { ...t, x: tx(t.x), y: ty(t.y), width: tw(t.width), height: th(t.height) } };
  }
  return el;
}

function renderGroupChildren(
  doc: SlidesDocument,
  group: GroupShape,
  groupKey: number,
  scale: number,
): React.ReactNode[] {
  return group.elements.map((child, index) => {
    const transformed = transformGroupChild(child, group);
    return renderElement(doc, transformed, groupKey * 1000 + index, scale, false, () => {});
  });
}

/** Render rich text paragraphs with per-run styling. Falls back to flat text if no paragraphs. */
function renderRichText(
  paragraphs: TextParagraph[] | undefined,
  flatText: string,
  scale: number,
  defaultStyle: { fontSize?: number | null; bold?: boolean; italic?: boolean; color?: string | null; align?: TextBox["align"] },
): React.ReactNode {
  if (!paragraphs || paragraphs.length === 0) {
    return flatText;
  }
  return paragraphs.map((para, pi) => (
    <div
      key={pi}
      style={{ textAlign: textAlignCss(para.align ?? defaultStyle.align ?? null) }}
    >
      {para.runs.length === 0 ? '\u00A0' : para.runs.map((run, ri) => (
        <span
          key={ri}
          style={{
            fontWeight: run.bold ? 700 : undefined,
            fontStyle: run.italic ? "italic" : undefined,
            fontSize: run.font_size ? `${run.font_size * scale}pt` : undefined,
            color: run.color ? `#${run.color}` : undefined,
          }}
        >
          {run.text}
        </span>
      ))}
    </div>
  ));
}

function renderElement(
  doc: SlidesDocument,
  el: SlideElement,
  i: number,
  scale: number,
  selected: boolean,
  onSelect: () => void,
  onEditText?: (text: string) => void,
  editable = false,
): React.ReactNode {
  if ("TextBox" in el) {
    const t = el.TextBox;
    const bgFill = fillToCss(t.fill);
    return (
      <div
        key={i}
        className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
        style={{
          left: px(t.x) * scale,
          top: px(t.y) * scale,
          width: px(t.width) * scale,
          height: px(t.height) * scale,
          background: bgFill,
          transform: t.rotation
            ? `rotate(${t.rotation / 60000}deg)`
            : undefined,
          boxShadow: t.outline?.color
            ? `inset 0 0 0 ${Math.max(1, t.outline.width_emu / EMU_PER_PX)}px #${t.outline.color}`
            : undefined,
        }}
        onMouseDown={onSelect}
      >
        <div
          className="pbo-el-text"
          contentEditable={editable}
          suppressContentEditableWarning
          style={{
            fontSize: t.font_size ? `${t.font_size * scale}pt` : undefined,
            fontWeight: t.bold ? 700 : undefined,
            fontStyle: t.italic ? "italic" : undefined,
            color: t.color ? `#${t.color}` : undefined,
            textAlign: textAlignCss(t.align),
          }}
          onBlur={(e) => onEditText?.(e.currentTarget.innerText)}
        >
          {renderRichText(t.paragraphs, t.text, scale, { fontSize: t.font_size, bold: t.bold, italic: t.italic, color: t.color, align: t.align })}
        </div>
      </div>
    );
  }
  if ("Shape" in el) {
    const s = el.Shape;
    const innerText = s.text;
    const path = presetShapePath(s.shape_type);
    const outlineColor = s.outline?.color ? `#${s.outline.color}` : "rgba(0,0,0,0.15)";
    const outlineWidth = Math.max(1, (s.outline?.width_emu ?? 0) / EMU_PER_PX);
    const fillCss = fillToCss(s.fill) ?? "transparent";

    // If preset path exists, render shape body via SVG and overlay text.
    if (path) {
      return (
        <div
          key={i}
          className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
          style={{
            left: px(s.x) * scale,
            top: px(s.y) * scale,
            width: px(s.width) * scale,
            height: px(s.height) * scale,
            transform: s.rotation ? `rotate(${s.rotation / 60000}deg)` : undefined,
          }}
          onMouseDown={onSelect}
        >
          <svg
            viewBox="0 0 100 100"
            preserveAspectRatio="none"
            style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
          >
            <path d={path} fill={fillCss} stroke={s.outline ? outlineColor : "none"} strokeWidth={outlineWidth / 2} />
          </svg>
          {innerText && (
            <span
              style={{
                position: "absolute",
                inset: 0,
                display: "grid",
                placeItems: "center",
                padding: 4,
                fontSize: innerText.font_size ? `${innerText.font_size * scale}pt` : undefined,
                fontWeight: innerText.bold ? 700 : undefined,
                fontStyle: innerText.italic ? "italic" : undefined,
                color: innerText.color ? `#${innerText.color}` : undefined,
                textAlign: textAlignCss(innerText.align),
                overflow: "hidden",
              }}
            >
              {renderRichText(innerText.paragraphs, innerText.text, scale, { fontSize: innerText.font_size, bold: innerText.bold, italic: innerText.italic, color: innerText.color, align: innerText.align })}
            </span>
          )}
        </div>
      );
    }

    return (
      <div
        key={i}
        className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
        style={{
          left: px(s.x) * scale,
          top: px(s.y) * scale,
          width: px(s.width) * scale,
          height: px(s.height) * scale,
          background: fillCss,
          borderRadius: shapeBorderRadius(s.shape_type),
          boxShadow: s.outline?.color
            ? `inset 0 0 0 ${outlineWidth}px ${outlineColor}`
            : undefined,
          transform: s.rotation
            ? `rotate(${s.rotation / 60000}deg)`
            : undefined,
          display: innerText ? "flex" : undefined,
          flexDirection: innerText ? "column" : undefined,
          justifyContent: innerText ? "center" : undefined,
          padding: innerText ? 4 : undefined,
          overflow: "hidden",
        }}
        onMouseDown={onSelect}
      >
        {innerText && (
          <div
            style={{
              fontSize: innerText.font_size
                ? `${innerText.font_size * scale}pt`
                : undefined,
              fontWeight: innerText.bold ? 700 : undefined,
              fontStyle: innerText.italic ? "italic" : undefined,
              color: innerText.color ? `#${innerText.color}` : undefined,
              textAlign: textAlignCss(innerText.align),
              width: "100%",
            }}
          >
            {renderRichText(innerText.paragraphs, innerText.text, scale, { fontSize: innerText.font_size, bold: innerText.bold, italic: innerText.italic, color: innerText.color, align: innerText.align })}
          </div>
        )}
      </div>
    );
  }
  if ("Picture" in el) {
    const p = el.Picture;
    const img = findImage(doc, p.image_id);
    return (
      <div
        key={i}
        className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
        style={{
          left: px(p.x) * scale,
          top: px(p.y) * scale,
          width: px(p.width) * scale,
          height: px(p.height) * scale,
          transform: p.rotation
            ? `rotate(${p.rotation / 60000}deg)`
            : undefined,
          backgroundImage: img
            ? `url(data:${img.mime};base64,${img.data_base64})`
            : undefined,
          backgroundSize: "contain",
          backgroundRepeat: "no-repeat",
          backgroundPosition: "center",
          backgroundColor: img ? undefined : "#eef0f5",
        }}
        onMouseDown={onSelect}
      >
        {!img && (
          <div
            style={{
              display: "grid",
              placeItems: "center",
              height: "100%",
              color: "var(--pbo-text-muted)",
              fontSize: 12,
            }}
          >
            🖼 {p.name ?? p.image_id}
          </div>
        )}
      </div>
    );
  }
  if ("Connector" in el) {
    const c = el.Connector;
    const col = c.outline?.color ? `#${c.outline.color}` : "#555";
    const sw = Math.max(1, (c.outline?.width_emu ?? 9525) / EMU_PER_PX);
    const dashArray = c.outline?.dash === "dash" ? "6 4" : c.outline?.dash === "dot" ? "2 4" : c.outline?.dash === "lgDash" ? "10 4" : undefined;
    const hasTailArrow = c.outline?.tail_end && c.outline.tail_end !== "none";
    const hasHeadArrow = c.outline?.head_end && c.outline.head_end !== "none";
    const markerId = `arrow-${i}`;
    return (
      <svg
        key={i}
        className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
        style={{
          left: px(c.x) * scale,
          top: px(c.y) * scale,
          width: Math.max(1, px(c.width) * scale),
          height: Math.max(1, px(c.height) * scale),
          overflow: "visible",
          transform: c.rotation ? `rotate(${c.rotation / 60000}deg)` : undefined,
        }}
        onMouseDown={onSelect}
        viewBox={`0 0 ${Math.max(1, px(c.width))} ${Math.max(1, px(c.height))}`}
        preserveAspectRatio="none"
      >
        {(hasTailArrow || hasHeadArrow) && (
          <defs>
            <marker id={markerId} viewBox="0 0 10 10" refX="10" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
              <path d="M 0 0 L 10 5 L 0 10 z" fill={col} />
            </marker>
          </defs>
        )}
        <line
          x1={c.flip_h ? px(c.width) : 0}
          y1={c.flip_v ? px(c.height) : 0}
          x2={c.flip_h ? 0 : px(c.width)}
          y2={c.flip_v ? 0 : px(c.height)}
          stroke={col}
          strokeWidth={sw}
          strokeDasharray={dashArray}
          markerEnd={hasTailArrow ? `url(#${markerId})` : undefined}
          markerStart={hasHeadArrow ? `url(#${markerId})` : undefined}
        />
      </svg>
    );
  }
  if ("Table" in el) {
    const t = el.Table;
    return (
      <div
        key={i}
        className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
        style={{
          left: px(t.x) * scale,
          top: px(t.y) * scale,
          width: px(t.width) * scale,
          height: px(t.height) * scale,
          overflow: "hidden",
          background: "#fff",
        }}
        onMouseDown={onSelect}
      >
        <table
          style={{
            width: "100%",
            height: "100%",
            borderCollapse: "collapse",
            fontSize: 10 * scale,
          }}
        >
          <tbody>
            {t.rows.map((row, ri) => (
              <tr key={ri}>
                {row.cells.map((cell, ci) => (
                  <td
                    key={ci}
                    colSpan={cell.col_span > 1 ? cell.col_span : undefined}
                    rowSpan={cell.row_span > 1 ? cell.row_span : undefined}
                    style={{
                      border: "1px solid #ccc",
                      padding: "2px 4px",
                      background: fillToCss(cell.fill),
                    }}
                  >
                    {cell.text}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }
  if ("Group" in el) {
    const g = el.Group;
    return [
      selected ? (
        <div
          key={`group-box-${i}`}
          className="pbo-slide-el pbo-el-selected"
          style={{
            left: px(g.x) * scale,
            top: px(g.y) * scale,
            width: px(g.width) * scale,
            height: px(g.height) * scale,
            pointerEvents: "none",
          }}
          onMouseDown={onSelect}
        />
      ) : null,
      ...renderGroupChildren(doc, g, i, scale),
    ];
  }
  return null;
}

function ThumbnailPreview({
  doc,
  slide,
  w,
  h,
}: {
  doc: SlidesDocument;
  slide: Slide;
  w: number;
  h: number;
}) {
  const scale = Math.min(180 / w, 96 / h);
  const elements = effectiveSlideElements(doc, slide);
  return (
    <div
      className="pbo-thumb-preview"
      style={{
        width: w * scale,
        height: h * scale,
        margin: "0 auto",
        background: fillToCss(slide.background) ?? "#fff",
      }}
    >
      {elements.map((el, i) =>
        renderElement(doc, el, i, scale, false, () => {}),
      )}
    </div>
  );
}

export default function Slides() {
  const nav = useNavigate();
  const [doc, setDoc] = useState<SlidesDocument>(emptyDoc());
  const [filename, setFilename] = useState<string>("Untitled.pptx");
  const [activeIdx, setActiveIdx] = useState(0);
  const [selEl, setSelEl] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(false);

  const slide = doc.slides[activeIdx] ?? doc.slides[0];
  const renderedElements = useMemo(
    () => effectiveSlideElements(doc, slide),
    [doc, slide],
  );
  const slideW = px(doc.slide_width);
  const slideH = px(doc.slide_height);

  const setSlides = useCallback(
    (update: (slides: Slide[]) => Slide[]) => {
      setDoc((d) => ({ ...d, slides: update(d.slides) }));
    },
    [],
  );

  const setActiveSlide = useCallback(
    (patch: (s: Slide) => Slide) => {
      setSlides((slides) =>
        slides.map((s, i) => (i === activeIdx ? patch(s) : s)),
      );
    },
    [activeIdx, setSlides],
  );

  // ── File ops ────────────────────────────────────
  const openFromFile = useCallback(async (file: File) => {
    setError(null);
    setBusy(true);
    try {
      const bytes = await readFileBytes(file);
      const parsed = await parsePptx(bytes);
      setDoc(parsed.slides.length ? parsed : emptyDoc());
      setActiveIdx(0);
      setSelEl(null);
      setFilename(file.name);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }, []);

  const { el: hiddenInput, open: openPicker } = useHiddenFilePicker(
    ".pptx",
    openFromFile,
  );

  const handleSave = useCallback(async () => {
    setError(null);
    setBusy(true);
    try {
      const bytes = await createPptx(doc);
      downloadBytes(bytes, filename, MIME.pptx);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }, [doc, filename]);

  const handleNew = useCallback(() => {
    setDoc(emptyDoc());
    setFilename("Untitled.pptx");
    setActiveIdx(0);
    setSelEl(null);
  }, []);

  // ── Slide ops ───────────────────────────────────
  const addSlide = () => {
    setSlides((slides) => [...slides, emptySlide()]);
    setActiveIdx(doc.slides.length);
    setSelEl(null);
  };
  const duplicateSlide = () => {
    const copy: Slide = {
      elements: slide.elements.map((e) => ({ ...e })),
      notes: slide.notes,
      layout_idx: slide.layout_idx,
      background: slide.background,
    };
    setSlides((slides) => [
      ...slides.slice(0, activeIdx + 1),
      copy,
      ...slides.slice(activeIdx + 1),
    ]);
    setActiveIdx((i) => i + 1);
    setSelEl(null);
  };
  const deleteSlide = (idx: number) => {
    if (doc.slides.length <= 1) {
      alert("A presentation needs at least one slide.");
      return;
    }
    const totalAfter = doc.slides.length - 1;
    setSlides((slides) => slides.filter((_, i) => i !== idx));
    setActiveIdx((cur) => {
      let next = cur;
      if (idx < cur) next = cur - 1;
      else if (idx === cur && cur >= totalAfter) next = totalAfter - 1;
      return Math.max(0, next);
    });
    setSelEl(null);
  };
  const moveSlide = (dir: -1 | 1) => {
    const j = activeIdx + dir;
    if (j < 0 || j >= doc.slides.length) return;
    setSlides((slides) => {
      const copy = slides.slice();
      [copy[activeIdx], copy[j]] = [copy[j], copy[activeIdx]];
      return copy;
    });
    setActiveIdx(j);
  };

  // ── Element ops ─────────────────────────────────
  const insertTextBox = () => {
    setActiveSlide((s) => ({ ...s, elements: [...s.elements, newTextBox()] }));
  };
  const insertShape = (kind: ShapeType, color: string) => {
    setActiveSlide((s) => ({
      ...s,
      elements: [...s.elements, newShape(kind, color)],
    }));
  };
  const deleteSelected = () => {
    if (selEl === null) return;
    setActiveSlide((s) => ({
      ...s,
      elements: s.elements.filter((_, i) => i !== selEl),
    }));
    setSelEl(null);
  };

  const updateTextBox = (idx: number, text: string) => {
    setActiveSlide((s) => ({
      ...s,
      elements: s.elements.map((el, i) => {
        if (i !== idx || !("TextBox" in el)) return el;
        const tb: TextBox = { ...el.TextBox, text };
        return { TextBox: tb };
      }),
    }));
  };

  const toggleBoldOnSelected = () => {
    if (selEl === null) return;
    setActiveSlide((s) => ({
      ...s,
      elements: s.elements.map((el, i) => {
        if (i !== selEl || !("TextBox" in el)) return el;
        const tb: TextBox = { ...el.TextBox, bold: !el.TextBox.bold };
        return { TextBox: tb };
      }),
    }));
  };

  const changeFontSize = (delta: number) => {
    if (selEl === null) return;
    setActiveSlide((s) => ({
      ...s,
      elements: s.elements.map((el, i) => {
        if (i !== selEl || !("TextBox" in el)) return el;
        const current = el.TextBox.font_size ?? 18;
        const tb: TextBox = {
          ...el.TextBox,
          font_size: Math.max(8, current + delta),
        };
        return { TextBox: tb };
      }),
    }));
  };

  const setShapeFill = (hex: string) => {
    if (selEl === null) return;
    setActiveSlide((s) => ({
      ...s,
      elements: s.elements.map((el, i) => {
        if (i !== selEl || !("Shape" in el)) return el;
        const sh: Shape = {
          ...el.Shape,
          fill: { Solid: hex.replace("#", "") },
        };
        return { Shape: sh };
      }),
    }));
  };

  const setNotes = (text: string) => {
    setActiveSlide((s) => ({ ...s, notes: text || null }));
  };

  // ── Keyboard shortcuts ──────────────────────────
  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
        e.preventDefault();
        handleSave();
      } else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "o") {
        e.preventDefault();
        openPicker();
      } else if (e.key === "Delete" && selEl !== null) {
        const target = e.target as HTMLElement | null;
        if (target && ["INPUT", "TEXTAREA"].includes(target.tagName)) return;
        if (target?.isContentEditable) return;
        e.preventDefault();
        deleteSelected();
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selEl, handleSave, openPicker]);

  const menus = useMemo<Menu[]>(
    () => [
      {
        label: "File",
        actions: [
          { label: "New", onClick: handleNew, kbd: "Ctrl+N" },
          { label: "Open…", onClick: openPicker, kbd: "Ctrl+O" },
          { label: "Open sample…", onClick: () => setPickerOpen(true) },
          { separator: true, label: "" },
          { label: "Save (.pptx)", onClick: handleSave, kbd: "Ctrl+S" },
          { separator: true, label: "" },
          { label: "Back to home", onClick: () => nav("/") },
        ],
      },
      {
        label: "Slide",
        actions: [
          { label: "New slide", onClick: addSlide },
          { label: "Duplicate slide", onClick: duplicateSlide },
          { label: "Delete slide", onClick: () => deleteSlide(activeIdx) },
          { separator: true, label: "" },
          { label: "Move up", onClick: () => moveSlide(-1) },
          { label: "Move down", onClick: () => moveSlide(1) },
        ],
      },
      {
        label: "Insert",
        actions: [
          { label: "Text box", onClick: insertTextBox },
          {
            label: "Rectangle",
            onClick: () => insertShape("Rectangle", "2b5ce6"),
          },
          {
            label: "Ellipse",
            onClick: () => insertShape("Ellipse", "c43d6b"),
          },
          { label: "Line", onClick: () => insertShape("Line", "1f8a4b") },
          { label: "Arrow", onClick: () => insertShape("Arrow", "b45309") },
        ],
      },
      {
        label: "Format",
        actions: [
          { label: "Toggle bold (selected)", onClick: toggleBoldOnSelected },
          { label: "Increase font size", onClick: () => changeFontSize(2) },
          { label: "Decrease font size", onClick: () => changeFontSize(-2) },
        ],
      },
    ],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [handleNew, openPicker, handleSave, nav, doc, activeIdx, selEl],
  );

  // ── Layout & scaling ─────────────────────────────
  const stageViewportWidth = 960;
  const stageViewportHeight = 540;
  const scale = Math.min(
    stageViewportWidth / slideW,
    stageViewportHeight / slideH,
  );

  const openAdvancedEditor = useCallback(async () => {
    setBusy(true);
    try {
      const bytes = await createPptx(doc);
      await launchAdvancedEditor({
        bytes,
        filename,
        kind: "pptx",
        navigate: nav,
      });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }, [doc, filename, nav]);

  return (
    <div className="pbo-app pbo-slides">
      {hiddenInput}
      <FixturePicker
        kind="pptx"
        open={pickerOpen}
        onClose={() => setPickerOpen(false)}
        onPick={async (f) => {
          setPickerOpen(false);
          setBusy(true);
          try {
            const bytes = await fetchFixture(f.kind, f.name);
            const parsed = await parsePptx(bytes);
            setDoc(parsed.slides.length ? parsed : emptyDoc());
            setActiveIdx(0);
            setSelEl(null);
            setFilename(f.name);
          } catch (e) {
            setError((e as Error).message);
          } finally {
            setBusy(false);
          }
        }}
      />

      <MenuBar
        menus={menus}
        appName="Slides"
        accent="var(--pbo-accent-slides)"
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
          <button className="pbo-btn" onClick={openPicker}>
            📂 Open
          </button>
          <button
            className="pbo-btn pbo-primary"
            onClick={handleSave}
            disabled={busy}
          >
            💾 Save
          </button>
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={addSlide}>
            ＋ Slide
          </button>
          <button className="pbo-btn" onClick={duplicateSlide}>
            ⎘ Duplicate
          </button>
          <button className="pbo-btn" onClick={() => deleteSlide(activeIdx)}>
            🗑 Delete
          </button>
        </div>
        <div className="pbo-group">
          <button className="pbo-btn" onClick={insertTextBox} title="Insert text">
            🅣 Text
          </button>
          <button
            className="pbo-btn"
            onClick={() => insertShape("Rectangle", "2b5ce6")}
            title="Insert rectangle"
          >
            ▭ Rect
          </button>
          <button
            className="pbo-btn"
            onClick={() => insertShape("Ellipse", "c43d6b")}
            title="Insert ellipse"
          >
            ◯ Ellipse
          </button>
        </div>
        <div className="pbo-group">
          <button
            className="pbo-btn"
            onClick={toggleBoldOnSelected}
            disabled={selEl === null}
          >
            <b>B</b>
          </button>
          <button
            className="pbo-btn"
            onClick={() => changeFontSize(-2)}
            disabled={selEl === null}
          >
            A−
          </button>
          <button
            className="pbo-btn"
            onClick={() => changeFontSize(2)}
            disabled={selEl === null}
          >
            A+
          </button>
          <input
            type="color"
            className="pbo-color"
            defaultValue="#2b5ce6"
            onChange={(e) => setShapeFill(e.target.value)}
            title="Shape fill"
            disabled={selEl === null}
          />
          <button
            className="pbo-btn"
            onClick={deleteSelected}
            disabled={selEl === null}
          >
            ✕ Delete
          </button>
        </div>
        <div className="pbo-group pbo-group-note-stack">
          <button className="pbo-btn pbo-btn-advanced" onClick={openAdvancedEditor} disabled={busy} title="Open the full OnlyOffice editor">
            ↗ Advanced Editor
          </button>
        </div>
      </div>

      {error && <div className="pbo-error">{error}</div>}

      <div className="pbo-slides-body">
        <aside className="pbo-slides-sidebar">
          <h3>Slides ({doc.slides.length})</h3>
          {doc.slides.map((s, i) => (
            <div
              key={i}
              className={`pbo-thumb${i === activeIdx ? " pbo-thumb-active" : ""}`}
              onClick={() => {
                setActiveIdx(i);
                setSelEl(null);
              }}
            >
              <button
                className="pbo-thumb-del"
                onClick={(e) => {
                  e.stopPropagation();
                  deleteSlide(i);
                }}
                title="Delete slide"
              >
                ✕
              </button>
              <ThumbnailPreview doc={doc} slide={s} w={slideW} h={slideH} />
              <div>
                Slide {i + 1} · {s.elements.length} element
                {s.elements.length === 1 ? "" : "s"}
              </div>
            </div>
          ))}
          <button className="pbo-btn" onClick={addSlide}>
            ＋ New slide
          </button>
        </aside>

        <section
          className="pbo-slides-stage"
          onMouseDown={() => setSelEl(null)}
        >
          <div className="pbo-slide-viewport">
            <div
              className="pbo-slide-canvas"
              style={{
                width: slideW * scale,
                height: slideH * scale,
                background: fillToCss(slide.background) ?? "#fff",
              }}
              onMouseDown={(e) => e.stopPropagation()}
            >
              {renderedElements.map((el, i) =>
                renderElement(
                  doc,
                  el,
                  i,
                  scale,
                  i < slide.elements.length && selEl === i,
                  () => {
                    if (i < slide.elements.length) {
                      setSelEl(i);
                    }
                  },
                  i < slide.elements.length ? (text) => updateTextBox(i, text) : undefined,
                  i < slide.elements.length && selEl === i && "TextBox" in el,
                ),
              )}
            </div>
          </div>
        </section>

        <aside className="pbo-slides-notes">
          <h3>Speaker notes</h3>
          <textarea
            value={slide.notes ?? ""}
            placeholder="Add notes for slide…"
            onChange={(e) => setNotes(e.target.value)}
          />
        </aside>
      </div>

      <div className="pbo-status">
        <span>
          Slide {activeIdx + 1} / {doc.slides.length}
        </span>
        <span>Elements: {slide.elements.length}</span>
        <div className="pbo-status-spacer" />
        <span>PPTX</span>
      </div>
    </div>
  );
}
