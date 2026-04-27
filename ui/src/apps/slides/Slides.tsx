import { useCallback, useEffect, useMemo, useRef, useState } from "react";
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
  renderPptx,
  readFileBytes,
} from "../../wasm/loader";
import type {
  ConnectorType,
  Fill,
  GeometryAdjustment,
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
const PT_TO_PX = 96 / 72;
const SHAPE_FILL_SWATCHES = ["2b5ce6", "c43d6b", "1f8a4b", "f59e0b", "22c3e6", "6d5bd0", "111827"];

function fillToCss(fill: Fill | null | undefined): string | undefined {
  if (!fill) return undefined;
  if (fill === "None") return "transparent";
  if ("Solid" in fill) return `#${fill.Solid}`;
  if ("SolidAlpha" in fill) return colorWithAlpha(fill.SolidAlpha.color, fill.SolidAlpha.alpha);
  if ("Gradient" in fill && fill.Gradient.length) {
    const stops = fill.Gradient.map(
      (g) => `#${g.color} ${(g.position / 1000).toFixed(1)}%`,
    ).join(", ");
    return `linear-gradient(135deg, ${stops})`;
  }
  if ("Pattern" in fill) {
    const fg = fill.Pattern.foreground ? `#${fill.Pattern.foreground}` : "rgba(0,0,0,0.18)";
    const bg = fill.Pattern.background ? `#${fill.Pattern.background}` : "transparent";
    const preset = fill.Pattern.preset ?? "pct20";
    if (preset.includes("Horz")) {
      return `repeating-linear-gradient(0deg, ${fg} 0 1px, ${bg} 1px 6px)`;
    }
    if (preset.includes("Vert")) {
      return `repeating-linear-gradient(90deg, ${fg} 0 1px, ${bg} 1px 6px)`;
    }
    if (preset.includes("Diag") || preset.includes("diag")) {
      return `repeating-linear-gradient(45deg, ${fg} 0 1px, ${bg} 1px 6px)`;
    }
    return `repeating-linear-gradient(45deg, ${fg} 0 2px, ${bg} 2px 8px)`;
  }
  return undefined;
}

function alphaFromOoxml(alpha: number | null | undefined, fallback = 62000): number {
  return Math.max(0, Math.min(1, (alpha ?? fallback) / 100000));
}

function fillToSvgPaint(fill: Fill | null | undefined, id: string): string {
  if (fill && fill !== "None" && "Gradient" in fill && fill.Gradient.length) {
    return `url(#${id})`;
  }
  if (fill && fill !== "None" && "Pattern" in fill) {
    return fill.Pattern.foreground
      ? `#${fill.Pattern.foreground}`
      : fill.Pattern.background
        ? `#${fill.Pattern.background}`
        : "transparent";
  }
  return fillToCss(fill) ?? "transparent";
}

function svgFillDefs(fill: Fill | null | undefined, id: string) {
  if (!fill || fill === "None" || !("Gradient" in fill) || !fill.Gradient.length) return null;
  return (
    <defs>
      <linearGradient id={id} x1="0%" y1="0%" x2="100%" y2="100%">
        {fill.Gradient.map((stop, index) => (
          <stop
            key={`${id}-${index}`}
            offset={`${stop.position / 1000}%`}
            stopColor={`#${stop.color}`}
          />
        ))}
      </linearGradient>
    </defs>
  );
}

function rgba(hex: string | null | undefined, alpha?: number | null, fallback = 62000): string {
  const clean = (hex ?? "000000").replace("#", "").padEnd(6, "0").slice(0, 6);
  const r = parseInt(clean.slice(0, 2), 16) || 0;
  const g = parseInt(clean.slice(2, 4), 16) || 0;
  const b = parseInt(clean.slice(4, 6), 16) || 0;
  return `rgba(${r}, ${g}, ${b}, ${alphaFromOoxml(alpha, fallback)})`;
}

function colorWithAlpha(hex: string | null | undefined, alpha?: number | null): string {
  return alpha === null || alpha === undefined ? `#${hex ?? "000000"}` : rgba(hex, alpha, 100000);
}

function effectFilter(effects: Shape["effects"] | undefined, scale: number): string | undefined {
  if (!effects) return undefined;
  const filters: string[] = [];
  if (effects.shadow) {
    const angle = (effects.shadow.direction / 60000) * Math.PI / 180;
    const distance = px(effects.shadow.distance) * scale;
    const blur = px(effects.shadow.blur_radius) * scale;
    filters.push(`drop-shadow(${Math.cos(angle) * distance}px ${Math.sin(angle) * distance}px ${blur}px ${rgba(effects.shadow.color, effects.shadow.alpha)})`);
  }
  if (effects.glow) {
    const radius = Math.max(1, px(effects.glow.radius) * scale);
    filters.push(`drop-shadow(0 0 ${radius}px ${rgba(effects.glow.color, effects.glow.alpha)})`);
  }
  if (!filters.length) return undefined;
  return filters.join(" ");
}

function transformCss(rotation: number, flipH?: boolean, flipV?: boolean): string | undefined {
  const parts: string[] = [];
  if (rotation) parts.push(`rotate(${rotation / 60000}deg)`);
  if (flipH) parts.push("scaleX(-1)");
  if (flipV) parts.push("scaleY(-1)");
  return parts.length ? parts.join(" ") : undefined;
}

function textBoxPadding(
  body: TextBox["body"] | null | undefined,
  scale: number,
  width: number,
  height: number,
  fallback = 0,
  minContentHeight = 0,
): React.CSSProperties {
  const finite = (value: number, fallbackValue = 0) =>
    Number.isFinite(value) ? value : fallbackValue;
  let paddingLeft = body?.inset_left !== null && body?.inset_left !== undefined
    ? finite(px(body.inset_left) * scale, fallback)
    : fallback;
  let paddingTop = body?.inset_top !== null && body?.inset_top !== undefined
    ? finite(px(body.inset_top) * scale, fallback)
    : fallback;
  let paddingRight = body?.inset_right !== null && body?.inset_right !== undefined
    ? finite(px(body.inset_right) * scale, fallback)
    : fallback;
  let paddingBottom = body?.inset_bottom !== null && body?.inset_bottom !== undefined
    ? finite(px(body.inset_bottom) * scale, fallback)
    : fallback;

  const clampPair = (start: number, end: number, available: number) => {
    const safeAvailable = finite(available);
    const safeStart = finite(start);
    const safeEnd = finite(end);
    if (safeAvailable <= 0) return [0, 0] as const;
    const total = safeStart + safeEnd;
    if (total <= safeAvailable) return [safeStart, safeEnd] as const;
    const factor = Math.max(0, (safeAvailable - 1) / Math.max(total, 1));
    return [safeStart * factor, safeEnd * factor] as const;
  };

  [paddingLeft, paddingRight] = clampPair(paddingLeft, paddingRight, width);
  [paddingTop, paddingBottom] = clampPair(
    paddingTop,
    paddingBottom,
    Math.max(0, finite(height) - finite(minContentHeight)),
  );

  return {
    paddingLeft,
    paddingTop,
    paddingRight,
    paddingBottom,
  };
}

function textVerticalAlign(anchor: string | null | undefined): React.CSSProperties["justifyContent"] {
  if (anchor === "ctr" || anchor === "mid") return "center";
  if (anchor === "b" || anchor === "bottom") return "flex-end";
  return "flex-start";
}

function textWritingMode(body: TextBox["body"] | null | undefined): React.CSSProperties["writingMode"] {
  if (!body?.vertical) return undefined;
  if (body.vertical.includes("vert")) return "vertical-rl";
  return undefined;
}

function textAllowsOverflow(body: TextBox["body"] | null | undefined): boolean {
  return body?.vertical_overflow === "overflow" || body?.horizontal_overflow === "overflow";
}

function textOverflowStyle(body: TextBox["body"] | null | undefined): React.CSSProperties {
  if (!textAllowsOverflow(body)) return { overflow: "hidden" };
  return {
    overflowX: body?.horizontal_overflow === "overflow" ? "visible" : "hidden",
    overflowY: body?.vertical_overflow === "overflow" ? "visible" : "hidden",
  };
}

function textLineHeight(body: TextBox["body"] | null | undefined): number {
  return body?.wrap === "none" ? 1 : 1;
}

function croppedImageStyle(crop: { left: number; top: number; right: number; bottom: number } | null | undefined): React.CSSProperties {
  const left = Math.max(0, (crop?.left ?? 0) / 100000);
  const top = Math.max(0, (crop?.top ?? 0) / 100000);
  const right = Math.max(0, (crop?.right ?? 0) / 100000);
  const bottom = Math.max(0, (crop?.bottom ?? 0) / 100000);
  const visibleW = Math.max(0.01, 1 - left - right);
  const visibleH = Math.max(0.01, 1 - top - bottom);
  return {
    position: "absolute",
    left: `${(-left / visibleW) * 100}%`,
    top: `${(-top / visibleH) * 100}%`,
    width: `${(1 / visibleW) * 100}%`,
    height: `${(1 / visibleH) * 100}%`,
    objectFit: "fill",
  };
}

function strokeLinecap(cap: string | null | undefined): "butt" | "round" | "square" {
  if (cap === "rnd") return "round";
  if (cap === "sq") return "square";
  return "butt";
}

function strokeLinejoin(join: string | null | undefined): "miter" | "round" | "bevel" {
  if (join === "round" || join === "bevel") return join;
  return "miter";
}

function strokeDashArray(dash: string | null | undefined, strokeWidth = 1): string | undefined {
  if (!dash || dash === "solid") return undefined;
  const sw = Math.max(1, strokeWidth);
  const pattern = (values: number[]) => values.map((value) => Math.max(1, value * sw).toFixed(1)).join(" ");
  if (dash === "dash" || dash === "sysDash") return pattern([4, 3]);
  if (dash === "dot" || dash === "sysDot") return pattern([1, 3]);
  if (dash === "lgDash") return pattern([8, 3]);
  if (dash === "dashDot" || dash === "sysDashDot") return pattern([4, 3, 1, 3]);
  if (dash === "lgDashDot") return pattern([8, 3, 1, 3]);
  if (dash === "lgDashDotDot") return pattern([8, 3, 1, 3, 1, 3]);
  return pattern([4, 3]);
}

function guideValue(adjustments: GeometryAdjustment[] | null | undefined, name: string, fallback: number): number {
  return adjustments?.find((item) => item.name === name)?.value ?? fallback;
}

function guideFraction(adjustments: GeometryAdjustment[] | null | undefined, name: string, fallback: number): number {
  return guideValue(adjustments, name, fallback) / 100000;
}

function markerPath(type: string | null | undefined): React.ReactNode {
  switch (type) {
    case "oval":
      return <circle cx="5" cy="5" r="4" />;
    case "diamond":
      return <path d="M 5 0 L 10 5 L 5 10 L 0 5 Z" />;
    case "stealth":
      return <path d="M 0 0 L 10 5 L 0 10 L 3 5 Z" />;
    case "triangle":
    case "arrow":
    default:
      return <path d="M 0 0 L 10 5 L 0 10 Z" />;
  }
}

function isRenderableImage(image: SlidesImage | undefined): image is SlidesImage {
  return Boolean(image && ["image/png", "image/jpeg", "image/gif", "image/bmp", "image/svg+xml", "image/webp"].includes(image.mime));
}

function adjustment(shape: Shape, name: string, fallback: number): number {
  return guideValue(shape.adjustments, name, fallback);
}

function calloutPolylinePoints(shape: Shape): string | null {
  const preset = shape.preset ?? "";
  if (!preset.toLowerCase().includes("callout")) return null;
  const pairs: Array<[string, string]> = [
    ["adj2", "adj1"],
    ["adj4", "adj3"],
  ];
  if (preset.includes("2") || preset.includes("3")) {
    pairs.push(["adj6", "adj5"]);
  }
  if (preset.includes("3")) {
    pairs.push(["adj8", "adj7"]);
  }
  return pairs
    .filter(([xName, yName]) => shape.adjustments.some((item) => item.name === xName || item.name === yName))
    .map(([xName, yName]) => `${adjustment(shape, xName, 0) / 1000},${adjustment(shape, yName, 0) / 1000}`)
    .join(" ") || null;
}

function connectorPath(
  preset: string | null | undefined,
  type: ConnectorType,
  width: number,
  height: number,
  adjustments: GeometryAdjustment[] | null | undefined,
): string {
  const w = Math.max(0, width);
  const h = Math.max(0, height);
  const name = preset ?? "";

  if (type === "Bent") {
    if (name === "bentConnector2") {
      return `M 0 0 L ${w} 0 L ${w} ${h}`;
    }
    if (name === "bentConnector3") {
      const x1 = w * guideFraction(adjustments, "adj1", 50000);
      return `M 0 0 L ${x1} 0 L ${x1} ${h} L ${w} ${h}`;
    }
    if (name === "bentConnector4") {
      const x1 = w * guideFraction(adjustments, "adj1", 50000);
      const y2 = h * guideFraction(adjustments, "adj2", 50000);
      return `M 0 0 L ${x1} 0 L ${x1} ${y2} L ${w} ${y2} L ${w} ${h}`;
    }
    const midX = w * guideFraction(adjustments, "adj1", 50000);
    return `M 0 0 L ${midX} 0 L ${midX} ${h} L ${w} ${h}`;
  }

  if (type === "Curved") {
    if (name === "curvedConnector2") {
      return `M 0 0 Q ${w} 0 ${w} ${h}`;
    }
    if (name === "curvedConnector3") {
      const x2 = w * guideFraction(adjustments, "adj1", 50000);
      const x1 = x2 / 2;
      const x3 = (w + x2) / 2;
      return `M 0 0 C ${x1} 0, ${x2} ${h / 4}, ${x2} ${h / 2} C ${x2} ${h * 0.75}, ${x3} ${h}, ${w} ${h}`;
    }
    return `M 0 0 C ${w * 0.5} 0, ${w * 0.5} ${h}, ${w} ${h}`;
  }

  return `M 0 0 L ${w} ${h}`;
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
      body: {
        inset_left: null,
        inset_top: null,
        inset_right: null,
        inset_bottom: null,
        anchor: null,
        rotation: null,
        vertical: null,
        wrap: null,
        vertical_overflow: null,
        horizontal_overflow: null,
      },
      flip_h: false,
      flip_v: false,
    },
  };
}

function newShape(kind: ShapeType, color: string): SlideElement {
  return {
    Shape: {
      shape_type: kind,
      preset: null,
      adjustments: [],
      x: emu(120),
      y: emu(200),
      width: emu(220),
      height: emu(140),
      rotation: 0,
      fill: { Solid: color },
      outline: null,
      text: null,
      custom_geometry: null,
      flip_h: false,
      flip_v: false,
      effects: { shadow: null, glow: null, soft_edge_radius: null },
    },
  };
}

function shapeBorderRadius(t: ShapeType): string | number {
  if (t === "Ellipse") return "50%";
  if (t === "RoundedRectangle") return 10;
  return 4;
}

/** Return an SVG path for a preset shape geometry, relative to a 100x100 box. */
function presetShapePath(shape: Shape): string | null {
  const preset = shape.preset ?? shape.shape_type;
  switch (preset) {
    case "rect":
    case "Rectangle":
      return "M 0 0 L 100 0 L 100 100 L 0 100 Z";
    case "roundRect":
    case "RoundedRectangle":
      return "M 12 0 L 88 0 Q 100 0 100 12 L 100 88 Q 100 100 88 100 L 12 100 Q 0 100 0 88 L 0 12 Q 0 0 12 0 Z";
    case "ellipse":
    case "oval":
    case "Ellipse":
      return "M 50 0 C 77.6 0 100 22.4 100 50 C 100 77.6 77.6 100 50 100 C 22.4 100 0 77.6 0 50 C 0 22.4 22.4 0 50 0 Z";
    case "Triangle":
    case "triangle":
      return "M 50 0 L 100 100 L 0 100 Z";
    case "RightTriangle":
    case "rtTriangle":
      return "M 0 0 L 100 100 L 0 100 Z";
    case "Diamond":
    case "diamond":
      return "M 50 0 L 100 50 L 50 100 L 0 50 Z";
    case "Parallelogram":
    case "parallelogram":
      return "M 20 0 L 100 0 L 80 100 L 0 100 Z";
    case "Trapezoid":
    case "trapezoid":
      return "M 20 0 L 80 0 L 100 100 L 0 100 Z";
    case "Pentagon":
    case "pentagon":
      return "M 50 0 L 100 38 L 80 100 L 20 100 L 0 38 Z";
    case "Hexagon":
    case "hexagon":
      return "M 25 0 L 75 0 L 100 50 L 75 100 L 25 100 L 0 50 Z";
    case "Octagon":
    case "octagon":
      return "M 30 0 L 70 0 L 100 30 L 100 70 L 70 100 L 30 100 L 0 70 L 0 30 Z";
    case "Star5":
    case "star5":
      return "M 50 0 L 61 35 L 98 35 L 68 57 L 79 91 L 50 70 L 21 91 L 32 57 L 2 35 L 39 35 Z";
    case "Arrow":
    case "rightArrow":
      return "M 0 35 L 60 35 L 60 10 L 100 50 L 60 90 L 60 65 L 0 65 Z";
    case "leftArrow":
      return "M 100 35 L 40 35 L 40 10 L 0 50 L 40 90 L 40 65 L 100 65 Z";
    case "upArrow":
      return "M 35 100 L 35 40 L 10 40 L 50 0 L 90 40 L 65 40 L 65 100 Z";
    case "downArrow":
      return "M 35 0 L 35 60 L 10 60 L 50 100 L 90 60 L 65 60 L 65 0 Z";
    case "upDownArrow":
      return "M 35 40 L 10 40 L 50 0 L 90 40 L 65 40 L 65 60 L 90 60 L 50 100 L 10 60 L 35 60 Z";
    case "can":
    case "flowChartMagneticDisk":
      return "M 0 15 C 0 6.7 22.4 0 50 0 C 77.6 0 100 6.7 100 15 L 100 85 C 100 93.3 77.6 100 50 100 C 22.4 100 0 93.3 0 85 Z M 0 15 C 0 23.3 22.4 30 50 30 C 77.6 30 100 23.3 100 15";
    case "donut":
      return "M 50 0 C 77.6 0 100 22.4 100 50 C 100 77.6 77.6 100 50 100 C 22.4 100 0 77.6 0 50 C 0 22.4 22.4 0 50 0 Z M 50 28 C 37.8 28 28 37.8 28 50 C 28 62.2 37.8 72 50 72 C 62.2 72 72 62.2 72 50 C 72 37.8 62.2 28 50 28 Z";
    case "mathPlus":
      return "M 35 0 L 65 0 L 65 35 L 100 35 L 100 65 L 65 65 L 65 100 L 35 100 L 35 65 L 0 65 L 0 35 L 35 35 Z";
    case "mathMultiply":
      return "M 20 0 L 50 30 L 80 0 L 100 20 L 70 50 L 100 80 L 80 100 L 50 70 L 20 100 L 0 80 L 30 50 L 0 20 Z";
    case "rightBrace":
      return "M 0 0 C 45 0 45 22 45 30 L 45 40 C 45 48 55 50 100 50 C 55 50 45 52 45 60 L 45 70 C 45 78 45 100 0 100";
    case "leftBrace":
      return "M 100 0 C 55 0 55 22 55 30 L 55 40 C 55 48 45 50 0 50 C 45 50 55 52 55 60 L 55 70 C 55 78 55 100 100 100";
    case "leftBracket":
      return "M 100 0 L 20 0 L 20 100 L 100 100";
    case "rightBracket":
      return "M 0 0 L 80 0 L 80 100 L 0 100";
    case "arc":
      return "M 10 80 A 70 70 0 0 1 90 20";
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
  thumbnail = false,
): React.ReactNode[] {
  return group.elements.map((child, index) => {
    const transformed = transformGroupChild(child, group);
    return renderElement(doc, transformed, groupKey * 1000 + index, scale, false, () => {}, undefined, false, thumbnail);
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
      style={{
        textAlign: textAlignCss(para.align ?? defaultStyle.align ?? null),
        marginLeft: para.margin_left ? px(para.margin_left) * scale : undefined,
        textIndent: para.indent ? px(para.indent) * scale : undefined,
      }}
    >
      {para.bullet && <span style={{ marginRight: 6 }}>{para.bullet}</span>}
      {para.runs.length === 0 ? '\u00A0' : para.runs.map((run, ri) => (
        <span
          key={ri}
          style={{
            fontWeight: run.bold ? 700 : undefined,
            fontStyle: run.italic ? "italic" : undefined,
            fontSize: run.font_size ? `${run.font_size * scale}pt` : undefined,
            color: run.color ? `#${run.color}` : undefined,
            fontFamily: run.font_family ?? undefined,
            textDecoration: [run.underline ? "underline" : "", run.strike ? "line-through" : ""].filter(Boolean).join(" ") || undefined,
            verticalAlign: run.baseline ? `${run.baseline / 1000}%` : undefined,
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
  thumbnail = false,
): React.ReactNode {
  if ("TextBox" in el) {
    const t = el.TextBox;
    const bgFill = fillToCss(t.fill);
    return (
      <div
        key={`textbox-${i}`}
        className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
        style={{
          left: px(t.x) * scale,
          top: px(t.y) * scale,
          width: px(t.width) * scale,
          height: px(t.height) * scale,
          overflow: textAllowsOverflow(t.body) ? "visible" : "hidden",
          background: bgFill,
          transform: transformCss(t.rotation, t.flip_h, t.flip_v),
          boxShadow: t.outline?.color
            ? `inset 0 0 0 ${Math.max(1, t.outline.width_emu / EMU_PER_PX)}px ${colorWithAlpha(t.outline.color, t.outline.alpha)}`
            : undefined,
          display: "flex",
          flexDirection: "column",
          justifyContent: textVerticalAlign(t.body?.anchor),
          ...textBoxPadding(
            t.body,
            scale,
            px(t.width) * scale,
            px(t.height) * scale,
            0,
            (t.font_size ?? 0) * scale * PT_TO_PX * textLineHeight(t.body),
          ),
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
            writingMode: textWritingMode(t.body),
            transform: t.body?.rotation ? `rotate(${t.body.rotation / 60000}deg)` : undefined,
            lineHeight: textLineHeight(t.body),
            whiteSpace: t.body?.wrap === "none" ? "pre" : undefined,
            ...textOverflowStyle(t.body),
            padding: 0,
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
    const path = presetShapePath(s);
    const outlineColor = s.outline?.color ? colorWithAlpha(s.outline.color, s.outline.alpha) : "rgba(0,0,0,0.15)";
    const outlineWidth = Math.max(1, (s.outline?.width_emu ?? 0) / EMU_PER_PX);
    const shapeStrokeWidth = Math.max(1, outlineWidth * scale);
    const fillCss = fillToCss(s.fill) ?? "transparent";
    const svgFillId = `pbo-fill-${i}`;
    const svgFill = fillToSvgPaint(s.fill, svgFillId);
    const svgDefs = svgFillDefs(s.fill, svgFillId);
    const dashArray = strokeDashArray(s.outline?.dash, shapeStrokeWidth);
    const pixelDashArray = strokeDashArray(s.outline?.dash, outlineWidth);

    if (s.custom_geometry?.paths.length) {
      const viewW = Math.max(1, s.custom_geometry.width || s.width);
      const viewH = Math.max(1, s.custom_geometry.height || s.height);
      return (
        <div
          key={`shape-custom-${i}`}
          className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
          style={{
            left: px(s.x) * scale,
            top: px(s.y) * scale,
            width: px(s.width) * scale,
            height: px(s.height) * scale,
            transform: transformCss(s.rotation, s.flip_h, s.flip_v),
            filter: effectFilter(s.effects, scale),
          }}
          onMouseDown={onSelect}
        >
          <svg
            viewBox={`0 0 ${viewW} ${viewH}`}
            preserveAspectRatio="none"
            style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
          >
            {svgDefs}
            {s.custom_geometry.paths.map((d, pathIndex) => (
              <path
                key={pathIndex}
                d={d}
                fill={svgFill}
                stroke={s.outline ? outlineColor : "none"}
                strokeWidth={shapeStrokeWidth}
                strokeDasharray={dashArray}
                strokeLinecap={strokeLinecap(s.outline?.cap)}
                strokeLinejoin={strokeLinejoin(s.outline?.join)}
                vectorEffect="non-scaling-stroke"
              />
            ))}
          </svg>
          {innerText && (
            <span
              style={{
                position: "absolute",
                inset: 0,
                display: "grid",
                justifyItems: "stretch",
                alignItems: textVerticalAlign(innerText.body?.anchor) === "center" ? "center" : textVerticalAlign(innerText.body?.anchor) === "flex-end" ? "end" : undefined,
                ...textBoxPadding(
                  innerText.body,
                  scale,
                  px(s.width) * scale,
                  px(s.height) * scale,
                  4,
                  (innerText.font_size ?? 0) * scale * PT_TO_PX * textLineHeight(innerText.body),
                ),
                fontSize: innerText.font_size ? `${innerText.font_size * scale}pt` : undefined,
                fontWeight: innerText.bold ? 700 : undefined,
                fontStyle: innerText.italic ? "italic" : undefined,
                color: innerText.color ? `#${innerText.color}` : undefined,
                textAlign: textAlignCss(innerText.align),
                writingMode: textWritingMode(innerText.body),
                lineHeight: textLineHeight(innerText.body),
                whiteSpace: innerText.body?.wrap === "none" ? "pre" : undefined,
                ...textOverflowStyle(innerText.body),
              }}
            >
              {renderRichText(innerText.paragraphs, innerText.text, scale, { fontSize: innerText.font_size, bold: innerText.bold, italic: innerText.italic, color: innerText.color, align: innerText.align })}
            </span>
          )}
        </div>
      );
    }

    if (s.shape_type === "Line") {
      const rawW = Math.max(0, px(s.width));
      const rawH = Math.max(0, px(s.height));
      const pad = Math.max(8, outlineWidth * 2);
      const viewW = Math.max(1, rawW) + pad * 2;
      const viewH = Math.max(1, rawH) + pad * 2;
      return (
        <svg
          key={`shape-line-${i}`}
          className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
          style={{
            left: (px(s.x) - pad) * scale,
            top: (px(s.y) - pad) * scale,
            width: viewW * scale,
            height: viewH * scale,
            overflow: "hidden",
            transform: transformCss(s.rotation, s.flip_h, s.flip_v),
            filter: effectFilter(s.effects, scale),
          }}
          onMouseDown={onSelect}
          viewBox={`${-pad} ${-pad} ${viewW} ${viewH}`}
          preserveAspectRatio="none"
        >
          <line
            x1={0}
            y1={0}
            x2={rawW}
            y2={rawH}
            stroke={outlineColor}
            strokeWidth={outlineWidth}
            strokeDasharray={pixelDashArray}
            strokeLinecap={strokeLinecap(s.outline?.cap)}
            strokeLinejoin={strokeLinejoin(s.outline?.join)}
          />
        </svg>
      );
    }

    const calloutPoints = calloutPolylinePoints(s);
    if (calloutPoints) {
      return (
        <div
          key={`shape-callout-${i}`}
          className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
          style={{
            left: px(s.x) * scale,
            top: px(s.y) * scale,
            width: px(s.width) * scale,
            height: px(s.height) * scale,
            transform: transformCss(s.rotation, s.flip_h, s.flip_v),
            filter: effectFilter(s.effects, scale),
            overflow: "visible",
          }}
          onMouseDown={onSelect}
        >
          <svg
            viewBox="0 0 100 100"
            preserveAspectRatio="none"
            style={{ position: "absolute", inset: 0, width: "100%", height: "100%", overflow: "visible" }}
          >
            {svgDefs}
            <rect
              x={0}
              y={0}
              width={100}
              height={100}
              fill={svgFill}
              stroke={s.outline ? outlineColor : "none"}
              strokeWidth={shapeStrokeWidth}
              strokeDasharray={dashArray}
              strokeLinecap={strokeLinecap(s.outline?.cap)}
              strokeLinejoin={strokeLinejoin(s.outline?.join)}
              vectorEffect="non-scaling-stroke"
            />
            <polyline
              points={calloutPoints}
              fill="none"
              stroke={outlineColor}
              strokeWidth={shapeStrokeWidth}
              strokeDasharray={dashArray}
              strokeLinecap={strokeLinecap(s.outline?.cap)}
              strokeLinejoin={strokeLinejoin(s.outline?.join)}
              vectorEffect="non-scaling-stroke"
            />
          </svg>
          {innerText && (
            <span
              style={{
                position: "absolute",
                inset: 0,
                display: "grid",
                justifyItems: "stretch",
                alignItems: textVerticalAlign(innerText.body?.anchor) === "center" ? "center" : textVerticalAlign(innerText.body?.anchor) === "flex-end" ? "end" : undefined,
                ...textBoxPadding(
                  innerText.body,
                  scale,
                  px(s.width) * scale,
                  px(s.height) * scale,
                  4,
                  (innerText.font_size ?? 0) * scale * PT_TO_PX * textLineHeight(innerText.body),
                ),
                fontSize: innerText.font_size ? `${innerText.font_size * scale}pt` : undefined,
                fontWeight: innerText.bold ? 700 : undefined,
                fontStyle: innerText.italic ? "italic" : undefined,
                color: innerText.color ? `#${innerText.color}` : undefined,
                textAlign: textAlignCss(innerText.align),
                writingMode: textWritingMode(innerText.body),
                lineHeight: textLineHeight(innerText.body),
                whiteSpace: innerText.body?.wrap === "none" ? "pre" : undefined,
                ...textOverflowStyle(innerText.body),
              }}
            >
              {renderRichText(innerText.paragraphs, innerText.text, scale, { fontSize: innerText.font_size, bold: innerText.bold, italic: innerText.italic, color: innerText.color, align: innerText.align })}
            </span>
          )}
        </div>
      );
    }

    // If preset path exists, render shape body via SVG and overlay text.
    if (path) {
      return (
        <div
          key={`shape-preset-${i}`}
          className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
          style={{
            left: px(s.x) * scale,
            top: px(s.y) * scale,
            width: px(s.width) * scale,
            height: px(s.height) * scale,
            transform: transformCss(s.rotation, s.flip_h, s.flip_v),
            filter: effectFilter(s.effects, scale),
          }}
          onMouseDown={onSelect}
        >
          <svg
            viewBox="0 0 100 100"
            preserveAspectRatio="none"
            style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
          >
            {svgDefs}
            <path
              d={path}
              fill={svgFill}
              fillRule="evenodd"
              stroke={s.outline ? outlineColor : "none"}
              strokeWidth={shapeStrokeWidth}
              strokeDasharray={dashArray}
              strokeLinecap={strokeLinecap(s.outline?.cap)}
              strokeLinejoin={strokeLinejoin(s.outline?.join)}
              vectorEffect="non-scaling-stroke"
            />
          </svg>
          {innerText && (
            <span
              style={{
                position: "absolute",
                inset: 0,
                display: "grid",
                justifyItems: "stretch",
                alignItems: textVerticalAlign(innerText.body?.anchor) === "center" ? "center" : textVerticalAlign(innerText.body?.anchor) === "flex-end" ? "end" : undefined,
                ...textBoxPadding(
                  innerText.body,
                  scale,
                  px(s.width) * scale,
                  px(s.height) * scale,
                  4,
                  (innerText.font_size ?? 0) * scale * PT_TO_PX * textLineHeight(innerText.body),
                ),
                fontSize: innerText.font_size ? `${innerText.font_size * scale}pt` : undefined,
                fontWeight: innerText.bold ? 700 : undefined,
                fontStyle: innerText.italic ? "italic" : undefined,
                color: innerText.color ? `#${innerText.color}` : undefined,
                textAlign: textAlignCss(innerText.align),
                writingMode: textWritingMode(innerText.body),
                lineHeight: textLineHeight(innerText.body),
                whiteSpace: innerText.body?.wrap === "none" ? "pre" : undefined,
                ...textOverflowStyle(innerText.body),
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
        key={`shape-basic-${i}`}
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
          transform: transformCss(s.rotation, s.flip_h, s.flip_v),
          filter: effectFilter(s.effects, scale),
          display: innerText ? "flex" : undefined,
          flexDirection: innerText ? "column" : undefined,
          justifyContent: innerText ? textVerticalAlign(innerText.body?.anchor) : undefined,
          ...(innerText
            ? textBoxPadding(
                innerText.body,
                scale,
                px(s.width) * scale,
                px(s.height) * scale,
                0,
                (innerText.font_size ?? 0) * scale * PT_TO_PX * textLineHeight(innerText.body),
              )
            : {}),
          overflow: innerText && textAllowsOverflow(innerText.body) ? "visible" : "hidden",
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
              writingMode: textWritingMode(innerText.body),
              lineHeight: textLineHeight(innerText.body),
              whiteSpace: innerText.body?.wrap === "none" ? "pre" : undefined,
              ...textOverflowStyle(innerText.body),
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
    const sourceImg = findImage(doc, p.image_id);
    const img = isRenderableImage(sourceImg) ? sourceImg : undefined;
    if (p.custom_geometry?.paths.length) {
      const viewW = Math.max(1, p.custom_geometry.width || p.width);
      const viewH = Math.max(1, p.custom_geometry.height || p.height);
      const clipId = `pbo-picture-clip-${i}`;
      return (
        <div
          key={`picture-custom-${i}`}
          className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
          style={{
            left: px(p.x) * scale,
            top: px(p.y) * scale,
            width: px(p.width) * scale,
            height: px(p.height) * scale,
            transform: transformCss(p.rotation, p.flip_h, p.flip_v),
            filter: effectFilter(p.effects, scale),
            backgroundColor: img ? undefined : "#eef0f5",
            pointerEvents: thumbnail ? "none" : undefined,
          }}
          onMouseDown={onSelect}
        >
          <svg
            viewBox={`0 0 ${viewW} ${viewH}`}
            preserveAspectRatio="none"
            style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
          >
            <defs>
              <clipPath id={clipId} clipPathUnits="userSpaceOnUse">
                {p.custom_geometry.paths.map((d, pathIndex) => (
                  <path key={pathIndex} d={d} />
                ))}
              </clipPath>
            </defs>
            {img ? (
              <image
                href={`data:${img.mime};base64,${img.data_base64}`}
                x={`${(-(p.crop?.left ?? 0) / Math.max(1, 100000 - (p.crop?.left ?? 0) - (p.crop?.right ?? 0))) * viewW}`}
                y={`${(-(p.crop?.top ?? 0) / Math.max(1, 100000 - (p.crop?.top ?? 0) - (p.crop?.bottom ?? 0))) * viewH}`}
                width={`${(100000 / Math.max(1, 100000 - (p.crop?.left ?? 0) - (p.crop?.right ?? 0))) * viewW}`}
                height={`${(100000 / Math.max(1, 100000 - (p.crop?.top ?? 0) - (p.crop?.bottom ?? 0))) * viewH}`}
                preserveAspectRatio="none"
                clipPath={`url(#${clipId})`}
              />
            ) : (
              <rect width={viewW} height={viewH} fill="#eef0f5" clipPath={`url(#${clipId})`} />
            )}
          </svg>
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
              🖼 {sourceImg ? `Unsupported ${sourceImg.mime}` : (p.name ?? p.image_id)}
            </div>
          )}
        </div>
      );
    }
    return (
      <div
        key={`picture-${i}`}
        className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
        style={{
          left: px(p.x) * scale,
          top: px(p.y) * scale,
          width: px(p.width) * scale,
          height: px(p.height) * scale,
          transform: transformCss(p.rotation, p.flip_h, p.flip_v),
          filter: effectFilter(p.effects, scale),
          overflow: "hidden",
          backgroundImage: img && p.crop?.tile ? `url(data:${img.mime};base64,${img.data_base64})` : undefined,
          backgroundSize: img && p.crop?.tile ? "auto" : undefined,
          backgroundRepeat: img && p.crop?.tile ? "repeat" : undefined,
          backgroundColor: img ? undefined : "#eef0f5",
          pointerEvents: thumbnail ? "none" : undefined,
        }}
        onMouseDown={onSelect}
      >
        {img && !p.crop?.tile && (
          <img
            src={`data:${img.mime};base64,${img.data_base64}`}
            alt={p.name ?? p.image_id}
            draggable={false}
            style={croppedImageStyle(p.crop)}
          />
        )}
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
            🖼 {sourceImg ? `Unsupported ${sourceImg.mime}` : (p.name ?? p.image_id)}
          </div>
        )}
      </div>
    );
  }
  if ("Connector" in el) {
    const c = el.Connector;
    const col = c.outline?.color ? colorWithAlpha(c.outline.color, c.outline.alpha) : "#555";
    const sw = Math.max(1, (c.outline?.width_emu ?? 9525) / EMU_PER_PX);
    const dashArray = strokeDashArray(c.outline?.dash, sw);
    const hasTailArrow = c.outline?.tail_end && c.outline.tail_end !== "none";
    const hasHeadArrow = c.outline?.head_end && c.outline.head_end !== "none";
    const tailMarkerId = `arrow-tail-${thumbnail ? "thumb" : "stage"}-${i}`;
    const headMarkerId = `arrow-head-${thumbnail ? "thumb" : "stage"}-${i}`;
    const rawW = Math.max(0, px(c.width));
    const rawH = Math.max(0, px(c.height));
    const pad = Math.max(10, sw * 4);
    const viewW = Math.max(1, rawW) + pad * 2;
    const viewH = Math.max(1, rawH) + pad * 2;
    const lineProps = {
      stroke: col,
      strokeWidth: sw,
      strokeDasharray: dashArray,
      strokeLinecap: strokeLinecap(c.outline?.cap),
      strokeLinejoin: strokeLinejoin(c.outline?.join),
      markerEnd: hasTailArrow ? `url(#${tailMarkerId})` : undefined,
      markerStart: hasHeadArrow ? `url(#${headMarkerId})` : undefined,
    };
    return (
      <svg
        key={`connector-${i}`}
        className={`pbo-slide-el${selected ? " pbo-el-selected" : ""}`}
        style={{
          left: (px(c.x) - pad) * scale,
          top: (px(c.y) - pad) * scale,
          width: viewW * scale,
          height: viewH * scale,
          overflow: "hidden",
          transform: transformCss(c.rotation, c.flip_h, c.flip_v),
          pointerEvents: thumbnail ? "none" : undefined,
        }}
        onMouseDown={onSelect}
        viewBox={`${-pad} ${-pad} ${viewW} ${viewH}`}
        preserveAspectRatio="none"
      >
        {(hasTailArrow || hasHeadArrow) && (
          <defs>
            {hasTailArrow && (
              <marker id={tailMarkerId} viewBox="0 0 10 10" refX="10" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
                <g fill={col}>{markerPath(c.outline?.tail_end)}</g>
              </marker>
            )}
            {hasHeadArrow && (
              <marker id={headMarkerId} viewBox="0 0 10 10" refX="10" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
                <g fill={col}>{markerPath(c.outline?.head_end)}</g>
              </marker>
            )}
          </defs>
        )}
        <path d={connectorPath(c.preset_geometry, c.connector_type, rawW, rawH, c.adjustments)} fill="none" {...lineProps} />
      </svg>
    );
  }
  if ("Table" in el) {
    const t = el.Table;
    return (
      <div
        key={`table-${i}`}
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
      ...renderGroupChildren(doc, g, i, scale, thumbnail),
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
        renderElement(doc, el, i, scale, false, () => {}, undefined, false, true),
      )}
    </div>
  );
}

export default function Slides() {
  const nav = useNavigate();
  const stageRef = useRef<HTMLElement | null>(null);
  const [doc, setDoc] = useState<SlidesDocument>(emptyDoc());
  const [filename, setFilename] = useState<string>("Untitled.pptx");
  const [activeIdx, setActiveIdx] = useState(0);
  const [selEl, setSelEl] = useState<number | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [showNotes, setShowNotes] = useState(false);
  const [stageSize, setStageSize] = useState({ width: 960, height: 540 });

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
      const parsed = await renderPptx(bytes).catch(() => parsePptx(bytes));
      setDoc(parsed.slides.length ? parsed : emptyDoc());
      setActiveIdx(0);
      setSelEl(null);
      setShowNotes(false);
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
    setShowNotes(false);
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

  useEffect(() => {
    const stage = stageRef.current;
    if (!stage) return;
    const update = () => {
      setStageSize({
        width: Math.max(320, stage.clientWidth - 48),
        height: Math.max(240, stage.clientHeight - 48),
      });
    };
    update();
    const resizeObserver = new ResizeObserver(update);
    resizeObserver.observe(stage);
    return () => resizeObserver.disconnect();
  }, []);

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
  const scale = Math.min(
    stageSize.width / slideW,
    stageSize.height / slideH,
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
            const parsed = await renderPptx(bytes).catch(() => parsePptx(bytes));
            setDoc(parsed.slides.length ? parsed : emptyDoc());
            setActiveIdx(0);
            setSelEl(null);
            setShowNotes(false);
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
          <div className="pbo-color-swatches" aria-label="Shape fill colors">
            {SHAPE_FILL_SWATCHES.map((color) => (
              <button
                key={color}
                type="button"
                className="pbo-color-swatch"
                style={{ backgroundColor: `#${color}` }}
                onClick={() => setShapeFill(`#${color}`)}
                title={`Shape fill #${color}`}
                disabled={selEl === null}
              />
            ))}
          </div>
          <button
            className="pbo-btn"
            onClick={deleteSelected}
            disabled={selEl === null}
          >
            ✕ Delete
          </button>
        </div>
        <div className="pbo-group pbo-group-note">
          <button
            className={`pbo-btn${showNotes ? " pbo-active" : ""}`}
            onClick={() => setShowNotes((value) => !value)}
            title="Toggle speaker notes"
          >
            📝 Notes
          </button>
          <button className="pbo-btn pbo-btn-advanced" onClick={openAdvancedEditor} disabled={busy} title="Open the full OnlyOffice editor">
            ↗ Advanced Editor
          </button>
        </div>
      </div>

      {error && <div className="pbo-error">{error}</div>}

      <div className={`pbo-slides-body${showNotes ? " pbo-notes-open" : ""}`}>
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
          ref={stageRef}
          className="pbo-slides-stage"
          onMouseDown={() => setSelEl(null)}
        >
          <div
            className="pbo-slide-viewport"
            style={{ width: slideW * scale, height: slideH * scale }}
          >
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

        {showNotes && (
          <aside className="pbo-slides-notes">
            <h3>Speaker notes</h3>
            <textarea
              value={slide.notes ?? ""}
              placeholder="Add notes for slide…"
              onChange={(e) => setNotes(e.target.value)}
            />
          </aside>
        )}
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
