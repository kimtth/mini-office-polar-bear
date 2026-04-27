// Type definitions mirroring the Rust internal document models
// from core/common/src/model/*.rs.
//
// Enum variants are serialized by serde as externally-tagged single-key
// objects (e.g. `{ "Run": {...} }`). All fields mirror the Rust field
// names exactly so unchanged JSON round-trips through the UI.

// ── Writer ──────────────────────────────────────────

export type Alignment = "Left" | "Center" | "Right" | "Justify";
export type VerticalAlign = "Baseline" | "Superscript" | "Subscript";

export interface RunStyle {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  font_size: number | null;
  font_family: string | null;
  color: string | null;
  highlight: string | null;
  vertical_align: VerticalAlign | null;
}

export interface Run {
  text: string;
  style: RunStyle | null;
}

export interface Hyperlink {
  rel_id: string | null;
  anchor: string | null;
  runs: Run[];
}

export interface InlineImage {
  image_id: string;
  alt: string | null;
  width_emu: number;
  height_emu: number;
}

export type RunElement =
  | { Run: Run }
  | { Hyperlink: Hyperlink }
  | { InlineImage: InlineImage };

export interface Indent {
  left: number;
  right: number;
  first_line: number;
  hanging: number;
}

export interface Spacing {
  before: number;
  after: number;
  line: number;
}

export interface ListRef {
  num_id: number;
  level: number;
}

export interface ParagraphStyle {
  alignment: Alignment | null;
  heading_level: number | null;
  style_id: string | null;
  indent: Indent | null;
  spacing: Spacing | null;
  list: ListRef | null;
}

export interface Bookmark {
  id: string;
  name: string;
}

export interface Paragraph {
  runs: RunElement[];
  style: ParagraphStyle | null;
  bookmarks: Bookmark[];
}

export type VMerge = "None" | "Restart" | "Continue";

export interface BorderLine {
  style: string;
  size: number;
  color: string | null;
}

export interface CellBorders {
  top: BorderLine | null;
  bottom: BorderLine | null;
  left: BorderLine | null;
  right: BorderLine | null;
}

export interface TableCell {
  content: Block[];
  width: number;
  grid_span: number;
  v_merge: VMerge;
  borders: CellBorders | null;
  shading: string | null;
}

export interface TableRow {
  cells: TableCell[];
  height: number;
  header: boolean;
}

export interface TableStyle {
  style_id: string | null;
  borders: CellBorders | null;
}

export interface Table {
  rows: TableRow[];
  grid_cols: number[];
  style: TableStyle | null;
}

export type Block =
  | { Paragraph: Paragraph }
  | { Table: Table }
  | "PageBreak";

export interface PageSize {
  width: number;
  height: number;
  landscape: boolean;
}

export interface PageMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
  header: number;
  footer: number;
  gutter: number;
}

export interface SectionProperties {
  page_size: PageSize | null;
  margins: PageMargins | null;
  columns: number | null;
}

export interface EmbeddedImage {
  id: string;
  mime: string;
  data_base64: string;
}

export interface HyperlinkTarget {
  id: string;
  url: string;
}

export interface WriterResources {
  images: EmbeddedImage[];
  hyperlinks: HyperlinkTarget[];
}

export interface NumInstance {
  num_id: number;
  abstract_num_id: number;
}

export type ListFormat =
  | "Bullet"
  | "Decimal"
  | "LowerLetter"
  | "UpperLetter"
  | "LowerRoman"
  | "UpperRoman"
  | "Ordinal"
  | "Other";

export interface NumLevel {
  ilvl: number;
  format: ListFormat;
  text: string;
}

export interface AbstractNum {
  abstract_num_id: number;
  levels: NumLevel[];
}

export interface NumberingTable {
  instances: NumInstance[];
  abstract_nums: AbstractNum[];
}

export type StyleType = "Paragraph" | "Character" | "Table" | "Numbering";

export interface StyleDefinition {
  style_id: string;
  name: string | null;
  style_type: StyleType;
  based_on: string | null;
  paragraph: ParagraphStyle | null;
  run: RunStyle | null;
}

export interface StyleCatalog {
  styles: StyleDefinition[];
}

export interface DocumentMetadata {
  title: string | null;
  author: string | null;
  last_modified_by: string | null;
  description: string | null;
  subject: string | null;
  keywords: string | null;
  created: string | null;
  modified: string | null;
}

export interface WriterDocument {
  metadata: DocumentMetadata;
  body: Block[];
  section: SectionProperties | null;
  resources: WriterResources;
  numbering: NumberingTable;
  styles: StyleCatalog;
}

// ── Calc ────────────────────────────────────────────

export type CellValue =
  | "Empty"
  | { Text: string }
  | { Number: number }
  | { Boolean: boolean }
  | { Date: string }
  | { Error: string };

export interface Cell {
  value: CellValue;
  formula: string | null;
  style_id: number | null;
}

export interface ColumnDim {
  min: number;
  max: number;
  width: number;
  custom_width: boolean;
  hidden: boolean;
}

export interface MergedRange {
  row_start: number;
  col_start: number;
  row_end: number;
  col_end: number;
}

export interface SheetHyperlink {
  cell_ref: string;
  target: string;
  display: string | null;
}

export interface SheetView {
  freeze_row: number;
  freeze_col: number;
  show_grid_lines: boolean;
  zoom_scale: number;
}

export interface Sheet {
  name: string;
  /** Row index (1-based) -> Column index (1-based) -> Cell */
  cells: Record<string, Record<string, Cell>>;
  columns: ColumnDim[];
  row_heights: Record<string, number>;
  column_widths: Record<string, number>;
  merges: MergedRange[];
  hyperlinks: SheetHyperlink[];
  view: SheetView | null;
  tab_color: string | null;
}

export interface SpreadsheetMetadata {
  title: string | null;
  author: string | null;
  last_modified_by: string | null;
  created: string | null;
  modified: string | null;
}

export interface NumFmt {
  id: number;
  format_code: string;
}

export interface CellFont {
  size: number | null;
  name: string | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: string | null;
}

export interface CellFill {
  pattern_type: string;
  fg_color: string | null;
  bg_color: string | null;
}

export interface BorderEdge {
  style: string;
  color: string | null;
}

export interface CellBorder {
  top: BorderEdge | null;
  bottom: BorderEdge | null;
  left: BorderEdge | null;
  right: BorderEdge | null;
}

export type HAlign = "Left" | "Center" | "Right" | "Justify" | "Fill" | "General";
export type VAlign = "Top" | "Middle" | "Bottom";

export interface CellAlignment {
  horizontal: HAlign | null;
  vertical: VAlign | null;
  wrap_text: boolean;
  text_rotation: number;
}

export interface CellXf {
  num_fmt_id: number | null;
  font_id: number | null;
  fill_id: number | null;
  border_id: number | null;
  alignment: CellAlignment | null;
  apply_number_format: boolean;
  apply_font: boolean;
  apply_fill: boolean;
  apply_border: boolean;
  apply_alignment: boolean;
}

export interface WorkbookStyles {
  num_fmts: NumFmt[];
  fonts: CellFont[];
  fills: CellFill[];
  borders: CellBorder[];
  cell_xfs: CellXf[];
}

export interface DefinedName {
  name: string;
  formula: string;
  local_sheet_id: number | null;
}

export interface CalcDocument {
  metadata: SpreadsheetMetadata;
  sheets: Sheet[];
  styles: WorkbookStyles;
  defined_names: DefinedName[];
}

// ── Slides ──────────────────────────────────────────

export type TextAlign = "Left" | "Center" | "Right" | "Justify";

export type Fill =
  | { Solid: string }
  | { SolidAlpha: SolidFill }
  | { Gradient: GradientStop[] }
  | { Pattern: PatternFill }
  | "None";

export interface SolidFill {
  color: string;
  alpha: number | null;
}

export interface GradientStop {
  position: number;
  color: string;
}

export interface Outline {
  width_emu: number;
  color: string | null;
  alpha: number | null;
  dash: string | null;
  head_end: string | null;
  tail_end: string | null;
  cap: string | null;
  join: string | null;
  compound: string | null;
  align: string | null;
}

export interface TextRun {
  text: string;
  bold: boolean;
  italic: boolean;
  font_size: number | null;
  color: string | null;
  font_family: string | null;
  underline: boolean;
  strike: boolean;
  baseline: number | null;
}

export interface TextParagraph {
  runs: TextRun[];
  align: TextAlign | null;
  bullet: string | null;
  margin_left: number | null;
  indent: number | null;
}

export interface TextBodyProperties {
  inset_left: number | null;
  inset_top: number | null;
  inset_right: number | null;
  inset_bottom: number | null;
  anchor: string | null;
  rotation: number | null;
  vertical: string | null;
  wrap: string | null;
  vertical_overflow: string | null;
  horizontal_overflow: string | null;
}

export interface TextBox {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  font_size: number | null;
  bold: boolean;
  italic: boolean;
  color: string | null;
  align: TextAlign | null;
  rotation: number;
  placeholder: string | null;
  placeholder_idx: string | null;
  paragraphs: TextParagraph[];
  fill: Fill | null;
  outline: Outline | null;
  body: TextBodyProperties;
  flip_h: boolean;
  flip_v: boolean;
}

export interface Picture {
  image_id: string;
  name: string | null;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  custom_geometry: CustomGeometry | null;
  crop: PictureCrop;
  flip_h: boolean;
  flip_v: boolean;
  effects: VisualEffects;
}

export interface PictureCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
  tile: boolean;
}

export interface CustomGeometry {
  width: number;
  height: number;
  paths: string[];
}

export interface GeometryAdjustment {
  name: string;
  formula: string;
  value: number | null;
}

export type ShapeType =
  | "Rectangle"
  | "RoundedRectangle"
  | "Ellipse"
  | "Triangle"
  | "RightTriangle"
  | "Diamond"
  | "Parallelogram"
  | "Trapezoid"
  | "Pentagon"
  | "Hexagon"
  | "Octagon"
  | "Star5"
  | "Arrow"
  | "Line"
  | "Callout"
  | "Other";

export interface Shape {
  shape_type: ShapeType;
  preset: string | null;
  adjustments: GeometryAdjustment[];
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  fill: Fill | null;
  outline: Outline | null;
  text: TextBox | null;
  custom_geometry: CustomGeometry | null;
  flip_h: boolean;
  flip_v: boolean;
  effects: VisualEffects;
}

export type ConnectorType = "Straight" | "Bent" | "Curved";

export interface Connector {
  connector_type: ConnectorType;
  preset_geometry?: string | null;
  adjustments?: GeometryAdjustment[];
  x: number;
  y: number;
  width: number;
  height: number;
  flip_h: boolean;
  flip_v: boolean;
  rotation: number;
  outline: Outline | null;
}

export interface SlideTableCell {
  text: string;
  col_span: number;
  row_span: number;
  fill: Fill | null;
}

export interface PatternFill {
  preset: string | null;
  foreground: string | null;
  background: string | null;
}

export interface VisualEffects {
  shadow: ShadowEffect | null;
  glow: GlowEffect | null;
  soft_edge_radius: number | null;
}

export interface ShadowEffect {
  blur_radius: number;
  distance: number;
  direction: number;
  color: string | null;
  alpha: number | null;
}

export interface GlowEffect {
  radius: number;
  color: string | null;
  alpha: number | null;
}

export interface SlideTableRow {
  height: number;
  cells: SlideTableCell[];
}

export interface SlideTable {
  x: number;
  y: number;
  width: number;
  height: number;
  rows: SlideTableRow[];
  columns: number[];
}

export interface GroupShape {
  x: number;
  y: number;
  width: number;
  height: number;
  ch_off_x: number;
  ch_off_y: number;
  ch_ext_cx: number;
  ch_ext_cy: number;
  elements: SlideElement[];
}

export type SlideElement =
  | { TextBox: TextBox }
  | { Picture: Picture }
  | { Shape: Shape }
  | { Table: SlideTable }
  | { Connector: Connector }
  | { Group: GroupShape };

export interface Slide {
  elements: SlideElement[];
  notes: string | null;
  layout_idx: number | null;
  background: Fill | null;
}

export interface SlideLayout {
  name: string;
  elements: SlideElement[];
  master_idx: number | null;
}

export interface SlideMaster {
  name: string;
  elements: SlideElement[];
  theme: Theme | null;
}

export interface ThemeColor {
  role: string;
  value: string;
}

export interface ThemeColorModifier {
  name: string;
  value: number;
}

export interface ThemeFillStop {
  position: number;
  color?: string | null;
  role?: string | null;
  alpha?: number | null;
  modifiers: ThemeColorModifier[];
}

export interface ThemeFillStyle {
  idx: number;
  is_gradient: boolean;
  stops: ThemeFillStop[];
}

export interface Theme {
  name: string;
  colors: ThemeColor[];
  fill_styles: ThemeFillStyle[];
  major_font: string | null;
  minor_font: string | null;
}

export interface SlidesImage {
  id: string;
  mime: string;
  data_base64: string;
}

export interface PresentationMetadata {
  title: string | null;
  author: string | null;
  last_modified_by: string | null;
  created: string | null;
  modified: string | null;
}

export interface SlidesDocument {
  metadata: PresentationMetadata;
  slides: Slide[];
  slide_width: number;
  slide_height: number;
  layouts: SlideLayout[];
  masters: SlideMaster[];
  theme: Theme | null;
  images: SlidesImage[];
}
