//! Internal model for presentation documents (Slides).

use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlidesDocument {
    pub metadata: PresentationMetadata,
    pub slides: Vec<Slide>,
    /// Slide width in EMU (914400 = 1 inch). 10" default = 9144000.
    pub slide_width: f64,
    pub slide_height: f64,
    /// Reusable slide layouts (resolved from `ppt/slideLayouts/`).
    #[serde(default)]
    pub layouts: Vec<SlideLayout>,
    /// Slide masters (`ppt/slideMasters/`). Layouts reference these indirectly.
    #[serde(default)]
    pub masters: Vec<SlideMaster>,
    /// Optional theme — the first referenced theme is loaded here.
    #[serde(default)]
    pub theme: Option<Theme>,
    /// Embedded images keyed by logical id.
    #[serde(default)]
    pub images: Vec<SlidesImage>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PresentationMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub last_modified_by: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Slide {
    pub elements: Vec<SlideElement>,
    pub notes: Option<String>,
    /// Index into `SlidesDocument::layouts`, if resolved.
    pub layout_idx: Option<usize>,
    /// Background fill override.
    pub background: Option<Fill>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlideLayout {
    pub name: String,
    pub elements: Vec<SlideElement>,
    pub master_idx: Option<usize>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlideMaster {
    pub name: String,
    pub elements: Vec<SlideElement>,
    /// Theme referenced by this master, if available.
    #[serde(default)]
    pub theme: Option<Theme>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Theme {
    pub name: String,
    pub colors: Vec<ThemeColor>,
    /// Theme format scheme fill styles (`a:fmtScheme/a:fillStyleLst`) used by `p:style/a:fillRef`.
    #[serde(default)]
    pub fill_styles: Vec<ThemeFillStyle>,
    pub major_font: Option<String>,
    pub minor_font: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ThemeColor {
    /// `dk1`, `lt1`, `accent1`, …
    pub role: String,
    pub value: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ThemeFillStyle {
    /// 1-based index in `fillStyleLst`, matching OOXML style references.
    pub idx: u32,
    #[serde(default)]
    pub is_gradient: bool,
    #[serde(default)]
    pub stops: Vec<ThemeFillStop>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ThemeFillStop {
    /// Gradient stop position; solid fills use 0.
    pub position: u32,
    /// Direct RGB value, when the style uses `srgbClr`.
    #[serde(default)]
    pub color: Option<String>,
    /// Theme role (`phClr`, `accent1`, …), when the style uses `schemeClr`.
    #[serde(default)]
    pub role: Option<String>,
    #[serde(default)]
    pub alpha: Option<i64>,
    #[serde(default)]
    pub modifiers: Vec<ThemeColorModifier>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ThemeColorModifier {
    pub name: String,
    pub value: i64,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlidesImage {
    pub id: String,
    pub mime: String,
    pub data_base64: String,
}

#[allow(clippy::large_enum_variant)]
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum SlideElement {
    TextBox(TextBox),
    Picture(Picture),
    Shape(Shape),
    Table(SlideTable),
    Connector(Connector),
    /// Nested group shape (allows recursion).
    Group(GroupShape),
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextBox {
    pub text: String,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub font_size: Option<f32>,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<String>,
    pub align: Option<TextAlign>,
    /// Rotation in 60000ths of a degree (OOXML convention).
    pub rotation: i32,
    /// Placeholder slot (title, body, ctrTitle, …) when this textbox comes
    /// from a layout/master placeholder.
    pub placeholder: Option<String>,
    /// Placeholder idx attribute (0, 1, 2, …) for disambiguation.
    #[serde(default)]
    pub placeholder_idx: Option<String>,
    /// Rich text paragraphs with per-run styling.
    #[serde(default)]
    pub paragraphs: Vec<TextParagraph>,
    /// Fill for placeholder shapes that carry fill info.
    #[serde(default)]
    pub fill: Option<Fill>,
    /// Outline for placeholder shapes.
    #[serde(default)]
    pub outline: Option<Outline>,
    /// DrawingML text body properties from `<a:bodyPr>`.
    #[serde(default)]
    pub body: TextBodyProperties,
    /// Horizontal flip from `<a:xfrm flipH="1">`.
    #[serde(default)]
    pub flip_h: bool,
    /// Vertical flip from `<a:xfrm flipV="1">`.
    #[serde(default)]
    pub flip_v: bool,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextParagraph {
    pub runs: Vec<TextRun>,
    pub align: Option<TextAlign>,
    /// Bullet character or numbering marker when parsed from `<a:bu*>`.
    #[serde(default)]
    pub bullet: Option<String>,
    /// Paragraph left margin in EMUs (`marL`).
    #[serde(default)]
    pub margin_left: Option<f64>,
    /// First-line indent in EMUs (`indent`).
    #[serde(default)]
    pub indent: Option<f64>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextRun {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub font_size: Option<f32>,
    pub color: Option<String>,
    #[serde(default)]
    pub font_family: Option<String>,
    #[serde(default)]
    pub underline: bool,
    #[serde(default)]
    pub strike: bool,
    /// Baseline offset in OOXML percentage units.
    #[serde(default)]
    pub baseline: Option<i32>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextBodyProperties {
    /// Text box left inset in EMUs (`lIns`).
    #[serde(default)]
    pub inset_left: Option<f64>,
    /// Text box top inset in EMUs (`tIns`).
    #[serde(default)]
    pub inset_top: Option<f64>,
    /// Text box right inset in EMUs (`rIns`).
    #[serde(default)]
    pub inset_right: Option<f64>,
    /// Text box bottom inset in EMUs (`bIns`).
    #[serde(default)]
    pub inset_bottom: Option<f64>,
    /// Vertical anchor (`t`, `ctr`, `b`, …).
    #[serde(default)]
    pub anchor: Option<String>,
    /// Body text rotation in 60000ths of a degree.
    #[serde(default)]
    pub rotation: Option<i32>,
    /// Vertical text mode (`vert`, `vert270`, `wordArtVert`, …).
    #[serde(default)]
    pub vertical: Option<String>,
    /// Whether wrapped text is enabled (`wrap`).
    #[serde(default)]
    pub wrap: Option<String>,
    /// Vertical text overflow behavior (`overflow`, `clip`, `ellipsis`).
    #[serde(default)]
    pub vertical_overflow: Option<String>,
    /// Horizontal text overflow behavior (`overflow`, `clip`).
    #[serde(default)]
    pub horizontal_overflow: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum TextAlign {
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Picture {
    /// Logical image id, resolved via `SlidesDocument::images`.
    pub image_id: String,
    pub name: Option<String>,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub rotation: i32,
    /// Parsed DrawingML custom geometry mask, when the picture is clipped by `<a:custGeom>`.
    #[serde(default)]
    pub custom_geometry: Option<CustomGeometry>,
    /// Source rectangle crop from `<a:srcRect>` in 1/1000 percent units.
    #[serde(default)]
    pub crop: PictureCrop,
    /// Horizontal flip from `<a:xfrm flipH="1">`.
    #[serde(default)]
    pub flip_h: bool,
    /// Vertical flip from `<a:xfrm flipV="1">`.
    #[serde(default)]
    pub flip_v: bool,
    /// Simplified visual effects parsed from `<a:effectLst>` / `<a:effectDag>`.
    #[serde(default)]
    pub effects: VisualEffects,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PictureCrop {
    #[serde(default)]
    pub left: f64,
    #[serde(default)]
    pub top: f64,
    #[serde(default)]
    pub right: f64,
    #[serde(default)]
    pub bottom: f64,
    #[serde(default)]
    pub tile: bool,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Shape {
    pub shape_type: ShapeType,
    /// Original DrawingML preset geometry name (`borderCallout2`, `rect`, …).
    #[serde(default)]
    pub preset: Option<String>,
    /// Preset geometry guide values from `<a:prstGeom><a:avLst><a:gd .../>`.
    #[serde(default)]
    pub adjustments: Vec<GeometryAdjustment>,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub rotation: i32,
    pub fill: Option<Fill>,
    pub outline: Option<Outline>,
    /// Text inside the shape (prstTxWarp or body text).
    pub text: Option<TextBox>,
    /// Parsed DrawingML custom geometry paths, when the shape uses `<a:custGeom>`.
    #[serde(default)]
    pub custom_geometry: Option<CustomGeometry>,
    /// Horizontal flip from `<a:xfrm flipH="1">`.
    #[serde(default)]
    pub flip_h: bool,
    /// Vertical flip from `<a:xfrm flipV="1">`.
    #[serde(default)]
    pub flip_v: bool,
    /// Simplified visual effects parsed from `<a:effectLst>` / `<a:effectDag>`.
    #[serde(default)]
    pub effects: VisualEffects,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CustomGeometry {
    /// Native geometry coordinate width from `<a:path w="...">`.
    pub width: f64,
    /// Native geometry coordinate height from `<a:path h="...">`.
    pub height: f64,
    /// SVG path data translated from DrawingML path commands.
    pub paths: Vec<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GeometryAdjustment {
    pub name: String,
    pub formula: String,
    #[serde(default)]
    pub value: Option<f64>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum ShapeType {
    #[default]
    Rectangle,
    RoundedRectangle,
    Ellipse,
    Triangle,
    RightTriangle,
    Diamond,
    Parallelogram,
    Trapezoid,
    Pentagon,
    Hexagon,
    Octagon,
    Star5,
    Arrow,
    Line,
    Callout,
    /// Any preset geometry not individually enumerated.
    Other,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Connector {
    pub connector_type: ConnectorType,
    /// Original DrawingML preset geometry (for example `straightConnector1`, `bentConnector2`).
    #[serde(default)]
    pub preset_geometry: Option<String>,
    /// Preset geometry guide values from `<a:prstGeom><a:avLst><a:gd .../>`.
    #[serde(default)]
    pub adjustments: Vec<GeometryAdjustment>,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub rotation: i32,
    pub outline: Option<Outline>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum ConnectorType {
    #[default]
    Straight,
    Bent,
    Curved,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GroupShape {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    /// Child coordinate space origin x (from `<a:chOff>`).
    #[serde(default)]
    pub ch_off_x: f64,
    /// Child coordinate space origin y.
    #[serde(default)]
    pub ch_off_y: f64,
    /// Child coordinate space width (from `<a:chExt>`).
    #[serde(default)]
    pub ch_ext_cx: f64,
    /// Child coordinate space height.
    #[serde(default)]
    pub ch_ext_cy: f64,
    pub elements: Vec<SlideElement>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlideTable {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub rows: Vec<SlideTableRow>,
    pub columns: Vec<f64>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlideTableRow {
    pub height: f64,
    pub cells: Vec<SlideTableCell>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SlideTableCell {
    pub text: String,
    pub col_span: u32,
    pub row_span: u32,
    pub fill: Option<Fill>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Fill {
    Solid(String),
    SolidAlpha(SolidFill),
    Gradient(Vec<GradientStop>),
    Pattern(PatternFill),
    None,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SolidFill {
    pub color: String,
    /// Alpha 0..=100000 from DrawingML `<a:alpha>`.
    #[serde(default)]
    pub alpha: Option<i64>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PatternFill {
    #[serde(default)]
    pub preset: Option<String>,
    #[serde(default)]
    pub foreground: Option<String>,
    #[serde(default)]
    pub background: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GradientStop {
    /// Position 0..=100000 (OOXML convention: 0..100 in %%).
    pub position: u32,
    pub color: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Outline {
    pub width_emu: i64,
    pub color: Option<String>,
    /// Alpha 0..=100000 from DrawingML line color `<a:alpha>`.
    #[serde(default)]
    pub alpha: Option<i64>,
    pub dash: Option<String>,
    /// Head-end type (e.g. "triangle", "arrow", "none").
    #[serde(default)]
    pub head_end: Option<String>,
    /// Tail-end type.
    #[serde(default)]
    pub tail_end: Option<String>,
    /// Line cap (`flat`, `rnd`, `sq`) from `<a:ln cap="...">`.
    #[serde(default)]
    pub cap: Option<String>,
    /// Line join (`round`, `bevel`, `miter`).
    #[serde(default)]
    pub join: Option<String>,
    /// Compound line type (`sng`, `dbl`, `thickThin`, …).
    #[serde(default)]
    pub compound: Option<String>,
    /// Pen alignment (`ctr`, `in`).
    #[serde(default)]
    pub align: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct VisualEffects {
    #[serde(default)]
    pub shadow: Option<ShadowEffect>,
    #[serde(default)]
    pub glow: Option<GlowEffect>,
    #[serde(default)]
    pub soft_edge_radius: Option<f64>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ShadowEffect {
    /// Blur radius in EMUs.
    #[serde(default)]
    pub blur_radius: f64,
    /// Distance in EMUs.
    #[serde(default)]
    pub distance: f64,
    /// Direction in 60000ths of a degree.
    #[serde(default)]
    pub direction: i32,
    #[serde(default)]
    pub color: Option<String>,
    /// Alpha 0..=100000.
    #[serde(default)]
    pub alpha: Option<i64>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct GlowEffect {
    /// Glow radius in EMUs.
    #[serde(default)]
    pub radius: f64,
    #[serde(default)]
    pub color: Option<String>,
    /// Alpha 0..=100000.
    #[serde(default)]
    pub alpha: Option<i64>,
}
