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
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Theme {
    pub name: String,
    pub colors: Vec<ThemeColor>,
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
pub struct SlidesImage {
    pub id: String,
    pub mime: String,
    pub data_base64: String,
}

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
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextParagraph {
    pub runs: Vec<TextRun>,
    pub align: Option<TextAlign>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TextRun {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub font_size: Option<f32>,
    pub color: Option<String>,
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
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Shape {
    pub shape_type: ShapeType,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub rotation: i32,
    pub fill: Option<Fill>,
    pub outline: Option<Outline>,
    /// Text inside the shape (prstTxWarp or body text).
    pub text: Option<TextBox>,
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
    Gradient(Vec<GradientStop>),
    None,
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
    pub dash: Option<String>,
    /// Head-end type (e.g. "triangle", "arrow", "none").
    #[serde(default)]
    pub head_end: Option<String>,
    /// Tail-end type.
    #[serde(default)]
    pub tail_end: Option<String>,
}
