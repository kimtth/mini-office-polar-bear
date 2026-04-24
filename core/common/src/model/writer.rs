//! Internal model for word-processing documents (Writer).

use serde::{Deserialize, Serialize};

/// Canonical in-memory representation of a DOCX document.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WriterDocument {
    pub metadata: DocumentMetadata,
    pub body: Vec<Block>,
    /// Section properties at the end of the body (page size, margins, etc.).
    #[serde(default)]
    pub section: Option<SectionProperties>,
    /// Relationship → hyperlink / image resolution table.
    #[serde(default)]
    pub resources: WriterResources,
    /// Numbering definitions (numbering.xml).
    #[serde(default)]
    pub numbering: NumberingTable,
    /// Named style catalogue (styles.xml).
    #[serde(default)]
    pub styles: StyleCatalog,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DocumentMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub last_modified_by: Option<String>,
    pub description: Option<String>,
    pub subject: Option<String>,
    pub keywords: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
}

/// External assets referenced by the document (images, hyperlinks).
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WriterResources {
    pub images: Vec<EmbeddedImage>,
    pub hyperlinks: Vec<HyperlinkTarget>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct EmbeddedImage {
    /// Logical identifier (e.g. "image1.png") used by `InlineImage::image_id`.
    pub id: String,
    /// MIME type inferred from the extension ("image/png", "image/jpeg", ...).
    pub mime: String,
    /// Image bytes, base64-encoded for JSON payload transit.
    pub data_base64: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct HyperlinkTarget {
    pub id: String,
    pub url: String,
}

/// Block-level content inside a body or table cell.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
    PageBreak,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Paragraph {
    pub runs: Vec<RunElement>,
    pub style: Option<ParagraphStyle>,
    #[serde(default)]
    pub bookmarks: Vec<Bookmark>,
}

/// Inline content inside a paragraph. DOCX allows more than just `<w:r>`.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub enum RunElement {
    Run(Run),
    Hyperlink(Hyperlink),
    InlineImage(InlineImage),
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Run {
    pub text: String,
    pub style: Option<RunStyle>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Hyperlink {
    /// Relationship id pointing into `WriterResources::hyperlinks`.
    pub rel_id: Option<String>,
    pub anchor: Option<String>,
    pub runs: Vec<Run>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct InlineImage {
    pub image_id: String,
    pub alt: Option<String>,
    /// Width in EMU (914400 per inch).
    pub width_emu: i64,
    pub height_emu: i64,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct RunStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub font_size: Option<f32>,
    pub font_family: Option<String>,
    pub color: Option<String>,
    pub highlight: Option<String>,
    pub vertical_align: Option<VerticalAlign>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum VerticalAlign {
    Baseline,
    Superscript,
    Subscript,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ParagraphStyle {
    pub alignment: Option<Alignment>,
    pub heading_level: Option<u8>,
    pub style_id: Option<String>,
    pub indent: Option<Indent>,
    pub spacing: Option<Spacing>,
    pub list: Option<ListRef>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct Indent {
    pub left: i32,
    pub right: i32,
    pub first_line: i32,
    pub hanging: i32,
}

/// Paragraph spacing: twips for before/after, 240ths for line.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct Spacing {
    pub before: i32,
    pub after: i32,
    pub line: i32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct ListRef {
    pub num_id: i32,
    pub level: u8,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Bookmark {
    pub id: String,
    pub name: String,
}

// ── Tables ──────────────────────────────────────────

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Table {
    pub rows: Vec<TableRow>,
    pub grid_cols: Vec<i32>,
    pub style: Option<TableStyle>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableRow {
    pub cells: Vec<TableCell>,
    pub height: i32,
    pub header: bool,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableCell {
    pub content: Vec<Block>,
    pub width: i32,
    pub grid_span: u32,
    pub v_merge: VMerge,
    pub borders: Option<CellBorders>,
    pub shading: Option<String>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum VMerge {
    #[default]
    None,
    Restart,
    Continue,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CellBorders {
    pub top: Option<BorderLine>,
    pub bottom: Option<BorderLine>,
    pub left: Option<BorderLine>,
    pub right: Option<BorderLine>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BorderLine {
    pub style: String,
    pub size: u32,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct TableStyle {
    pub style_id: Option<String>,
    pub borders: Option<CellBorders>,
}

// ── Sections ────────────────────────────────────────

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SectionProperties {
    pub page_size: Option<PageSize>,
    pub margins: Option<PageMargins>,
    pub columns: Option<u32>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct PageSize {
    pub width: i32,
    pub height: i32,
    pub landscape: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct PageMargins {
    pub top: i32,
    pub bottom: i32,
    pub left: i32,
    pub right: i32,
    pub header: i32,
    pub footer: i32,
    pub gutter: i32,
}

// ── Numbering ───────────────────────────────────────

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumberingTable {
    pub instances: Vec<NumInstance>,
    pub abstract_nums: Vec<AbstractNum>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumInstance {
    pub num_id: i32,
    pub abstract_num_id: i32,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AbstractNum {
    pub abstract_num_id: i32,
    pub levels: Vec<NumLevel>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumLevel {
    pub ilvl: u8,
    pub format: ListFormat,
    pub text: String,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum ListFormat {
    #[default]
    Bullet,
    Decimal,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
    Ordinal,
    Other,
}

// ── Style catalogue ─────────────────────────────────

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct StyleCatalog {
    pub styles: Vec<StyleDefinition>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct StyleDefinition {
    pub style_id: String,
    pub name: Option<String>,
    pub style_type: StyleType,
    pub based_on: Option<String>,
    pub paragraph: Option<ParagraphStyle>,
    pub run: Option<RunStyle>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub enum StyleType {
    #[default]
    Paragraph,
    Character,
    Table,
    Numbering,
}

impl NumberingTable {
    /// Resolve a list marker for a `(num_id, level)` reference, if any.
    pub fn resolve(&self, list_ref: ListRef) -> Option<&NumLevel> {
        let abstract_id = self
            .instances
            .iter()
            .find(|i| i.num_id == list_ref.num_id)?
            .abstract_num_id;
        self.abstract_nums
            .iter()
            .find(|a| a.abstract_num_id == abstract_id)?
            .levels
            .iter()
            .find(|lvl| lvl.ilvl == list_ref.level)
    }
}