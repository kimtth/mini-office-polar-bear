//! Internal model for spreadsheet documents (Calc).

use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CalcDocument {
    pub metadata: SpreadsheetMetadata,
    pub sheets: Vec<Sheet>,
    /// Workbook-level style catalogue (styles.xml).
    #[serde(default)]
    pub styles: WorkbookStyles,
    /// Workbook-level defined names (`<definedNames>`).
    #[serde(default)]
    pub defined_names: Vec<DefinedName>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SpreadsheetMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub last_modified_by: Option<String>,
    pub created: Option<String>,
    pub modified: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Sheet {
    pub name: String,
    /// Row index (1-based) → column index (1-based) → cell.
    pub cells: BTreeMap<u32, BTreeMap<u32, Cell>>,
    /// Column dimension ranges ordered by `min` (inclusive, 1-based).
    #[serde(default)]
    pub columns: Vec<ColumnDim>,
    /// Row index → height (points).
    #[serde(default)]
    pub row_heights: BTreeMap<u32, f64>,
    /// Legacy per-column width map, kept for simple cases.
    #[serde(default)]
    pub column_widths: BTreeMap<u32, f64>,
    /// Merged cell regions.
    #[serde(default)]
    pub merges: Vec<MergedRange>,
    /// Hyperlinks keyed by cell reference ("A1").
    #[serde(default)]
    pub hyperlinks: Vec<SheetHyperlink>,
    #[serde(default)]
    pub view: Option<SheetView>,
    #[serde(default)]
    pub tab_color: Option<String>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Serialize, Deserialize)]
pub struct ColumnDim {
    pub min: u32,
    pub max: u32,
    pub width: f64,
    pub custom_width: bool,
    pub hidden: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct MergedRange {
    pub row_start: u32,
    pub col_start: u32,
    pub row_end: u32,
    pub col_end: u32,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SheetHyperlink {
    /// A1-style reference ("A1", "B2:C4").
    pub cell_ref: String,
    pub target: String,
    pub display: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct SheetView {
    pub freeze_row: u32,
    pub freeze_col: u32,
    pub show_grid_lines: bool,
    pub zoom_scale: u32,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Cell {
    pub value: CellValue,
    pub formula: Option<String>,
    /// Index into `WorkbookStyles::cell_xfs`.
    pub style_id: Option<u32>,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub enum CellValue {
    #[default]
    Empty,
    Text(String),
    Number(f64),
    Boolean(bool),
    /// ISO-8601 date string; Excel stores dates as serial numbers, but we
    /// normalize on the way in when number format matches.
    Date(String),
    Error(String),
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct DefinedName {
    pub name: String,
    pub formula: String,
    /// Sheet-local scope (sheet index) or `None` for workbook-global.
    pub local_sheet_id: Option<u32>,
}

// ── styles.xml catalogue ─────────────────────────────

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WorkbookStyles {
    pub num_fmts: Vec<NumFmt>,
    pub fonts: Vec<CellFont>,
    pub fills: Vec<CellFill>,
    pub borders: Vec<CellBorder>,
    pub cell_xfs: Vec<CellXf>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct NumFmt {
    pub id: u32,
    pub format_code: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CellFont {
    pub size: Option<f32>,
    pub name: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CellFill {
    pub pattern_type: String,
    pub fg_color: Option<String>,
    pub bg_color: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct CellBorder {
    pub top: Option<BorderEdge>,
    pub bottom: Option<BorderEdge>,
    pub left: Option<BorderEdge>,
    pub right: Option<BorderEdge>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct BorderEdge {
    pub style: String,
    pub color: Option<String>,
}

/// Per-cell XF (eXtended Format) — indices into the parallel arrays.
#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize)]
pub struct CellXf {
    pub num_fmt_id: Option<u32>,
    pub font_id: Option<u32>,
    pub fill_id: Option<u32>,
    pub border_id: Option<u32>,
    pub alignment: Option<CellAlignment>,
    pub apply_number_format: bool,
    pub apply_font: bool,
    pub apply_fill: bool,
    pub apply_border: bool,
    pub apply_alignment: bool,
}

#[derive(Debug, Clone, Copy, Default, Serialize, Deserialize)]
pub struct CellAlignment {
    pub horizontal: Option<HAlign>,
    pub vertical: Option<VAlign>,
    pub wrap_text: bool,
    pub text_rotation: i32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum HAlign {
    Left,
    Center,
    Right,
    Justify,
    Fill,
    General,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum VAlign {
    Top,
    Middle,
    Bottom,
}

// ── Built-in number format resolver ──────────────────

/// Return the format code for a built-in numFmtId, per ECMA-376 §18.8.30.
pub fn builtin_num_fmt(id: u32) -> Option<&'static str> {
    Some(match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

impl WorkbookStyles {
    /// Resolve the effective number format for a given style index.
    pub fn number_format_for(&self, xf_idx: u32) -> Option<&str> {
        let xf = self.cell_xfs.get(xf_idx as usize)?;
        let nfid = xf.num_fmt_id?;
        if let Some(s) = builtin_num_fmt(nfid) {
            return Some(s);
        }
        self.num_fmts
            .iter()
            .find(|n| n.id == nfid)
            .map(|n| n.format_code.as_str())
    }
}
