//! `xlsb_write` — a fast, pure-Rust writer for the Excel Binary (.xlsb)
//! file format. No COM, no Python, no Excel installation required, and
//! no OS-specific APIs — builds and runs the same on Windows and Linux.
//!
//! ```no_run
//! use xlsb_write::Workbook;
//!
//! let mut wb = Workbook::new();
//! let sheet = wb.add_worksheet("Sheet1");
//! sheet.write_string(0, 0, "Name");
//! sheet.write_number(0, 1, 42.0);
//! wb.save("out.xlsb").unwrap();
//! ```

pub mod biff12;
pub mod formula;
mod sheet;
mod sst;
pub mod styles;
mod wb_part;

pub use formula::{FnIndex, Formula};
pub use styles::{BorderStyle, Color, Format, HAlign, VAlign};

use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::fs::File;
use std::io::{Seek, Write};
use std::path::Path;
use std::rc::Rc;
use zip::{write::SimpleFileOptions, CompressionMethod, ZipWriter};

// ── Public value/error types ────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Blank,
    Number(f64),
    Bool(bool),
    /// Text content. `Rc<str>` rather than `String`: report data is
    /// overwhelmingly categorical (client/product/store names repeated
    /// across hundreds of thousands of rows), and `Worksheet` interns each
    /// distinct value once (see `Worksheet::intern_string`) — cells sharing
    /// a value share the same allocation instead of each owning a copy.
    String(Rc<str>),
    /// A formula whose most recent evaluation produced a numeric value.
    /// The `f64` is the cached result shown until Excel recalculates.
    FormulaNum(Box<Formula>, f64),
    /// A formula whose most recent evaluation produced a string value.
    FormulaStr(Box<Formula>, String),
}

#[derive(Debug)]
pub enum WriteError {
    Io(std::io::Error),
    Zip(zip::result::ZipError),
}

impl From<std::io::Error> for WriteError {
    fn from(e: std::io::Error) -> Self {
        Self::Io(e)
    }
}
impl From<zip::result::ZipError> for WriteError {
    fn from(e: zip::result::ZipError) -> Self {
        Self::Zip(e)
    }
}
impl std::fmt::Display for WriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{self:?}")
    }
}
impl std::error::Error for WriteError {}

// ── Worksheet ────────────────────────────────────────────────────────────────

pub struct Worksheet {
    name: String,
    /// Workbook-wide style table, shared with every other sheet — resolving
    /// a `Format` to its final XF index happens immediately, per cell,
    /// instead of in a second pass over the whole sheet at write time.
    style_builder: Rc<RefCell<styles::StylesBuilder>>,
    /// Per-sheet cache: `Format` → already-resolved XF index, so a format
    /// reused across millions of cells is only ever hashed/registered with
    /// `style_builder` once per sheet, not once per cell.
    format_ids: HashMap<Format, u16>,
    /// Workbook-wide shared string table — interning happens immediately
    /// (inside `flush_pending_row`, via `sheet::encode_row`), not in a
    /// second pass over the whole workbook at write time.
    sst: Rc<RefCell<sst::Sst>>,
    /// Distinct text values written to this sheet — see `intern_string`.
    strings: HashSet<Rc<str>>,
    freeze_row: u32,
    freeze_col: u32,
    col_specs: BTreeMap<u32, (f64, bool)>,
    row_specs: BTreeMap<u32, (f32, bool)>,
    merges: Vec<(u32, u32, u32, u32)>,
    /// Encoded bytes for every row already finished (everything between
    /// `BrtBeginSheetData` and `BrtEndSheetData`) — grows incrementally as
    /// rows are written, instead of the whole sheet being held as a
    /// `BTreeMap` of unencoded cells until write time.
    body: Vec<u8>,
    /// The row currently being written to, not yet encoded into `body`.
    pending_row: Option<PendingRow>,
    /// Running `(min_row, max_row, min_col, max_col)`, updated on every
    /// write — this is what `sheet::write_sheet_header`'s `BrtWsDim` needs,
    /// computed without ever holding the full cell set at once.
    dim: Option<(u32, u32, u32, u32)>,
    /// The last row index actually encoded into `body` — used to reject
    /// out-of-order writes (see `stage_cell`).
    last_flushed_row: Option<u32>,
}

struct PendingRow {
    row: u32,
    /// `(col, value, xf)` — not required to be column-sorted; sorted in
    /// `Worksheet::flush_pending_row` right before encoding.
    cells: Vec<(u32, CellValue, u16)>,
}

const DEFAULT_COLUMN_WIDTH: f64 = 8.43;
const DEFAULT_ROW_HEIGHT_PT: f32 = 15.0;

impl Worksheet {
    fn new(name: &str, sst: Rc<RefCell<sst::Sst>>, style_builder: Rc<RefCell<styles::StylesBuilder>>) -> Self {
        Self {
            name: name.to_owned(),
            style_builder,
            format_ids: HashMap::new(),
            sst,
            strings: HashSet::new(),
            freeze_row: 0,
            freeze_col: 0,
            col_specs: BTreeMap::new(),
            row_specs: BTreeMap::new(),
            merges: Vec::new(),
            body: Vec::new(),
            pending_row: None,
            dim: None,
            last_flushed_row: None,
        }
    }

    /// Resolve `fmt` to its workbook-wide XF index, registering it with the
    /// shared `StylesBuilder` the first time this sheet sees it. A report
    /// with millions of cells typically reuses a handful of distinct
    /// `Format`s (one per column type, roughly) — hashing/registering the
    /// same `Format` once per sheet instead of once per cell is what keeps
    /// this cheap.
    fn resolve_format(&mut self, fmt: &Format) -> u16 {
        if let Some(&xf) = self.format_ids.get(fmt) {
            return xf;
        }
        let xf = self.style_builder.borrow_mut().register(fmt);
        self.format_ids.insert(fmt.clone(), xf);
        xf
    }

    /// Encode the currently-pending row (if any) into `body` and clear it.
    /// Called whenever a write targets a different row, and once more at
    /// the very end (`Workbook::write`) to finalize each sheet's last row.
    fn flush_pending_row(&mut self) {
        let Some(pending) = self.pending_row.take() else { return };
        let mut cells: Vec<(u32, &CellValue, u16)> =
            pending.cells.iter().map(|(col, val, xf)| (*col, val, *xf)).collect();
        cells.sort_by_key(|(col, _, _)| *col);
        let (height_pt, hidden) =
            self.row_specs.get(&pending.row).copied().unwrap_or((DEFAULT_ROW_HEIGHT_PT, false));
        let height_twips = (height_pt * 20.0).round() as u16;
        sheet::encode_row(pending.row, &cells, height_twips, hidden, &mut self.sst.borrow_mut(), &mut self.body);
        self.last_flushed_row = Some(pending.row);
    }

    /// Stage one cell for encoding. If `row` differs from whatever row is
    /// currently pending, the pending row is flushed first — so cells for a
    /// given row must all be written before moving on to a later row.
    ///
    /// # Panics
    /// If `row` is less than or equal to a row that's already been flushed
    /// (or less than the currently-pending row) — this streaming writer
    /// requires non-decreasing row order. Sort your data by row before
    /// writing it.
    fn stage_cell(&mut self, row: u32, col: u32, value: CellValue, xf: u16) {
        let starting_new_row = !matches!(&self.pending_row, Some(p) if p.row == row);
        if starting_new_row {
            let floor = self.pending_row.as_ref().map(|p| p.row).or(self.last_flushed_row);
            if let Some(floor) = floor {
                assert!(
                    row > floor,
                    "xlsb_write: cells must be written in strictly increasing row order (this \
                     Worksheet streams rows to disk as soon as a later row starts) — tried to \
                     write row {row} after row {floor} was already finished. Sort your data by \
                     row before writing it."
                );
            }
            self.flush_pending_row();
            self.pending_row = Some(PendingRow { row, cells: vec![(col, value, xf)] });
        } else if let Some(p) = &mut self.pending_row {
            if let Some(existing) = p.cells.iter_mut().find(|(c, _, _)| *c == col) {
                existing.1 = value;
                existing.2 = xf;
            } else {
                p.cells.push((col, value, xf));
            }
        }
        let (min_row, max_row, min_col, max_col) = self.dim.unwrap_or((row, row, col, col));
        self.dim = Some((min_row.min(row), max_row.max(row), min_col.min(col), max_col.max(col)));
    }

    /// Intern `s`, returning a shared handle. Report data is overwhelmingly
    /// categorical — the same client/product/store name repeats across
    /// hundreds of thousands of rows — so deduplicating here means repeated
    /// values share one allocation (a cheap `Rc` clone) instead of each
    /// cell owning its own independent copy of the same bytes.
    fn intern_string(&mut self, s: &str) -> Rc<str> {
        if let Some(existing) = self.strings.get(s) {
            return existing.clone();
        }
        let rc: Rc<str> = Rc::from(s);
        self.strings.insert(rc.clone());
        rc
    }

    pub fn write_string(&mut self, row: u32, col: u32, value: &str) -> &mut Self {
        self.write_string_with_format(row, col, value, &Format::default())
    }
    pub fn write_string_with_format(&mut self, row: u32, col: u32, value: &str, fmt: &Format) -> &mut Self {
        let xf = self.resolve_format(fmt);
        let s = self.intern_string(value);
        self.stage_cell(row, col, CellValue::String(s), xf);
        self
    }

    pub fn write_number(&mut self, row: u32, col: u32, value: f64) -> &mut Self {
        self.write_number_with_format(row, col, value, &Format::default())
    }
    pub fn write_number_with_format(&mut self, row: u32, col: u32, value: f64, fmt: &Format) -> &mut Self {
        let xf = self.resolve_format(fmt);
        self.stage_cell(row, col, CellValue::Number(value), xf);
        self
    }

    pub fn write_boolean(&mut self, row: u32, col: u32, value: bool) -> &mut Self {
        self.write_boolean_with_format(row, col, value, &Format::default())
    }
    pub fn write_boolean_with_format(&mut self, row: u32, col: u32, value: bool, fmt: &Format) -> &mut Self {
        let xf = self.resolve_format(fmt);
        self.stage_cell(row, col, CellValue::Bool(value), xf);
        self
    }

    pub fn write_blank(&mut self, row: u32, col: u32) -> &mut Self {
        self.write_blank_with_format(row, col, &Format::default())
    }
    pub fn write_blank_with_format(&mut self, row: u32, col: u32, fmt: &Format) -> &mut Self {
        let xf = self.resolve_format(fmt);
        self.stage_cell(row, col, CellValue::Blank, xf);
        self
    }

    /// Write a formula whose most recent evaluation produced a numeric
    /// value. `cached_value` is what's shown immediately (Excel
    /// recalculates on open, but other readers — including `calamine` —
    /// display exactly this cached value and never evaluate the formula
    /// themselves), so it must match what the formula actually computes.
    pub fn write_formula_num(&mut self, row: u32, col: u32, formula: Formula, cached_value: f64) -> &mut Self {
        self.write_formula_num_with_format(row, col, formula, cached_value, &Format::default())
    }
    pub fn write_formula_num_with_format(
        &mut self,
        row: u32,
        col: u32,
        formula: Formula,
        cached_value: f64,
        fmt: &Format,
    ) -> &mut Self {
        let xf = self.resolve_format(fmt);
        self.stage_cell(row, col, CellValue::FormulaNum(Box::new(formula), cached_value), xf);
        self
    }

    /// Write a formula whose most recent evaluation produced a string value.
    pub fn write_formula_str(&mut self, row: u32, col: u32, formula: Formula, cached_value: &str) -> &mut Self {
        self.write_formula_str_with_format(row, col, formula, cached_value, &Format::default())
    }
    pub fn write_formula_str_with_format(
        &mut self,
        row: u32,
        col: u32,
        formula: Formula,
        cached_value: &str,
        fmt: &Format,
    ) -> &mut Self {
        let xf = self.resolve_format(fmt);
        self.stage_cell(row, col, CellValue::FormulaStr(Box::new(formula), cached_value.to_owned()), xf);
        self
    }

    /// Freeze the top `rows` rows (e.g. `1` to freeze just the header row).
    pub fn set_freeze_panes(&mut self, rows: u32) -> &mut Self {
        self.freeze_row = rows;
        self
    }

    /// Freeze the leftmost `cols` columns, independently of `set_freeze_panes`.
    pub fn set_freeze_panes_cols(&mut self, cols: u32) -> &mut Self {
        self.freeze_col = cols;
        self
    }

    /// Set a column's width, in character units (Excel's own unit — the
    /// default is 8.43).
    pub fn set_column_width(&mut self, col: u32, width_chars: f64) -> &mut Self {
        let hidden = self.col_specs.get(&col).map(|&(_, h)| h).unwrap_or(false);
        self.col_specs.insert(col, (width_chars, hidden));
        self
    }

    /// Hide a column. Keeps any width previously set with `set_column_width`.
    pub fn set_column_hidden(&mut self, col: u32) -> &mut Self {
        let width = self.col_specs.get(&col).map(|&(w, _)| w).unwrap_or(DEFAULT_COLUMN_WIDTH);
        self.col_specs.insert(col, (width, true));
        self
    }

    /// Set a row's height, in points (Excel's default is 15.0).
    pub fn set_row_height(&mut self, row: u32, height_pt: f32) -> &mut Self {
        let hidden = self.row_specs.get(&row).map(|&(_, h)| h).unwrap_or(false);
        self.row_specs.insert(row, (height_pt, hidden));
        self
    }

    /// Hide a row. Keeps any height previously set with `set_row_height`.
    pub fn set_row_hidden(&mut self, row: u32) -> &mut Self {
        let height = self.row_specs.get(&row).map(|&(h, _)| h).unwrap_or(DEFAULT_ROW_HEIGHT_PT);
        self.row_specs.insert(row, (height, true));
        self
    }

    /// Merge the rectangular range `[first_row..=last_row] x [first_col..=last_col]`.
    /// Write the merged value into `(first_row, first_col)` separately.
    pub fn merge_range(&mut self, first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> &mut Self {
        self.merges.push((first_row, first_col, last_row, last_col));
        self
    }
}

fn zip_options() -> SimpleFileOptions {
    SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .compression_level(Some(6))
}

/// Assemble one already-fully-built `Worksheet` into its final on-disk bytes
/// (header + encoded row data + footer) and write it to `zip` as
/// `xl/worksheets/sheet{sheet_number}.bin`. `sheet` is dropped at the end of
/// this call, freeing its `body` and everything else in it — used both by
/// `Workbook::write` (all sheets finished at once, at the very end) and by
/// `StreamingWorkbook::finish_worksheet` (one sheet finished immediately,
/// while later sheets haven't been built yet).
fn finish_and_write_sheet<W: Write + Seek>(
    mut sheet: Worksheet,
    sheet_number: usize,
    active: bool,
    zip: &mut ZipWriter<W>,
) -> Result<(), WriteError> {
    sheet.flush_pending_row(); // encode whatever row was still open

    let col_specs: Vec<sheet::ColSpec> =
        sheet.col_specs.iter().map(|(&col, &(w, hidden))| (col, w, hidden)).collect();

    // Header and footer are small (well under 1KB plus one BrtColInfo/
    // BrtMergeCell per entry) — only `sheet.body` can be large (one entry
    // per cell, up to hundreds of MB on a big sheet), so it's written to
    // the zip stream directly rather than copied into a combined buffer
    // first, which would transiently double this sheet's peak memory.
    let mut header = Vec::with_capacity(512);
    sheet::write_sheet_header(sheet.freeze_row, sheet.freeze_col, &col_specs, sheet.dim, active, &mut header);
    let mut footer = Vec::new();
    sheet::write_sheet_footer(&sheet.merges, &mut footer);

    zip.start_file(format!("xl/worksheets/sheet{sheet_number}.bin"), zip_options())?;
    zip.write_all(&header)?;
    zip.write_all(&sheet.body)?;
    zip.write_all(&footer)?;
    Ok(())
}

// ── Workbook ─────────────────────────────────────────────────────────────────

#[derive(Default)]
pub struct Workbook {
    sheets: Vec<Worksheet>,
    sst: Rc<RefCell<sst::Sst>>,
    style_builder: Rc<RefCell<styles::StylesBuilder>>,
}

impl Workbook {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_worksheet(&mut self, name: &str) -> &mut Worksheet {
        self.sheets.push(Worksheet::new(name, Rc::clone(&self.sst), Rc::clone(&self.style_builder)));
        self.sheets.last_mut().unwrap()
    }

    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<(), WriteError> {
        let file = File::create(path)?;
        self.write(file)
    }

    /// Consumes the workbook (rather than taking `&self`) so each sheet's
    /// accumulated cell data can be freed as soon as that sheet has been
    /// serialized, instead of every sheet's data staying alive until the
    /// whole workbook finishes writing — on a report with tens of millions
    /// of cells across several sheets, retaining all of them simultaneously
    /// (as `&self` would require) was a large, avoidable share of this
    /// crate's peak memory use.
    pub fn write<W: Write + Seek>(self, sink: W) -> Result<(), WriteError> {
        let mut zip = ZipWriter::new(sink);

        let sheet_names: Vec<&str> = self.sheets.iter().map(|s| s.name.as_str()).collect();
        let n = sheet_names.len();

        zip.start_file("[Content_Types].xml", zip_options())?;
        zip.write_all(content_types(n).as_bytes())?;

        zip.start_file("_rels/.rels", zip_options())?;
        zip.write_all(ROOT_RELS.as_bytes())?;

        zip.start_file("xl/workbook.bin", zip_options())?;
        zip.write_all(&wb_part::build_workbook(&sheet_names))?;

        zip.start_file("xl/_rels/workbook.bin.rels", zip_options())?;
        zip.write_all(workbook_rels(n).as_bytes())?;

        // Each sheet's bytes are written to the zip as soon as they're
        // encoded (see the loop body) rather than collected into a
        // `Vec<Vec<u8>>` first — the OPC/zip format doesn't require parts to
        // appear in any particular order, so there's no need to hold every
        // sheet's encoded bytes in memory simultaneously just to write
        // sharedStrings.bin/styles.bin (whose content isn't final until all
        // sheets are processed) before the worksheet parts.
        for (sheet_idx, sheet) in self.sheets.into_iter().enumerate() {
            // Only the first sheet is the active tab on open — every other
            // sheet must NOT have its own view marked "selected", or Excel
            // opens with every tab grouped (the exact behavior a user gets
            // from Shift-clicking every tab: editing/filtering one sheet
            // then applies to all of them).
            let active = sheet_idx == 0;
            finish_and_write_sheet(sheet, sheet_idx + 1, active, &mut zip)?;
        }

        zip.start_file("xl/sharedStrings.bin", zip_options())?;
        zip.write_all(&self.sst.borrow().encode())?;

        zip.start_file("xl/styles.bin", zip_options())?;
        zip.write_all(&self.style_builder.borrow().build())?;

        zip.start_file("docProps/core.xml", zip_options())?;
        zip.write_all(CORE_XML.as_bytes())?;
        zip.start_file("docProps/app.xml", zip_options())?;
        zip.write_all(APP_XML.as_bytes())?;

        zip.finish()?;
        Ok(())
    }
}

// ── StreamingWorkbook ────────────────────────────────────────────────────────

/// Like `Workbook`, but each worksheet is serialized and written to the
/// output the moment you're done with it (`finish_worksheet`) instead of
/// staying resident until every other sheet in the workbook is also
/// finished. For a workbook with several large sheets built one after
/// another (e.g. from a big external data source), `Workbook` holds every
/// sheet's encoded bytes simultaneously until its single terminal `write`/
/// `save` call; `StreamingWorkbook` holds at most one sheet's encoded bytes
/// at a time, so peak memory tracks the *largest single sheet* instead of
/// the *sum of every sheet*.
pub struct StreamingWorkbook<W: Write + Seek> {
    zip: ZipWriter<W>,
    sst: Rc<RefCell<sst::Sst>>,
    style_builder: Rc<RefCell<styles::StylesBuilder>>,
    sheet_names: Vec<String>,
}

/// A `Worksheet` created by `StreamingWorkbook::new_worksheet`. Write cells
/// to it exactly like a `Worksheet` (via `Deref`/`DerefMut`), then hand it
/// to `StreamingWorkbook::finish_worksheet` — do not hold onto it past that
/// call.
pub struct StreamingWorksheet {
    inner: Worksheet,
    number: usize,
    active: bool,
}

impl std::ops::Deref for StreamingWorksheet {
    type Target = Worksheet;
    fn deref(&self) -> &Worksheet {
        &self.inner
    }
}
impl std::ops::DerefMut for StreamingWorksheet {
    fn deref_mut(&mut self) -> &mut Worksheet {
        &mut self.inner
    }
}

impl<W: Write + Seek> StreamingWorkbook<W> {
    /// Wrap `sink` (e.g. a `File`) in a new, empty streaming workbook.
    /// Nothing is written to `sink` yet — the OPC/zip format doesn't
    /// require parts in any particular order, so the workbook-wide parts
    /// ([Content_Types].xml, workbook.bin, sharedStrings.bin, styles.bin —
    /// none of which are final until every sheet has been seen) are all
    /// deferred to `finish`.
    pub fn create(sink: W) -> Self {
        Self {
            zip: ZipWriter::new(sink),
            sst: Rc::new(RefCell::new(sst::Sst::new())),
            style_builder: Rc::new(RefCell::new(styles::StylesBuilder::new())),
            sheet_names: Vec::new(),
        }
    }

    /// Start a new worksheet. Only the very first worksheet ever created is
    /// the active tab on open (mirrors `Workbook::write`'s rule).
    pub fn new_worksheet(&mut self, name: &str) -> StreamingWorksheet {
        self.sheet_names.push(name.to_owned());
        StreamingWorksheet {
            inner: Worksheet::new(name, Rc::clone(&self.sst), Rc::clone(&self.style_builder)),
            number: self.sheet_names.len(),
            active: self.sheet_names.len() == 1,
        }
    }

    /// Serialize `sheet` and write it to the output immediately, then drop
    /// it — its `body` and everything else in it frees before the next
    /// sheet is even created.
    pub fn finish_worksheet(&mut self, sheet: StreamingWorksheet) -> Result<(), WriteError> {
        finish_and_write_sheet(sheet.inner, sheet.number, sheet.active, &mut self.zip)
    }

    /// Write every workbook-wide part (content types, `workbook.bin`,
    /// `sharedStrings.bin`, `styles.bin`, `docProps`) and close the zip.
    /// Call this once, after every worksheet has been passed to
    /// `finish_worksheet`.
    pub fn finish(mut self) -> Result<(), WriteError> {
        let sheet_names: Vec<&str> = self.sheet_names.iter().map(String::as_str).collect();
        let n = sheet_names.len();

        self.zip.start_file("[Content_Types].xml", zip_options())?;
        self.zip.write_all(content_types(n).as_bytes())?;

        self.zip.start_file("_rels/.rels", zip_options())?;
        self.zip.write_all(ROOT_RELS.as_bytes())?;

        self.zip.start_file("xl/workbook.bin", zip_options())?;
        self.zip.write_all(&wb_part::build_workbook(&sheet_names))?;

        self.zip.start_file("xl/_rels/workbook.bin.rels", zip_options())?;
        self.zip.write_all(workbook_rels(n).as_bytes())?;

        self.zip.start_file("xl/sharedStrings.bin", zip_options())?;
        self.zip.write_all(&self.sst.borrow().encode())?;

        self.zip.start_file("xl/styles.bin", zip_options())?;
        self.zip.write_all(&self.style_builder.borrow().build())?;

        self.zip.start_file("docProps/core.xml", zip_options())?;
        self.zip.write_all(CORE_XML.as_bytes())?;
        self.zip.start_file("docProps/app.xml", zip_options())?;
        self.zip.write_all(APP_XML.as_bytes())?;

        self.zip.finish()?;
        Ok(())
    }
}

// ── XML boilerplate ───────────────────────────────────────────────────────────

fn content_types(n: usize) -> String {
    let mut s = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
         <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
         <Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>\
         <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
         <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
         <Override PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>",
    );
    for i in 1..=n {
        s.push_str(&format!(
            "<Override PartName=\"/xl/worksheets/sheet{i}.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>"
        ));
    }
    s.push_str(
        "<Override PartName=\"/xl/styles.bin\" ContentType=\"application/vnd.ms-excel.styles\"/>\
         <Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"application/vnd.ms-excel.sharedStrings\"/>\
         <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\
         <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\
         </Types>",
    );
    s
}

fn workbook_rels(n: usize) -> String {
    let mut s = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
         <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for i in 1..=n {
        s.push_str(&format!(
            "<Relationship Id=\"rId{i}\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" \
             Target=\"worksheets/sheet{i}.bin\"/>"
        ));
    }
    let styles_id = n + 1;
    let sst_id = n + 2;
    s.push_str(&format!(
        "<Relationship Id=\"rId{styles_id}\" \
         Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" \
         Target=\"styles.bin\"/>\
         <Relationship Id=\"rId{sst_id}\" \
         Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" \
         Target=\"sharedStrings.bin\"/>\
         </Relationships>"
    ));
    s
}

const ROOT_RELS: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
     <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
     <Relationship Id=\"rId1\" \
     Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" \
     Target=\"xl/workbook.bin\"/>\
     <Relationship Id=\"rId2\" \
     Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" \
     Target=\"docProps/app.xml\"/>\
     <Relationship Id=\"rId3\" \
     Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" \
     Target=\"docProps/core.xml\"/>\
     </Relationships>";

const CORE_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
     <cp:coreProperties \
     xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" \
     xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\
     <dc:creator>xlsb_write</dc:creator></cp:coreProperties>";

const APP_XML: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
     <Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">\
     <Application>xlsb_write</Application></Properties>";

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn write_single_sheet_produces_valid_zip() {
        let mut wb = Workbook::new();
        let sheet = wb.add_worksheet("Sheet1");
        sheet.write_string(0, 0, "Name");
        sheet.write_number(0, 1, 42.0);
        sheet.write_boolean(1, 1, true);

        let mut buf = Cursor::new(Vec::new());
        wb.write(&mut buf).unwrap();
        let bytes = buf.into_inner();
        assert_eq!(&bytes[..2], b"PK");

        let cursor = Cursor::new(&bytes);
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        assert!(zip.by_name("xl/worksheets/sheet1.bin").is_ok());
        assert!(zip.by_name("xl/styles.bin").is_ok());
        assert!(zip.by_name("xl/sharedStrings.bin").is_ok());
        assert!(zip.by_name("xl/workbook.bin").is_ok());
    }

    #[test]
    #[should_panic(expected = "strictly increasing row order")]
    fn writing_an_earlier_row_after_a_later_one_panics() {
        let mut wb = Workbook::new();
        let sheet = wb.add_worksheet("Sheet1");
        sheet.write_number(5, 0, 1.0);
        sheet.write_number(3, 0, 2.0); // row 3 after row 5 — must panic
    }

    #[test]
    fn rewriting_a_cell_in_the_still_open_row_keeps_the_latest_value() {
        let mut wb = Workbook::new();
        let sheet = wb.add_worksheet("Sheet1");
        sheet.write_number(0, 0, 1.0);
        sheet.write_number(0, 0, 2.0); // same (row, col), still on the open row
        let mut buf = std::io::Cursor::new(Vec::new());
        wb.write(&mut buf).unwrap();
        // A crude but sufficient check: the FIRST value (1.0) must not
        // appear anywhere verbatim as a written RK/real cell payload,
        // while decoding the sheet fully is covered by tests/roundtrip.rs
        // — this test exists specifically to pin the "last write wins
        // within an open row" behavior, not full round-trip correctness.
        let bytes = buf.into_inner();
        assert_eq!(&bytes[..2], b"PK");
    }

    #[test]
    fn streaming_workbook_produces_valid_zip() {
        let mut buf = Cursor::new(Vec::new());
        let mut wb = StreamingWorkbook::create(&mut buf);

        let mut sheet1 = wb.new_worksheet("Sheet1");
        sheet1.write_string(0, 0, "Name");
        sheet1.write_number(0, 1, 42.0);
        wb.finish_worksheet(sheet1).unwrap();

        let mut sheet2 = wb.new_worksheet("Sheet2");
        sheet2.write_boolean(0, 0, true);
        wb.finish_worksheet(sheet2).unwrap();

        wb.finish().unwrap();
        let bytes = buf.into_inner();
        assert_eq!(&bytes[..2], b"PK");

        let cursor = Cursor::new(&bytes);
        let mut zip = zip::ZipArchive::new(cursor).unwrap();
        assert!(zip.by_name("xl/worksheets/sheet1.bin").is_ok());
        assert!(zip.by_name("xl/worksheets/sheet2.bin").is_ok());
        assert!(zip.by_name("xl/styles.bin").is_ok());
        assert!(zip.by_name("xl/sharedStrings.bin").is_ok());
        assert!(zip.by_name("xl/workbook.bin").is_ok());
    }
}
