# xlsb_write — Complete Guide

This is the full reference for `xlsb_write`: every public type and method,
how the pieces fit together, what it does well, and what it deliberately
does not do. The [README](README.md) is the pitch; this document is
everything you need to actually build something with it.

Table of contents:

- [Why `.xlsb`](#why-xlsb)
- [Installation](#installation)
- [Core concepts](#core-concepts)
- [Choosing `Workbook` vs `StreamingWorkbook`](#choosing-workbook-vs-streamingworkbook)
- [Writing cells](#writing-cells)
- [Formatting (`Format`, `Color`, borders, alignment, number formats)](#formatting)
- [Number formats in depth](#number-formats-in-depth)
- [Sheet layout (freeze panes, column/row sizing, merges)](#sheet-layout)
- [Embedding images](#embedding-images)
- [Autofilter](#autofilter)
- [Defined names (named ranges)](#defined-names-named-ranges)
- [Formulas](#formulas)
- [Full worked example](#full-worked-example)
- [Error handling](#error-handling)
- [Performance model](#performance-model)
- [Correctness: how this crate is verified](#correctness-how-this-crate-is-verified)
- [Limitations — what's not supported](#limitations--whats-not-supported)
- [Troubleshooting](#troubleshooting)
- [Internals, for contributors](#internals-for-contributors)

---

## Why `.xlsb`

`.xlsb` is Excel's **binary** workbook format (BIFF12) — functionally the
same as `.xlsx` (same features: formulas, formatting, multiple sheets), but:

- **Smaller on disk** — binary records instead of verbose XML, typically
  significantly smaller for data-heavy workbooks.
- **Faster for Excel to open and save** — no XML parsing.
- **Still a real `.xlsx`-family file** — it's a ZIP/OPC package like
  `.xlsx`, just with `xl/worksheets/sheet1.bin` (BIFF12 records) instead of
  `xl/worksheets/sheet1.xml`.

This crate writes those `.bin` parts directly, byte-for-byte, without going
through Excel, COM, or any XML intermediate. That means:

- **No Excel installation required** — works in CI, on Linux servers,
  inside containers, anywhere Rust runs.
- **No Python, no COM automation** — no `xlwings`, no `pywin32`, no
  platform-specific glue.
- **Fast** — encoding is a straight in-memory walk over your data; there's
  no interpreter or automation layer in the loop.

The tradeoff: this is a **writer**, not a full spreadsheet engine. It emits
a valid, spec-conformant `.xlsb` that Excel opens and displays correctly,
but it does not read existing workbooks, and it does not implement every
corner of the OOXML/XLSB feature surface (see
[Limitations](#limitations--whats-not-supported)).

## Installation

```toml
[dependencies]
xlsb_write = "0.1"
```

Everything lives under the `xlsb_write` crate root:

```rust
use xlsb_write::{
    Workbook, StreamingWorkbook, Worksheet,
    Format, Color, BorderStyle, HAlign, VAlign,
    Formula, FnIndex,
    WriteError, CellValue,
};
```

You'll rarely need `CellValue` or `FnIndex` directly — they exist mostly so
the compiler can talk about what the crate is doing internally.

## Core concepts

- **Rows and columns are zero-based `u32`.** `(0, 0)` is cell `A1`. There is
  no `"A1"`-style string addressing anywhere in the API — always `(row,
  col)` pairs.
- **A `Format` is a value, not a handle.** You build one with the builder
  pattern (`Format::new().set_bold()...`), and pass it *by reference* to
  every `write_*_with_format` call. Two structurally-equal `Format`s
  (including two independently-built ones with identical settings) are
  automatically deduplicated into a single style entry in the output file —
  you don't need to intern or cache formats yourself; the crate does that
  for you (see [Performance model](#performance-model)).
- **Text is deduplicated automatically.** Repeated strings (the same
  category name, region, product code, etc. across many rows) are written
  once to Excel's shared string table and referenced by index — you don't
  need to do anything to get this; every `write_string`/
  `write_string_with_format` call benefits from it.
- **Two ways to build a workbook**: `Workbook` (build everything, write
  once at the end) and `StreamingWorkbook` (finish and flush one sheet at a
  time). See the next section for when to use which.
- **Formulas are built, not parsed.** There's no string formula parser —
  you construct a `Formula` value with method calls
  (`Formula::cell(0,0).add(Formula::num(1.0))`) and the crate encodes it
  directly to the binary token stream Excel expects. See
  [Formulas](#formulas).

## Choosing `Workbook` vs `StreamingWorkbook`

Both produce byte-identical output for the same sequence of calls — the
only difference is *when* memory is freed.

### `Workbook`

```rust
use xlsb_write::Workbook;

let mut wb = Workbook::new();
let sheet = wb.add_worksheet("Sheet1");
sheet.write_string(0, 0, "Name");
sheet.write_number(0, 1, 42.0);

let sheet2 = wb.add_worksheet("Sheet2");
sheet2.write_string(0, 0, "Another sheet");

wb.save("out.xlsb")?;          // or wb.write(some_writer)?;
```

- `add_worksheet` returns a `&mut Worksheet` borrowed from the `Workbook`,
  so all your sheets stay alive (and resident in memory) until the single
  terminal `.save()`/`.write()` call.
- Simplest API — reach for this unless you have a specific memory reason
  not to.
- Peak memory ≈ the **sum** of every sheet's encoded size, since nothing is
  flushed to the output until the very end.

### `StreamingWorkbook`

```rust
use xlsb_write::StreamingWorkbook;
use std::fs::File;

let file = File::create("out.xlsb")?;
let mut wb = StreamingWorkbook::create(file);

let mut sheet1 = wb.new_worksheet("Sheet1");
sheet1.write_string(0, 0, "Name");
sheet1.write_number(0, 1, 42.0);
wb.finish_worksheet(sheet1)?;   // encodes + writes sheet1 to disk now, frees it

let mut sheet2 = wb.new_worksheet("Sheet2");
sheet2.write_string(0, 0, "Another sheet");
wb.finish_worksheet(sheet2)?;

wb.finish()?;                   // writes workbook-wide parts, closes the zip
```

- `new_worksheet` returns an owned `StreamingWorksheet` (a thin wrapper
  around `Worksheet` via `Deref`/`DerefMut` — every `Worksheet` method
  works on it unchanged). You must hand it to `finish_worksheet` when done;
  don't hold onto it past that call.
- Only **one sheet's** encoded bytes are resident at a time — peak memory
  tracks the size of your **largest single sheet**, not the sum of all of
  them. For a workbook built from an external data source (e.g. reading
  rows from a database or a large Parquet/CSV file, sheet by sheet), this
  is the difference between "works" and "OOM" on a big report.
- Requires `W: Write + Seek` (the underlying `ZipWriter` needs to seek to
  patch central-directory offsets) — a `File`, an in-memory `Cursor<Vec<u8>>`,
  or anything else implementing both traits.
- You must call `new_worksheet` / `finish_worksheet` in the order you want
  sheets to appear, and call `.finish()` exactly once, after every sheet.

**Rule of thumb:** start with `Workbook`. Switch to `StreamingWorkbook` only
once you've confirmed memory is actually a problem (very large sheets, or
several large sheets in one workbook) — it's a strictly more constrained
API (streamed, ordered) for the same output.

## Writing cells

Every write method exists in a plain form and a `_with_format` form. The
plain form is exactly `write_x_with_format(row, col, value, &Format::default())`.

```rust
sheet.write_string(row, col, "text");
sheet.write_string_with_format(row, col, "text", &header_fmt);

sheet.write_number(row, col, 42.0);
sheet.write_number_with_format(row, col, 42.0, &int_fmt);

sheet.write_boolean(row, col, true);
sheet.write_boolean_with_format(row, col, true, &fmt);

sheet.write_blank(row, col);                    // an empty cell, still formatted/bordered
sheet.write_blank_with_format(row, col, &fmt);

sheet.write_formula_num(row, col, formula, cached_f64);
sheet.write_formula_num_with_format(row, col, formula, cached_f64, &fmt);

sheet.write_formula_str(row, col, formula, cached_str);
sheet.write_formula_str_with_format(row, col, formula, cached_str, &fmt);
```

All of these return `&mut Self`, so calls chain:

```rust
sheet
    .write_string(0, 0, "A")
    .write_string(0, 1, "B")
    .write_number(1, 0, 1.0);
```

Notes:

- **An empty string writes as blank.** `write_string(r, c, "")` produces the
  same result as `write_blank(r, c)` (checked internally — you don't get an
  empty shared-string entry).
- **Writing the same `(row, col)` twice while that row is still open**
  (i.e. before you've moved on to writing a later row) keeps the **last**
  value written — this lets you build a row incrementally and correct a
  cell before moving on.
- **Rows must be written in non-decreasing order, and once you move past a
  row it's finalized.** Concretely: you can interleave writes to columns
  within the *current* row in any order, but once you write to a strictly
  later row, every earlier row is flushed (encoded to the internal buffer,
  in `StreamingWorkbook`'s case immediately to disk) and can never be
  touched again. Writing to a row at or before one that's already been
  flushed **panics**:

  ```text
  xlsb_write: cells must be written in strictly increasing row order (this
  Worksheet streams rows to disk as soon as a later row starts) — tried to
  write row 3 after row 5 was already finished. Sort your data by row
  before writing it.
  ```

  This applies to `Workbook` too, not just `StreamingWorkbook` — both share
  the same `Worksheet` internals, and the row-streaming buffer is what
  makes memory use scale with *row width* rather than *total cell count
  held as a map*. **If your source data isn't already row-ordered, sort it
  by row before writing.**

## Formatting

`Format` is a builder. Build one, then pass `&format` to any
`_with_format` call — build it once outside your write loop and reuse the
reference; the crate deduplicates identical `Format`s into one style entry
regardless of how many times you pass it.

```rust
use xlsb_write::{Format, Color, BorderStyle, HAlign, VAlign};

let header = Format::new()
    .set_bold()
    .set_font_color(Color::WHITE)
    .set_background_color(Color::rgb(31, 73, 125))
    .set_align(HAlign::Center, VAlign::Center)
    .set_border(BorderStyle::Thin)
    .set_border_color(Color::BLACK);
```

### `Format` methods (all consume `self` and return `Self` — chain freely)

| Method | Effect |
|---|---|
| `Format::new()` | Start from the default (unformatted) style. |
| `.set_bold()` | Bold text. |
| `.set_italic()` | Italic text. |
| `.set_underline()` | Single underline. |
| `.set_strikeout()` | Strikethrough. |
| `.set_font_color(Color)` | Text color. Default: theme "Text 1" (usually black). |
| `.set_font_size(points: f32)` | Font size in points (e.g. `12.0`). Default: 11pt. Rounded to the nearest half-point internally. |
| `.set_font_name(&str)` | Font family name (e.g. `"Arial"`). Default: `"Calibri"`. |
| `.set_background_color(Color)` | Solid cell fill color. |
| `.set_border(BorderStyle)` | Border style, applied to **all four sides** identically (there's no per-side control). |
| `.set_border_color(Color)` | Border color. Only takes effect if `.set_border(...)` is also set to something other than `None`. |
| `.set_align(HAlign, VAlign)` | Horizontal + vertical alignment (set together, one call). |
| `.set_text_wrap()` | Wrap text within the cell. |
| `.set_num_format(&str)` | Number format — a shorthand keyword or a raw Excel format code. See [Number formats in depth](#number-formats-in-depth). |

`Format::default()` (equivalently `Format::new()` with nothing else called)
is special-cased to always map to Excel's built-in style index `0` — using
the default format everywhere costs nothing extra in the output file.

### `Color`

```rust
Color::rgb(r: u8, g: u8, b: u8) -> Color   // any RGB color
Color::BLACK
Color::WHITE
Color::RED
Color::GREEN   // 0,128,0 — the standard "dark-ish" web green, not (0,255,0)
Color::BLUE
```

There's no named-color table beyond these five — for anything else, use
`Color::rgb(...)` directly.

### `BorderStyle`

```rust
pub enum BorderStyle {
    None,     // default
    Thin, Medium, Thick,
    Dashed, Dotted,
    Double,
    Hair,
    MediumDashed, DashDot, MediumDashDot, DashDotDot, MediumDashDotDot,
    SlantDashDot,
}
```

Applied uniformly to all four sides of the cell — this crate doesn't (yet)
support different styles per side, or diagonal borders.

### `HAlign` / `VAlign`

```rust
pub enum HAlign {
    General,   // default
    Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed,
}
pub enum VAlign {
    Top, Center, Bottom,   // default is Bottom, matching Excel's own default
    Justify, Distributed,
}
```

## Number formats in depth

`.set_num_format(&str)` accepts either a **shorthand keyword** or a **raw
Excel number-format code**. Shorthand keywords (case-insensitive) resolve
to the format codes below:

| Shorthand | Resolves to | Looks like |
|---|---|---|
| `"general"` | `General` | `1234.5` |
| `"int"` | `#,##0` | `1,235` |
| `"int0"` | `0` | `1235` |
| `"float1"` | `#,##0.0` | `1,234.5` |
| `"float2"` | `#,##0.00` | `1,234.50` |
| `"float3"` | `#,##0.000` | `1,234.500` |
| `"float4"` | `#,##0.0000` | `1,234.5000` |
| `"pct"` | `0%` | `12%` |
| `"pct1"` | `0.0%` | `12.3%` |
| `"pct2"` | `0.00%` | `12.34%` |
| `"sci"` | `0.00E+00` | `1.23E+03` |
| `"date"` | `YYYY-MM-DD` | `2026-09-13` |
| `"datetime"` | `YYYY-MM-DD HH:MM:SS` | `2026-09-13 14:30:00` |
| `"time"` | `HH:MM:SS` | `14:30:00` |
| `"text"` | `@` | forces text display, no numeric interpretation |
| `"currency"` | `$#,##0.00` | `$1,234.50` |
| `"euro"` | `€#,##0.00` | `€1,234.50` |
| `"accounting"` | `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` | accounting-style column alignment |

Anything else you pass through `.set_num_format(...)` that isn't one of
these keywords is treated as a **raw Excel number-format code** and used
verbatim (e.g. `.set_num_format("0.0\" kg\"")`). A handful of the most
common raw codes (`"General"`, `"0"`, `"0.00"`, `"#,##0"`, `"#,##0.00"`,
`"0%"`, `"0.0%"`/`"0.00%"`, `"0.00E+00"`, `"m/d/yyyy"`, `"@"`) map to
Excel's **built-in** format IDs and cost nothing extra; anything else is
registered as a custom format in the styles table (deduplicated the same
way `Format`s are — using the same custom code twice reuses one entry).

**Dates**: there's no `CellValue::Date` — write a date as a **number**
(the Excel serial date: days since 1899-12-30) with a date-shaped
`.set_num_format(...)`. If you're converting from a Unix-epoch value
(days since 1970-01-01, e.g. what Parquet's `DATE` logical type stores),
add `25569.0`:

```rust
let excel_serial = unix_epoch_days as f64 + 25569.0;
sheet.write_number_with_format(row, col, excel_serial, &date_fmt);
```

## Sheet layout

```rust
sheet.set_freeze_panes(1);           // freeze the top 1 row (e.g. a header)
sheet.set_freeze_panes_cols(1);      // independently freeze the left 1 column

sheet.set_column_width(col, 14.0);   // width in Excel's own character units (default 8.43)
sheet.set_column_hidden(col);        // hide a column (keeps any width you set separately)

sheet.set_row_height(row, 20.0);     // height in points (default 15.0)
sheet.set_row_hidden(row);           // hide a row (keeps any height you set separately)

sheet.merge_range(first_row, first_col, last_row, last_col);
```

- `set_freeze_panes` and `set_freeze_panes_cols` are independent — call
  both for a classic "frozen header row + frozen first column" layout
  (e.g. `set_freeze_panes(1)` + `set_freeze_panes_cols(1)` freezes row 1
  and column A simultaneously, like Excel's "Freeze Panes" at cell B2).
- For `merge_range`, write the visible value into `(first_row, first_col)`
  yourself with a normal `write_*` call — merging only affects the visual
  range, not which cell holds the value.
- All of these can be called in any order relative to your cell writes,
  **except** that column/row specs and merges are only finalized when the
  sheet itself is finished (`Workbook::write`/`StreamingWorkbook::finish_worksheet`)
  — so call them any time before that point, there's no ordering constraint
  relative to `write_*` calls the way rows themselves have.

## Embedding images

```rust
use xlsb_write::ImageFormat;

let png_bytes: Vec<u8> = std::fs::read("logo.png")?;
sheet.embed_image(first_row, first_col, last_row, last_col, &png_bytes, ImageFormat::Png);
```

- `embed_image(first_row, first_col, last_row, last_col, image_bytes, format)`
  anchors the image to the inclusive rectangular cell range
  `[first_row..=last_row] x [first_col..=last_col]` — write the range you'd
  select in Excel before choosing Insert → Picture, using this crate's
  usual zero-based `(row, col)` addressing.
- **This is a "two-cell anchor"**: the image's on-screen size tracks the
  actual column widths/row heights of the range it's anchored to. Widen a
  column or heighten a row it spans, and the image grows with it — the
  same behavior as inserting a picture "into" a cell range in Excel
  itself, not a fixed-size picture that merely starts at that cell.
  (Verified empirically against real Excel — see
  [Correctness](#correctness-how-this-crate-is-verified).)
- `ImageFormat` is `Png` or `Jpeg` — pass whichever matches your bytes.
  `embed_image` checks the file's magic number against `format` and
  panics on a mismatch (a PNG's bytes passed with `ImageFormat::Jpeg`, or
  vice versa), since that would silently produce a picture Excel can't
  decode.
- `image_bytes` is written verbatim into `xl/media/imageN.<ext>` — this
  crate never decodes, validates, resizes, or re-encodes image data.
- A sheet can have any number of images (in different, or overlapping,
  ranges); each becomes its own anchor inside that sheet's one shared
  drawing part.
- Works identically on `Worksheet`, `StreamingWorksheet` (via
  `Deref`/`DerefMut`, like every other `Worksheet` method), and
  `SizedStreamingWorksheet` — call it any time before the sheet is
  finished; unlike `set_column_width`/`set_freeze_panes` on
  `SizedStreamingWorksheet`, it has no "before the header is sent"
  restriction, since an image isn't part of the header.

```rust
use xlsb_write::{ImageFormat, Workbook};

let mut wb = Workbook::new();
let sheet = wb.add_worksheet("Sheet1");
sheet.write_string(0, 0, "Logo:");
sheet.embed_image(1, 1, 5, 3, &png_bytes, ImageFormat::Png); // anchored to B2:D6
wb.save("out.xlsb")?;
```

Not supported (out of scope for this first pass — see
[Limitations](#limitations--whats-not-supported)): image resizing/
cropping/rotation, transparency effects beyond what the source PNG/JPEG
already encodes, formats other than PNG/JPEG (BMP, GIF, TIFF, EMF/WMF),
and charts (a related but separate, larger subsystem — not implemented at
all).

## Autofilter

```rust
sheet.set_autofilter(first_row, first_col, last_row, last_col);
```

- Turns the inclusive rectangular range `[first_row..=last_row] x
  [first_col..=last_col]` into an autofilter — the same effect as selecting
  that range and choosing Data > AutoFilter in Excel. This only turns on
  the dropdown arrows on the header row; it does not pre-set any filter
  criteria on any column (there's no API for that yet).
- A sheet has at most one autofilter range — calling `set_autofilter` again
  replaces the previous one.
- Works identically on `Worksheet`, `StreamingWorksheet` (via
  `Deref`/`DerefMut`), and `SizedStreamingWorksheet` — call it any time
  before the sheet is finished; like `embed_image`, it has no "before the
  header is sent" restriction on `SizedStreamingWorksheet`, since
  autofilter isn't part of the header.

```rust
use xlsb_write::Workbook;

let mut wb = Workbook::new();
let sheet = wb.add_worksheet("Sheet1");
sheet.write_string(0, 0, "Name");
sheet.write_string(0, 1, "Qty");
sheet.write_string(1, 0, "Alpha");
sheet.write_number(1, 1, 10.0);
sheet.set_autofilter(0, 0, 1, 1); // header + one data row, A1:B2
wb.save("out.xlsb")?;
```

Not supported: pre-set filter criteria on any column (values/conditions
selected in a dropdown) — this only shows the dropdowns, matching "Data >
AutoFilter" with nothing filtered yet.

## Defined names (named ranges)

```rust
wb.define_name(name, sheet_index, first_row, first_col, last_row, last_col);
```

- A **workbook-level** named range (`Formulas > Define Name` in Excel) —
  this is why `define_name` lives on `Workbook`/`StreamingWorkbook`, not
  `Worksheet`: named ranges are workbook-scoped even though the range
  itself lives on one specific sheet.
- `sheet_index` is the 0-based index of the sheet the range lives on, in
  the order sheets were added (`add_worksheet`/`new_worksheet`/
  `new_worksheet_sized` call order) — not validated until `write`/`save`/
  `finish`, so `define_name` can be called before the target sheet exists
  yet, as long as it exists by the time the workbook is finished.
- The range is `[first_row..=last_row] x [first_col..=last_col]`, same
  zero-based inclusive addressing as everywhere else in this crate.

```rust
use xlsb_write::Workbook;

let mut wb = Workbook::new();
let sheet = wb.add_worksheet("Sheet1");
sheet.write_string(0, 0, "Q1");
sheet.write_number(1, 0, 42.0);

wb.define_name("Q1Total", 0, 1, 0, 1, 0); // Sheet1!$A$2
wb.save("out.xlsb")?;
```

```rust
use xlsb_write::StreamingWorkbook;
use std::fs::File;

let file = File::create("out.xlsb")?;
let mut wb = StreamingWorkbook::create(file);
let mut sheet = wb.new_worksheet("Sheet1");
sheet.write_string(0, 0, "Header");
wb.finish_worksheet(sheet)?;
wb.define_name("HeaderCell", 0, 0, 0, 0, 0); // can be called before or after finish_worksheet
wb.finish()?;
```

Excel opens the file with the name visible in the Name Box (top-left of
the formula bar) and resolvable in formulas (`=SUM(MyRange)`) — confirmed
via COM automation (see
[Correctness](#correctness-how-this-crate-is-verified)).

Not supported:
- Sheet-scoped names (a name only visible/usable from one specific sheet)
  — every name this crate writes is workbook-scoped.
- Names referring to a non-contiguous selection, a formula/constant instead
  of a cell range, or more than one area.
- Print areas (`_xlnm.Print_Area`) or other `_xlnm`-prefixed built-in
  names — `define_name` is for ordinary user-visible named ranges only.

## Formulas

`Formula` is a small expression builder — there's no text-formula parser
anywhere in this crate. You build an expression tree with method calls,
and it's encoded directly to the binary token stream (Rgce) Excel expects.

### Building expressions

```rust
use xlsb_write::Formula;

Formula::cell(row, col)                              // a single-cell reference
Formula::range(r0, c0, r1, c1)                        // a rectangular range
Formula::num(3.14)                                    // a numeric literal
Formula::str("hello")                                 // a string literal
Formula::boolean(true)                                // a boolean literal (PtgBool)

// arithmetic (all consume self, return a new Formula — chain freely)
a.add(b)   // a + b
a.sub(b)   // a - b
a.mul(b)   // a * b
a.div(b)   // a / b
a.concat(b) // a & b (string concatenation)

// comparisons
a.lt(b)  a.le(b)  a.eq(b)  a.ge(b)  a.gt(b)  a.ne(b)

// SUM over a range — the common case gets a shortcut:
Formula::sum_range(first_row, first_col, last_row, last_col)
// equivalent to (but use the shortcut — see note below):
// Formula::Func(FnIndex::SUM, vec![Formula::range(...)])

// branching
Formula::if_then_else(cond, then, else_)              // IF(cond, then, else)
Formula::iferror(expr, default)                       // IFERROR(expr, default)

// string, logical, conditional-aggregate, text, date, and lookup
// functions — see "More built-in functions" below.
Formula::concatenate(vec![a, b, c])                   // CONCATENATE(a, b, c)
Formula::and(vec![a, b])   Formula::or(vec![a, b])   Formula::not(a)
Formula::sumif(range, criteria)   Formula::countif(range, criteria)
Formula::left(text, n)   Formula::right(text, n)   Formula::mid(text, start, n)
Formula::len(text)   Formula::text(value, format_code)
Formula::today()   Formula::now()   Formula::date(year, month, day)
Formula::vlookup(lookup_value, table, col_index, exact_match)
Formula::index(array, row_num, col_num)
Formula::match_(lookup_value, array, match_type)      // trailing `_`: `match` is a keyword
```

Example — `(A1 + 5) * 2`:

```rust
let f = Formula::cell(0, 0).add(Formula::num(5.0)).mul(Formula::num(2.0));
sheet.write_formula_num(row, col, f, computed_value);
```

### Writing a formula cell — you must supply the cached value

```rust
sheet.write_formula_num(row, col, formula, cached_value: f64);
sheet.write_formula_str(row, col, formula, cached_value: &str);
sheet.write_formula_bool(row, col, formula, cached_value: bool);
```

Use `write_formula_bool` for a formula whose result is a boolean —
`AND`/`OR`/`NOT`, or a bare comparison (`a.gt(b)`) used as the whole
formula. Real Excel writes these as a distinct record type (`BrtFmlaBool`,
with a 1-byte cached value) rather than `BrtFmlaNum`'s 8-byte float, and
this crate matches that — confirmed byte-for-byte against real Excel's own
`=AND(...)` output.

`cached_value` is what's displayed **immediately**, before any
recalculation happens. Excel itself recalculates on open (so a wrong
cached value self-corrects the moment a human opens the file in Excel),
but **other readers — including `calamine`, the library this crate's own
tests use to verify output — display exactly the cached value and never
evaluate the formula themselves.** If you're consuming the file
programmatically rather than through Excel, a wrong cached value is a real
bug, not a cosmetic one. **Always compute `cached_value` from the same
data you built the formula from** (as the `sales_report.rs` example does —
every formula's cached value is computed in plain Rust alongside the
`Formula` expression, from the same source values).

### Built-in functions (`FnIndex`)

```rust
FnIndex::COUNT, FnIndex::ISNA, FnIndex::ISERROR,
FnIndex::SUM, FnIndex::AVERAGE, FnIndex::MIN, FnIndex::MAX, FnIndex::ROUND,
FnIndex::CONCATENATE, FnIndex::AND, FnIndex::OR, FnIndex::NOT,
FnIndex::SUMIF, FnIndex::COUNTIF,
FnIndex::LEFT, FnIndex::RIGHT, FnIndex::MID, FnIndex::LEN, FnIndex::TEXT,
FnIndex::TODAY, FnIndex::NOW, FnIndex::DATE,
FnIndex::VLOOKUP, FnIndex::INDEX, FnIndex::MATCH
```

These are the only function indices this crate's own test suite has
verified against real Excel output — use one of them via `Formula::Func`/
`Formula::FuncFixed` (or, better, the dedicated `Formula::and`/
`Formula::sumif`/... constructors below) whenever possible.

For a function not listed here, `FnIndex`'s inner value is public:
`Formula::Func(FnIndex(0x0182), vec![...])` works without forking this
crate. This is **unverified by this crate** — look the real index up
yourself in the published MS-XLS `Ftab` enumeration (don't guess), and
double-check both the argument count (`PtgFuncVar`'s `cparams` byte is
written from `args.len()`) **and** whether real Excel actually uses the
general `PtgFuncVar` form at all: some functions (`NOT`, `COUNTIF`, `MID`,
`LEN`, `TEXT`, `DATE`, ...) are *fixed*-argument-count in real Excel and
are encoded with the narrower `PtgFunc` form instead (no `cparams` byte).
Guessing wrong on either point is exactly the class of subtle bug this
crate has hit before with its own built-in functions — see
`src/formula.rs`'s `Formula::Func`/`Formula::FuncFixed` doc comments for
exactly which of the functions below are which, and why argument count
alone doesn't tell you (`SUMIF` and `COUNTIF` both commonly take 2 args,
but only `COUNTIF`'s is fixed).

**Use `Formula::sum_range` for a single-range `SUM`, not a hand-built
`Formula::Func`.** A single-range `SUM` is the one case real Excel encodes
with a dedicated shortcut token (`PtgAttrSum`) rather than the general
function-call form — `sum_range` produces exactly that byte-for-byte match
against real Excel output. Using the general form is spec-legal and opens
fine, but was found (during this crate's own hardening) to make Excel's
dynamic-array engine insert a spurious `@` into the formula on load. Stick
to the provided constructor and you won't hit this.

### More built-in functions

Every constructor below was checked against real Excel output the same
way `SUM`/`IF`/`IFERROR` were: build a tiny reference `.xlsb` with real
Excel via COM automation, then compare this crate's encoding byte-for-byte
against Excel's own `Rgce` bytes for the equivalent formula (via
`examples/dump_sheet.rs`/`examples/rgce_disasm.rs`).

**String concatenation:**

```rust
a.concat(b)                          // a & b — PtgConcat, a new binary infix Ptg
Formula::concatenate(vec![a, b, c])  // CONCATENATE(a, b, c) — variadic function
```

Prefer `a.concat(b)` for two values (one token shorter, and what Excel's
own UI produces for `=A1&B1`); use `concatenate` for 3+ values or to match
literal `CONCATENATE(...)` formula text.

**Logical functions:**

```rust
Formula::and(vec![cond1, cond2, ...])  // AND(...) — variadic
Formula::or(vec![cond1, cond2, ...])   // OR(...) — variadic
Formula::not(cond)                     // NOT(cond)
```

`AND`/`OR`/`NOT` (and any bare comparison used as a whole formula) produce
a **boolean** result — write them with `write_formula_bool`, not
`write_formula_num`.

**Conditional aggregates:**

```rust
Formula::sumif(range, criteria)     // SUMIF(range, criteria) — 2-arg form
Formula::countif(range, criteria)   // COUNTIF(range, criteria) — always 2 args
```

`criteria` is a normal `Formula` — a comparison-expression string
(`Formula::str(">10")`), a plain value to match (`Formula::str("apples")`,
`Formula::num(10.0)`), etc.; real Excel encodes it as a plain string/number
literal, nothing special. `SUMIF` also has a 3-arg form
(`SUMIF(range, criteria, sum_range)`, to total a *different* column than
the one being matched) — not wrapped in its own constructor, but reachable
directly (and separately byte-verified) via
`Formula::Func(FnIndex::SUMIF, vec![range, criteria, sum_range])`.

**Text functions:**

```rust
Formula::left(text, num_chars)             // LEFT(text, num_chars)
Formula::right(text, num_chars)            // RIGHT(text, num_chars)
Formula::mid(text, start_num, num_chars)   // MID(text, start_num, num_chars)
Formula::len(text)                         // LEN(text)
Formula::text(value, format_text)          // TEXT(value, format_text)
```

`format_text` (e.g. `Formula::str("0.00")`) is a plain string literal, same
as everywhere else — but note that Excel evaluates number-format codes
using the **current user's regional settings** for the decimal/thousands
separator characters (`.`/`,`); a format code written with US-style `.`
can render unexpectedly on a machine set to a locale that uses `,` as the
decimal separator. This is a general Excel behavior, not something this
crate's encoding controls or can compensate for — the bytes written are
exactly what real Excel itself writes for the same format string.

**Date functions:**

```rust
Formula::today()                        // TODAY() — no arguments
Formula::now()                          // NOW() — no arguments
Formula::date(year, month, day)         // DATE(year, month, day)
```

`TODAY`/`NOW` are **volatile**: real Excel recalculates them on every
calculation pass, not just when a dependency changes, because a formula
with no cell references would otherwise never be flagged for
recalculation. This crate reproduces both halves of how real Excel marks
that: a `PtgAttrSemi` token wrapping the call, and a bit set in the cell
record's `grbitFlags` (`Formula::is_volatile()` detects this
automatically — you don't need to do anything beyond using `today()`/
`now()`). `DATE(...)` is an ordinary, non-volatile function.

**Lookup functions:**

```rust
Formula::vlookup(lookup_value, table_array, col_index_num, range_lookup: bool)
Formula::index(array, row_num, column_num)
Formula::match_(lookup_value, lookup_array, match_type)   // `match` is a keyword, hence `match_`
```

`vlookup`'s 4th argument is a plain Rust `bool` (encoded as `Formula::Bool`
— `PtgBool`, confirmed byte-for-byte against real Excel's own
`VLOOKUP(...,FALSE)`), covering the overwhelmingly common case of a
literal `TRUE`/`FALSE`; build `Formula::Func(FnIndex::VLOOKUP, vec![...])`
directly if you need a computed 4th argument instead. `index`/`match_`
cover the common 2-4/2-3-arg forms real Excel itself uses; `INDEX`'s rarer
1-arg area-only form and 4-arg `area_num` form aren't wrapped in their own
constructor but are reachable the same way.

### `IF` / `IFERROR` semantics

- `Formula::if_then_else(cond, then, else_)` encodes to genuine branching
  tokens (`PtgAttrIf`/`PtgAttrGoto`) — only the taken branch actually
  evaluates in Excel, exactly like a real `IF()`.
- `Formula::iferror(expr, default)` is **desugared** to
  `IF(ISERROR(expr), default, expr)` rather than using XLSB's newer
  `PtgAttrIfError` token. This means **`expr` is evaluated twice** (once
  inside `ISERROR`, once as the else-branch) if the condition is false.
  For pure arithmetic on cell values (the intended use — e.g.
  `IFERROR(A1/B1, "")`) this is invisible and harmless; avoid relying on
  `IFERROR` around anything with side effects (not a concern for anything
  this crate itself generates, since formulas are pure expressions, but
  worth knowing if you're composing complex nested formulas).

### A worked formula example — variance columns

This is the pattern used throughout `examples/sales_report.rs`: a
"VAR (U)" (absolute change) and "VAR (%)" (percent change) column,
comparing two other cells, written with proper `""`-on-zero handling:

```rust
use xlsb_write::Formula;

let curr_col = 5;
let prev_col = 4;

// VAR (U) = IF(curr=0, "", curr-prev)
let var_u = Formula::if_then_else(
    Formula::cell(row, curr_col).eq(Formula::num(0.0)),
    Formula::str(""),
    Formula::cell(row, curr_col).sub(Formula::cell(row, prev_col)),
);
// Cached value must match what the formula computes — including the
// "" branch when curr is genuinely zero:
match if curr == 0.0 { None } else { Some(curr - prev) } {
    Some(v) => sheet.write_formula_num_with_format(row, var_u_col, var_u, v, &value_fmt),
    None    => sheet.write_formula_str_with_format(row, var_u_col, var_u, "", &value_fmt),
};

// VAR (%) = IFERROR(IF(curr=0,"",curr/prev-1), "")
let inner_if = Formula::if_then_else(
    Formula::cell(row, curr_col).eq(Formula::num(0.0)),
    Formula::str(""),
    Formula::cell(row, curr_col).div(Formula::cell(row, prev_col)).sub(Formula::num(1.0)),
);
let var_pct = Formula::iferror(inner_if, Formula::str(""));
match if curr == 0.0 || prev == 0.0 { None } else { Some(curr / prev - 1.0) } {
    Some(v) => sheet.write_formula_num_with_format(row, var_pct_col, var_pct, v, &pct_fmt),
    None    => sheet.write_formula_str_with_format(row, var_pct_col, var_pct, "", &pct_fmt),
};
```

Note the pattern: **whenever a formula's result could be a string in one
branch and a number in another** (like `""` vs. a computed number here),
you have to pick which `write_formula_*` call to use *in Rust*, based on
which branch actually fires for this row's data — the cell's on-disk type
(`BrtFmlaNum` vs `BrtFmlaString`) is determined by which cached value you
hand it, not inferred from the formula text.

## Full worked example

See [`examples/sales_report.rs`](examples/sales_report.rs) for a complete,
runnable, from-scratch example exercising nearly everything in this guide
at once: synthetic data generation, two sheets, colored/bold headers,
borders, column widths, frozen panes (rows and columns), a merged banner
cell, and a full SUM/IF/IFERROR variance table.

See [`examples/formula_functions_demo.rs`](examples/formula_functions_demo.rs)
for a runnable demo of every function added in the 2026-09-13
formula-coverage expansion (`&`/`CONCATENATE`, `AND`/`OR`/`NOT`,
`SUMIF`/`COUNTIF` including the 3-arg form, `LEFT`/`RIGHT`/`MID`/`LEN`/
`TEXT`, `TODAY`/`NOW`/`DATE`, `VLOOKUP`/`INDEX`/`MATCH`), each with a
hand-computed cached value next to it.

```
cargo run --example sales_report
```

produces `sales_report.xlsb` — open it in Excel to see the result, or run

```
cargo run --example validate_xlsb -- sales_report.xlsb
```

to structurally verify it (checks every binary part parses as well-formed
BIFF12, then opens it with `calamine` as an independent value-level check)
without needing Excel installed at all.

Minimal end-to-end example, for reference:

```rust
use xlsb_write::{Workbook, Format, Color, HAlign, VAlign};

fn main() -> Result<(), xlsb_write::WriteError> {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");

    let header = Format::new()
        .set_bold()
        .set_font_color(Color::WHITE)
        .set_background_color(Color::rgb(31, 73, 125))
        .set_align(HAlign::Center, VAlign::Center);

    sheet.write_string_with_format(0, 0, "Name", &header);
    sheet.write_string_with_format(0, 1, "Score", &header);
    sheet.set_freeze_panes(1);

    sheet.write_string(1, 0, "Alpha");
    sheet.write_number(1, 1, 91.5);
    sheet.write_string(2, 0, "Beta");
    sheet.write_number(2, 1, 88.0);

    wb.save("out.xlsb")
}
```

## Error handling

Every fallible operation returns `Result<_, WriteError>`:

```rust
pub enum WriteError {
    Io(std::io::Error),
    Zip(zip::result::ZipError),
}
```

It implements `std::error::Error` + `Display` (via `{:?}`), and has
`From<std::io::Error>`/`From<zip::result::ZipError>` impls, so `?` works
naturally in a function returning `Result<_, WriteError>` or
`Result<_, Box<dyn std::error::Error>>`.

The only *non*-`Result` failure mode is the **row-ordering panic** described
in [Writing cells](#writing-cells) — that's treated as a programming error
(you sorted your data wrong), not a recoverable I/O condition, so it's a
panic rather than an `Err`.

## Performance model

- **Formats are deduplicated per-sheet, then registered workbook-wide.**
  Passing the same `Format` (by value-equality, not pointer-equality) to
  thousands or millions of cells costs one hash-map lookup per cell and one
  style-table entry total — not one entry per cell.
- **Strings are deduplicated per-sheet** the same way, into Excel's shared
  string table (`xl/sharedStrings.bin`), which is also how Excel itself
  expects repeated text to be stored.
- **Rows are streamed to an internal buffer as soon as they're complete**
  (see the row-ordering rule above) — a `Worksheet` never holds a
  `HashMap`/`BTreeMap` of "every cell ever written" in memory; it holds one
  pending row plus the already-encoded bytes of every prior row.
- **`Workbook` vs `StreamingWorkbook`** is the one memory decision left to
  you: `Workbook` holds every sheet's encoded bytes until you call
  `.write()`/`.save()`; `StreamingWorkbook` writes each sheet to the output
  the moment you call `finish_worksheet`, so peak memory tracks your
  largest single sheet rather than the sum of all sheets. See
  [Choosing `Workbook` vs `StreamingWorkbook`](#choosing-workbook-vs-streamingworkbook).

None of this requires any action from you beyond picking `Workbook` vs
`StreamingWorkbook` — the string/format interning is automatic and always
on.

## Correctness: how this crate is verified

This crate's test suite (`cargo test`) checks correctness at three
independent levels, and all of it runs without needing Excel installed:

1. **Structural**: every `.bin` part of a generated file is parsed back as
   a well-formed BIFF12 record stream (`biff12::try_parse_records` /
   `examples/validate_xlsb.rs`) — this is the same class of check that
   would catch a truncated record or a length field that doesn't match
   actual content (the kind of bug that makes Excel silently "repair" or
   discard a worksheet part on open).
2. **Value-level, via an independent reader**: `tests/roundtrip.rs` and
   `tests/robustness.rs` open generated files with
   [`calamine`](https://crates.io/crates/calamine) — a completely separate
   Rust XLSB parser, sharing no code with this crate — and assert the
   values, formulas' cached results, and structural properties (row/column
   counts, blank cells, hidden columns, merges) read back correctly.
3. **Randomized shape coverage**: `tests/robustness.rs`'s
   `random_matrix_shapes_survive_generic_write_and_reopen` runs every
   existing feature (formulas, formatting, freeze panes, hidden columns,
   merges, both `Workbook` and `StreamingWorkbook`) against a fixed,
   deliberately tricky matrix of row/column-count/type-mix combinations
   (generated by `cargo run --example gen_random_datasets`, seeded so any
   failure is reproducible) — not a fuzzer, but broader coverage than any
   one hand-written test case.

Beyond the automated suite, output has also been verified by opening real
generated files in **actual Excel via COM automation** — checking the file
opens without a repair prompt and that formula cells recalculate to the
correct values live. If you want to do the same locally on Windows with
Excel installed, `examples/validate_xlsb.rs` is the quickest automated
proxy for "will Excel accept this file," and PowerShell's
`New-Object -ComObject Excel.Application` is the way to actually drive
Excel itself if you want the real thing.

**Embedded images specifically** (`embed_image`, 2026-09-13): this
feature's entire OPC/drawing-XML shape was derived from a real Excel-
authored reference, not the published spec text alone — the MS-XLSB HTML
pages don't render the actual ABNF grammar file that would say where
`BrtDrawing` belongs in the worksheet record stream. This session built a
`.xlsb` from scratch with real Excel (COM automation: `Shapes.AddPicture`
into a cell range, `SaveAs` format 50), inspected the produced zip
directly (`xl/drawings/drawing1.xml`'s exact shape, both `_rels` files,
`[Content_Types].xml`'s entries, and `BrtDrawing`'s exact byte position in
`sheet1.bin` via `examples/dump_sheet.rs`), and confirmed this crate's own
output opens with no repair prompt, reports the expected `Shapes.Count`
and anchor cells (`Shape.TopLeftCell`/`BottomRightCell`), and genuinely
resizes when the anchor range's column width/row height changes
(distinguishing a real two-cell anchor from a fixed-size picture that
merely starts at the right cell) — across all three worksheet types
(`Workbook`, `StreamingWorkbook`, `SizedStreamingWorksheet`).

**Autofilter and defined names specifically** (`set_autofilter`/
`define_name`, 2026-09-13): both features' exact record shape and position
were derived the same real-Excel-reference way, not from spec text alone —
the published MS-XLSB HTML pages don't render the ABNF grammar that would
say where `BrtBeginAFilter`/`BrtEndAFilter` belong in the worksheet record
stream, or that `BrtName`'s range is a full `PtgArea3d` Rgce token stream
rather than a direct row/col field. This session built several reference
files with real Excel (COM automation: `Range.AutoFilter`, `Workbook.Names.Add`)
varying the things most likely to matter — filter range not starting at
row 0, a multi-sheet workbook with the filter/name on a non-first sheet,
autofilter combined with an embedded image in the same sheet — and
inspected each with `examples/dump_sheet.rs`. Confirmed via COM automation
on this crate's own output across all three worksheet types (`Workbook`,
`StreamingWorkbook`, `SizedStreamingWorksheet`): the file opens with no
repair prompt, `Worksheet.AutoFilterMode` is `true` with
`AutoFilter.Range.Address` matching exactly, and `Workbook.Names` resolves
each defined name's `RefersToRange` to the correct sheet and address —
including a name pointing at a sheet other than the first, which exercises
this crate's dynamically-grown `BrtExternSheet` XTI table, not just the
single hardcoded entry every workbook already had. Defined names are also
covered by an independent-reader check in `tests/roundtrip.rs`: `calamine`
exposes `Reader::defined_names()` and decodes this crate's `BrtName`/
`BrtExternSheet` output back into `"Sheet!$A$1:$B$2"`-style strings,
matching exactly. Autofilter has no `calamine`-level check — like embedded
images, `calamine` is a cell-data reader and doesn't surface autofilter at
all — so `tests/autofilter.rs` checks the structural shape (the exact
`BrtBeginAFilter` payload and position) instead, and the real-Excel-COM
check above is what actually proves correctness for that feature.

**Formula-coverage expansion specifically** (`&`/`CONCATENATE`, `AND`/`OR`/
`NOT`, `SUMIF`/`COUNTIF`, `LEFT`/`RIGHT`/`MID`/`LEN`/`TEXT`, `TODAY`/`NOW`/
`DATE`, `VLOOKUP`/`INDEX`/`MATCH`, 2026-09-13): every `Ftab` index was
cross-checked against Apache POI's `functionMetadata.txt` (a
long-established, independently-maintained mapping of Excel's built-in
function table) and then confirmed byte-for-byte against a real
Excel-authored `.xlsb` — this session built one via COM automation
(`Range.Formula = "=AND(...)"` etc. for all eighteen new functions in one
pass, `SaveAs` format 50) and disassembled every formula cell's raw `Rgce`
bytes with `examples/dump_sheet.rs`/`examples/rgce_disasm.rs` (the latter's
own `PtgStr` decoding was found to still have the stale 1-byte-cch BIFF8
shape from before this crate's own `PtgStr` fix, and was corrected in the
same pass). This is how three non-obvious facts were discovered, none of
which could have been safely guessed from the spec or from argument count
alone: `PtgConcat` (`&`) is a plain binary infix Ptg, no different in shape
from `PtgAdd`; some functions (`NOT`, `COUNTIF`, `MID`, `LEN`, `TEXT`,
`DATE`, `TODAY`, `NOW`) are fixed-argument-count in real Excel and use the
narrower `PtgFunc` form instead of `PtgFuncVar` (`COUNTIF` and `SUMIF` both
commonly take 2 args, but only `COUNTIF`'s is fixed — this really did
require checking each function's actual bytes, not just picking one and
assuming the rest matched); and `TODAY`/`NOW` are volatile, requiring both
a `PtgAttrSemi` wrapper token and a cell-record `grbitFlags` bit this crate
would otherwise never emit (without which Excel has no reason to ever
recalculate a formula with no cell dependencies). `SUMIF`'s 3-arg form and
`VLOOKUP`'s `PtgBool` argument were each separately confirmed against their
own real-Excel reference formula, not assumed from the 2-arg/plain-value
cases. Every function is also covered by `tests/roundtrip.rs`'s
`roundtrip_new_formula_functions_through_calamine`/
`roundtrip_volatile_functions_through_calamine` (independent-reader
value-level checks) and was verified via COM automation on the actual
`.xlsb` this crate produces (`examples/formula_functions_demo.rs`): opens
with no repair prompt, every formula bar shows the expected formula text,
and Excel's own live recalculation matches every hand-computed value
except one locale-specific display quirk — `TEXT(value,"$0.00")` rendered
oddly on the Spanish-locale Excel installation used for this check (a
general Excel behavior where number-format codes are interpreted using the
current regional decimal/thousands-separator settings, not something this
crate's byte encoding controls; the underlying `PtgStr("$0.00")` bytes
match real Excel's own output exactly, confirmed separately by direct
disassembly).

## Limitations — what's not supported

This crate covers "data + formatting + formulas across one or many
sheets" — the common case for generating reports. It deliberately does
**not** implement:

- **Reading `.xlsb` files.** This is a writer only. (For reading, see
  [`calamine`](https://crates.io/crates/calamine) or
  [`pyxlsb`](https://github.com/willtrnr/pyxlsb) for Python.)
- **Print areas** and other `_xlnm`-prefixed built-in names. Ordinary
  workbook-scoped named ranges *are* supported — see
  [Defined names (named ranges)](#defined-names-named-ranges).
- **Autofilter column criteria** (pre-selecting which values a dropdown
  filters to) — `set_autofilter` only turns on the dropdowns themselves,
  see [Autofilter](#autofilter).
- **Charts, embedded objects (OLE, ActiveX), image resizing/cropping/rotation,
  image formats other than PNG/JPEG.** Embedding a plain PNG/JPEG image
  anchored to a cell range *is* supported — see
  [Embedding images](#embedding-images).
- **Conditional formatting, data validation.**
- **Hyperlinks.**
- **Pivot tables.**
- **Cell comments/notes.**
- **Per-side borders or diagonal borders** — `Format::set_border` applies
  one style to all four sides.
- **No text-formula parser** — there's no way to pass `"=VLOOKUP(...)"` as
  a string; formulas are always built via `Formula`'s method calls. Any
  `Ftab` function can be referenced via `FnIndex(raw_index)` (see
  [Formulas](#formulas)), but only the listed constants are verified
  against real Excel output by this crate's own tests.
- **Shared formulas / array formulas** — every formula cell is written as
  its own independent `BrtFmlaNum`/`BrtFmlaString` record.
- **Relative cell references in formulas** (`$`-style semantics) — every
  cell/range reference `Formula` encodes is absolute; this only matters if
  you were hoping to copy/paste a written formula around in Excel and have
  it auto-adjust — formulas this crate writes are otherwise fully normal
  and recalculate correctly in place.

If you need any of these, this crate isn't the right tool yet — consider
building the workbook with a full OOXML library, or via Excel automation,
for that specific need.

## Troubleshooting

**Panic: "cells must be written in strictly increasing row order"** — see
[Writing cells](#writing-cells). Sort your source data by row before
writing. This applies per-sheet; different sheets are entirely independent
(no shared row-ordering constraint across sheets).

**A formula cell shows a stale/wrong value when opened by a non-Excel
tool** — check that `cached_value` in `write_formula_num`/
`write_formula_str` actually matches what the formula computes for that
row's data. Excel recalculates on open and self-corrects; other readers
(including `calamine`, and this crate's own test suite) do not evaluate
formulas — they only ever see the cached value you supplied.

**A `SUM` formula reads back with a spurious `@` in front of it** — you
likely hand-built a `Formula::Func(FnIndex::SUM, vec![...])` for a single
range instead of using `Formula::sum_range(...)`. Use the dedicated
constructor; see the [note under Formulas](#built-in-functions-fnindex).

**Wide/merged/frozen layout doesn't look right** — remember
`merge_range` only affects the *visual* merge; write the value into the
merge's top-left cell separately with a normal `write_*` call.
`set_freeze_panes` (rows) and `set_freeze_panes_cols` (columns) are
independent settings — set both if you want a frozen header row *and* a
frozen first column simultaneously.

**File opens but a column/row I set width/height on doesn't reflect it** —
column/row specs (`set_column_width`, `set_column_hidden`,
`set_row_height`, `set_row_hidden`) can be called any time before the
sheet is finished, but make sure you're calling them on the right `col`/
`row` index (zero-based, same indexing as cell writes).

## Internals, for contributors

The crate is organized as:

- `src/lib.rs` — public API surface: `Workbook`, `StreamingWorkbook`,
  `Worksheet`, `CellValue`, `WriteError`, and the OPC/ZIP package assembly
  (`[Content_Types].xml`, relationships, `docProps`).
- `src/biff12.rs` — low-level BIFF12 primitives: varint encoding, generic
  record read/write, per-cell-type record encoders (`write_cell_rk`,
  `write_cell_isst`, ...), the RK-encoding heuristic for compact numbers.
- `src/sheet.rs` — worksheet binary framing: the sheet header (view state,
  freeze panes, column info), row encoding, and the sheet footer (merge
  cells, the fixed reference-derived tail bytes).
- `src/sst.rs` — the shared string table builder (`IndexMap<String, u32>`
  for O(1) lookup with preserved insertion order, which `BrtSst`'s
  index-ordering requirement needs).
- `src/drawing.rs` — OPC parts for embedded images: `xl/media/imageN.<ext>`,
  `xl/drawings/drawingN.xml` (regular DrawingML XML, per [MS-XLSB]'s own
  scope — even inside an otherwise-binary `.xlsb`) and its `_rels`, and the
  worksheet's own `_rels` file. The one binary-side hook (`BrtDrawing`) is
  encoded in `src/sheet.rs`'s `write_sheet_footer`, not here.
- `src/styles.rs` — `xl/styles.bin`: splices custom fonts/fills/borders/
  number-formats/cell-XFs onto a byte-verified reference base blob at
  dynamically-found offsets (not hardcoded — stays correct if the base
  blob ever changes).
- `src/formula.rs` — the `Formula` builder and its Rgce/Ptg token encoder.
- `src/wb_part.rs` — `xl/workbook.bin`: patches a fixed prefix/suffix
  template with one `BrtBundleSh` record per sheet, plus (when
  `Workbook::define_name`/`StreamingWorkbook::define_name` is used) a
  dynamically-grown `BrtExternSheet` XTI table and one `BrtName` record per
  defined name.

Dev tools in `examples/` for working on the binary format itself:

- `dump_sheet.rs` — parse a raw worksheet `.bin` and print its BIFF12
  records one per line, for diffing two files record-by-record.
- `dump_styles.rs` — same, for the styles.bin reference blob.
- `dump_schema.rs` — print a Parquet file's column list (used when
  building test fixtures from external data).
- `rgce_disasm.rs` — disassemble a raw formula token stream from a hex
  string, for checking `PtgAttrIf`/`PtgAttrGoto` offset math by hand.
- `validate_xlsb.rs` — the structural + calamine validator described above.
- `gen_random_datasets.rs` — generates the fixed, seeded matrix of
  randomly-shaped/typed Parquet fixtures `tests/robustness.rs` runs
  against.

Started from [`xlsb-writer`](https://crates.io/crates/xlsb-writer) by
kotucha (MIT) — see [README.md](README.md#acknowledgments) for the full
attribution and what's changed since.
