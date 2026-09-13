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

// arithmetic (all consume self, return a new Formula — chain freely)
a.add(b)   // a + b
a.sub(b)   // a - b
a.mul(b)   // a * b
a.div(b)   // a / b

// comparisons
a.lt(b)  a.le(b)  a.eq(b)  a.ge(b)  a.gt(b)  a.ne(b)

// SUM over a range — the common case gets a shortcut:
Formula::sum_range(first_row, first_col, last_row, last_col)
// equivalent to (but use the shortcut — see note below):
// Formula::Func(FnIndex::SUM, vec![Formula::range(...)])

// branching
Formula::if_then_else(cond, then, else_)              // IF(cond, then, else)
Formula::iferror(expr, default)                       // IFERROR(expr, default)
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
```

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
FnIndex::SUM, FnIndex::AVERAGE, FnIndex::MIN, FnIndex::MAX, FnIndex::ROUND
```

These are the only function indices this crate's own test suite has
verified byte-for-byte against real Excel output — use one of them via
`Formula::Func(FnIndex::SUM, args)` whenever possible.

For a function not listed here (`VLOOKUP`, `SUMIF`, text/date functions,
...), `FnIndex`'s inner value is public: `Formula::Func(FnIndex(0x0182),
vec![...])` works without forking this crate. This is **unverified by
this crate** — look the real index up yourself in the published MS-XLS
`Ftab` enumeration (don't guess), and double-check the argument count you
pass matches what the function actually requires (`PtgFuncVar`'s
`cparams` byte is written from `args.len()` — a mismatch there is exactly
the class of subtle bug this crate has hit before with its own built-in
functions).

**Use `Formula::sum_range` for a single-range `SUM`, not a hand-built
`Formula::Func`.** A single-range `SUM` is the one case real Excel encodes
with a dedicated shortcut token (`PtgAttrSum`) rather than the general
function-call form — `sum_range` produces exactly that byte-for-byte match
against real Excel output. Using the general form is spec-legal and opens
fine, but was found (during this crate's own hardening) to make Excel's
dynamic-array engine insert a spurious `@` into the formula on load. Stick
to the provided constructor and you won't hit this.

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

## Limitations — what's not supported

This crate covers "data + formatting + formulas across one or many
sheets" — the common case for generating reports. It deliberately does
**not** implement:

- **Reading `.xlsb` files.** This is a writer only. (For reading, see
  [`calamine`](https://crates.io/crates/calamine) or
  [`pyxlsb`](https://github.com/willtrnr/pyxlsb) for Python.)
- **Named ranges, print areas.**
- **Charts, images, embedded objects.**
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
- `src/styles.rs` — `xl/styles.bin`: splices custom fonts/fills/borders/
  number-formats/cell-XFs onto a byte-verified reference base blob at
  dynamically-found offsets (not hardcoded — stays correct if the base
  blob ever changes).
- `src/formula.rs` — the `Formula` builder and its Rgce/Ptg token encoder.
- `src/wb_part.rs` — `xl/workbook.bin`: patches a fixed prefix/suffix
  template with one `BrtBundleSh` record per sheet.

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
