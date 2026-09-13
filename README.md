# xlsb_write

Fast, pure-Rust writer for the Excel Binary (`.xlsb`) file format — no COM, no
Excel installation, no Python, no OS-specific APIs. Builds and runs the same
on Windows and Linux.

`.xlsb` is Excel's binary (BIFF12) workbook format: same feature set as
`.xlsx`, but smaller on disk and faster for Excel to open, which matters for
large reports. This crate writes it directly, encoding BIFF12 records
byte-for-byte rather than shelling out to Excel or converting from another
format.

## Features

- Cell types: strings (deduplicated via a shared string table), numbers,
  booleans, blanks, and formulas with a cached value for immediate display.
- Formulas: a small builder (`Formula::sum_range`, `.add()`, `.sub()`,
  `.eq()`, `if_then_else()`, `iferror()`, ...) rendered to native BIFF12
  Rgce tokens — no formula text parsing involved.
- Formatting: bold/italic/underline/strikeout, font color/size/name,
  background color, borders (style + color), horizontal/vertical alignment,
  text wrap, and number formats (built-in codes plus common shorthands).
- Layout: column width, hidden columns, row height, hidden rows, frozen
  panes, and merged cell ranges.
- Two write modes:
  - `Workbook` — build everything in memory, then `.save()` or `.write()`.
  - `StreamingWorkbook` — write one worksheet at a time straight to the
    output sink, for large reports where holding every sheet in memory
    isn't practical.

## Install

```toml
[dependencies]
xlsb_write = "0.1"
```

## Quick example

```rust
use xlsb_write::Workbook;

let mut wb = Workbook::new();
let sheet = wb.add_worksheet("Sheet1");
sheet.write_string(0, 0, "Name");
sheet.write_number(0, 1, 42.0);
wb.save("out.xlsb").unwrap();
```

See `examples/sales_report.rs` for a fuller demo — colored/bold headers,
borders, column widths, frozen panes, merged cells, and SUM/IF/IFERROR
formulas, using `StreamingWorkbook`:

```
cargo run --example sales_report
```

## Status

Byte-level output is verified against real Excel-produced reference files
and cross-checked with two independent readers ([`calamine`] and
[`pyxlsb`]'s record definitions) as part of the test suite. The public API
covers the common case (data + formatting + formulas across one or many
sheets) rather than the full `.xlsb`/[MS-XLSB] surface — named ranges, print
areas, charts, and conditional formatting aren't implemented yet.

[`calamine`]: https://crates.io/crates/calamine
[`pyxlsb`]: https://github.com/willtrnr/pyxlsb

## Acknowledgments

This crate started from [`xlsb-writer`](https://crates.io/crates/xlsb-writer)
by [kotucha](https://github.com/kotucha) (MIT License) — its byte-verified
BIFF12 framing was the starting point for the low-level encoding in
`src/biff12.rs`, `src/sheet.rs`, `src/sst.rs`, `src/wb_part.rs`, and
`src/styles.rs`. Everything built on top of that since — the imperative
cell-by-cell API, formulas, formatting, streaming writes, and the rest of
the public surface — is new. See [LICENSE](LICENSE) for both projects'
copyright notices.

## License

MIT — see [LICENSE](LICENSE).
