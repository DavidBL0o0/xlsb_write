//! Demonstrates `StreamingWorkbook::new_worksheet_sized` — the caller
//! declares the sheet's exact row/col extent upfront, so the header
//! (including a real `BrtWsDim`) can go out before any row data exists, and
//! every row streams straight to the zip entry with no whole-sheet buffer
//! (see `docs/superpowers/plans/2026-09-13-round6-streaming-and-parallel-write.md`).
//!
//! Builds one large, fully synthetic "Sales Ledger" sheet: one row per
//! (region, product) transaction, with a `Revenue = Units * Price` formula
//! column, formatted headers, frozen header row, and set column widths —
//! all set up before the first row is written, as `new_worksheet_sized`
//! requires. Row count defaults to a small demo size; pass a number on the
//! command line to generate more rows (used to re-measure peak RSS on a
//! large synthetic dataset — see the Round 6 plan's Task 6).
//!
//! Run with: cargo run --release --example sized_streaming_report [rows]

use std::fs::File;
use xlsb_write::{BorderStyle, Color, Format, Formula, HAlign, StreamingWorkbook, VAlign};

const HEADER_BLUE: Color = Color::rgb(31, 73, 125);
const REGIONS: &[&str] = &["North", "South", "East", "West"];
const PRODUCTS: &[&str] = &["Widget", "Gadget", "Gizmo", "Doohickey"];

/// Small deterministic xorshift PRNG — same approach as `sales_report.rs`,
/// no RNG crate needed for a demo/measurement dataset.
struct Rng(u64);
impl Rng {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }
}

fn main() {
    let rows: u32 = std::env::args().nth(1).and_then(|s| s.parse().ok()).unwrap_or(2_000);
    // Row 0 is the header; data occupies rows 1..=rows, so the declared
    // extent's last row is `rows` itself.
    let last_row = rows;
    let last_col = 4u32; // Region, Product, Units, Price, Revenue -> cols 0..=4

    let out_path = "sized_streaming_report.xlsb";
    let out_file = File::create(out_path).expect("create output file");
    let mut wb = StreamingWorkbook::create(out_file);

    let header_fmt = Format::new()
        .set_bold()
        .set_font_color(Color::WHITE)
        .set_background_color(HEADER_BLUE)
        .set_align(HAlign::Center, VAlign::Center)
        .set_border(BorderStyle::Thin)
        .set_border_color(Color::BLACK);
    let int_fmt = Format::new().set_num_format("int");
    let money_fmt = Format::new().set_num_format("float2");

    let mut sheet = wb.new_worksheet_sized("Sales Ledger", last_row, last_col);

    // Layout calls must happen before the first row is flushed (i.e. before
    // a second row starts, since `stage_cell` flushes the *previous* row) —
    // set them all up front, right after creating the sheet, exactly like
    // any other `Worksheet`/`StreamingWorksheet`. Calling these after the
    // header has gone out would panic (see `SizedStreamingWorksheet`'s doc
    // comment).
    sheet.set_freeze_panes(1);
    for col in 0..=last_col {
        sheet.set_column_width(col, 14.0);
    }

    let headers = ["Region", "Product", "Units", "Price", "Revenue"];
    for (i, name) in headers.iter().enumerate() {
        sheet.write_string_with_format(0, i as u32, name, &header_fmt).expect("write header cell");
    }

    let mut rng = Rng(42);
    for row in 1..=rows {
        let region = REGIONS[(rng.next() % REGIONS.len() as u64) as usize];
        let product = PRODUCTS[(rng.next() % PRODUCTS.len() as u64) as usize];
        let units = 1.0 + (rng.next() % 500) as f64;
        let price = 5.0 + (rng.next() % 200) as f64 / 4.0;
        let revenue = units * price;

        sheet.write_string(row, 0, region).expect("write region");
        sheet.write_string(row, 1, product).expect("write product");
        sheet.write_number_with_format(row, 2, units, &int_fmt).expect("write units");
        sheet.write_number_with_format(row, 3, price, &money_fmt).expect("write price");
        let formula = Formula::cell(row, 2).mul(Formula::cell(row, 3));
        sheet.write_formula_num_with_format(row, 4, formula, revenue, &money_fmt).expect("write revenue formula");
    }

    sheet.finish().expect("finish Sales Ledger sheet");
    wb.finish().expect("finish xlsb");
    println!("Wrote {out_path} ({rows} data rows, declared extent 0..={last_row} x 0..={last_col})");
}
