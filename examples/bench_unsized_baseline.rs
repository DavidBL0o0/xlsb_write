//! Dev tool / measurement baseline: builds the exact same synthetic "Sales
//! Ledger" dataset as `sized_streaming_report.rs`, but through the older,
//! still-supported `StreamingWorkbook::new_worksheet` (unsized) path, which
//! buffers the whole sheet's encoded bytes in `Worksheet.body` until
//! `finish_worksheet` writes them out. Exists so the Round 6 "before" and
//! "after" peak-RSS numbers can be measured on an apples-to-apples dataset
//! (same row/col counts, same values) instead of comparing across
//! differently-shaped runs — see
//! `docs/superpowers/plans/2026-09-13-round6-streaming-and-parallel-write.md`.
//!
//! Run with: cargo run --release --example bench_unsized_baseline [rows]

use std::fs::File;
use xlsb_write::{BorderStyle, Color, Format, Formula, HAlign, StreamingWorkbook, VAlign};

const HEADER_BLUE: Color = Color::rgb(31, 73, 125);
const REGIONS: &[&str] = &["North", "South", "East", "West"];
const PRODUCTS: &[&str] = &["Widget", "Gadget", "Gizmo", "Doohickey"];

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

    let out_path = "bench_unsized_baseline.xlsb";
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

    let mut sheet = wb.new_worksheet("Sales Ledger");
    sheet.set_freeze_panes(1);
    for col in 0..=4u32 {
        sheet.set_column_width(col, 14.0);
    }

    let headers = ["Region", "Product", "Units", "Price", "Revenue"];
    for (i, name) in headers.iter().enumerate() {
        sheet.write_string_with_format(0, i as u32, name, &header_fmt);
    }

    let mut rng = Rng(42);
    for row in 1..=rows {
        let region = REGIONS[(rng.next() % REGIONS.len() as u64) as usize];
        let product = PRODUCTS[(rng.next() % PRODUCTS.len() as u64) as usize];
        let units = 1.0 + (rng.next() % 500) as f64;
        let price = 5.0 + (rng.next() % 200) as f64 / 4.0;
        let revenue = units * price;

        sheet.write_string(row, 0, region);
        sheet.write_string(row, 1, product);
        sheet.write_number_with_format(row, 2, units, &int_fmt);
        sheet.write_number_with_format(row, 3, price, &money_fmt);
        let formula = Formula::cell(row, 2).mul(Formula::cell(row, 3));
        sheet.write_formula_num_with_format(row, 4, formula, revenue, &money_fmt);
    }

    wb.finish_worksheet(sheet).expect("write Sales Ledger sheet");
    wb.finish().expect("finish xlsb");
    println!("Wrote {out_path} ({rows} data rows) via unsized new_worksheet (baseline)");
}
