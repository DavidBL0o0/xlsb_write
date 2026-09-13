//! Dev tool: writes a sheet via `new_worksheet_sized` where the declared
//! extent is deliberately far larger than what's actually written, to
//! empirically answer Round 6's flagged open question — is a `BrtWsDim`
//! that's a superset of the real content spec-legal/Excel-safe, or does
//! Excel's strict validator reject it the same way it rejects a too-small
//! dim ("Replaced Part")? Checked by opening the output in real Excel via
//! COM automation (see
//! `docs/superpowers/plans/2026-09-13-round6-streaming-and-parallel-write.md`).
//! Confirmed safe: opens with no repair dialog, no `%TEMP%\errorNNN.xml`
//! recovery log, and Excel computes its own (correct, tighter) UsedRange
//! from the actual content regardless of what `BrtWsDim` declared.
//!
//! Declares a 1001-row x 51-col extent but writes only 3 rows x 3 cols.
//!
//! Run with: cargo run --example verify_superset_dim
use std::fs::File;
use xlsb_write::StreamingWorkbook;

fn main() {
    let out_file = File::create("verify_superset_dim.xlsb").expect("create output file");
    let mut wb = StreamingWorkbook::create(out_file);

    let mut sheet = wb.new_worksheet_sized("Sheet1", 1000, 50);
    sheet.write_string(0, 0, "Only").unwrap();
    sheet.write_string(0, 1, "a").unwrap();
    sheet.write_string(0, 2, "few").unwrap();
    sheet.write_number(1, 0, 1.0).unwrap();
    sheet.write_number(1, 1, 2.0).unwrap();
    sheet.write_number(2, 0, 3.0).unwrap();
    sheet.finish().expect("finish sheet");

    wb.finish().expect("finish workbook");
    println!("Wrote verify_superset_dim.xlsb: declared extent 0..=1000 x 0..=50, actually wrote 3 rows x 3 cols");
}
