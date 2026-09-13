//! Demo: `Worksheet::set_autofilter` (per-sheet, turns a header row into
//! filterable dropdowns) and `Workbook::define_name` (workbook-level named
//! ranges, `Formulas > Define Name` in Excel).
//!
//! Two sheets:
//! - "Data": a small table with autofilter dropdowns on its header row.
//! - "Summary": references the first sheet only by name (open the file in
//!   Excel and check the Name Box dropdown, top-left of the formula bar).
//!
//! Run with: cargo run --example autofilter_and_names_demo
//! Then check the result structurally:
//! cargo run --example validate_xlsb -- autofilter_and_names_demo.xlsb

use xlsb_write::{Format, HAlign, VAlign, Workbook};

fn main() -> Result<(), xlsb_write::WriteError> {
    let mut wb = Workbook::new();

    let header = Format::new().set_bold().set_align(HAlign::Left, VAlign::Center);

    let data = wb.add_worksheet("Data");
    data.write_string_with_format(0, 0, "Name", &header);
    data.write_string_with_format(0, 1, "Qty", &header);
    data.write_string_with_format(0, 2, "Price", &header);
    let rows: &[(&str, f64, f64)] = &[("Alpha", 10.0, 9.99), ("Beta", 5.0, 19.99), ("Gamma", 20.0, 4.5)];
    for (i, &(name, qty, price)) in rows.iter().enumerate() {
        let row = i as u32 + 1;
        data.write_string(row, 0, name);
        data.write_number(row, 1, qty);
        data.write_number(row, 2, price);
    }
    // Turn the header + data rows into filterable dropdowns (A1:C4).
    data.set_autofilter(0, 0, rows.len() as u32, 2);

    let summary = wb.add_worksheet("Summary");
    summary.write_string_with_format(0, 0, "See the \"DataTable\" name in the Name Box \u{2192}", &header);

    // A workbook-level named range pointing at the Data sheet's table body
    // (sheet_index 0 == "Data", the first sheet added above) — named
    // ranges are workbook-scoped, which is why `define_name` lives on
    // `Workbook`, not `Worksheet`.
    wb.define_name("DataTable", 0, 0, 0, rows.len() as u32, 2);

    wb.save("autofilter_and_names_demo.xlsb")?;
    println!(
        "Wrote autofilter_and_names_demo.xlsb — open it in Excel: the \"Data\" sheet's \
         header row has filter dropdowns, and \"DataTable\" appears in the Name Box dropdown \
         (top-left of the formula bar) resolving to Data!$A$1:$C$4."
    );
    Ok(())
}
