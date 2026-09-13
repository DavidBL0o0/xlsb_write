//! End-to-end demo: builds a small, fully synthetic "sales report" workbook
//! exercising most of this crate's feature surface — colored/bold headers,
//! borders, alignment, column widths, frozen panes, merged cells, and
//! SUM/IF/IFERROR formulas — via `StreamingWorkbook`.
//!
//! Two sheets:
//! - "Raw Data": one row per (region, product, month) sale record.
//! - "Summary": a Region x Month pivot of Revenue, with a TOTAL column and
//!   month-over-month VAR (U)/VAR (%) formulas for the last two months.
//!
//! All data below is made up for this demo; it has no connection to any
//! real dataset.
//!
//! Run with: cargo run --example sales_report

use std::collections::BTreeMap;
use std::fs::File;
use xlsb_write::{BorderStyle, Color, Format, Formula, HAlign, StreamingWorkbook, VAlign};

const HEADER_BLUE: Color = Color::rgb(31, 73, 125);
const REGIONS: &[&str] = &["North", "South", "East", "West"];
const PRODUCTS: &[&str] = &["Widget", "Gadget", "Gizmo"];
const MONTHS: &[&str] = &["Jan", "Feb", "Mar", "Apr", "May", "Jun"];

struct Record {
    region: &'static str,
    product: &'static str,
    month: &'static str,
    units: i64,
    revenue: f64,
}

/// A small deterministic dataset (xorshift, no RNG crate needed for a demo)
/// — enough rows to make the Summary sheet's pivot/variance columns
/// meaningful, small enough to read at a glance.
fn synthetic_records() -> Vec<Record> {
    let mut seed: u64 = 42;
    let mut next = move || {
        seed ^= seed << 13;
        seed ^= seed >> 7;
        seed ^= seed << 17;
        seed
    };

    let mut records = Vec::new();
    for &region in REGIONS {
        for &product in PRODUCTS {
            for &month in MONTHS {
                let units = 50 + (next() % 450) as i64;
                let unit_price = 8.0 + (next() % 40) as f64 / 4.0;
                records.push(Record {
                    region,
                    product,
                    month,
                    units,
                    revenue: units as f64 * unit_price,
                });
            }
        }
    }
    records
}

fn write_raw_data_sheet(wb: &mut StreamingWorkbook<File>, records: &[Record]) {
    let header_fmt = Format::new()
        .set_bold()
        .set_font_color(Color::WHITE)
        .set_background_color(HEADER_BLUE)
        .set_align(HAlign::Center, VAlign::Center)
        .set_border(BorderStyle::Thin)
        .set_border_color(Color::BLACK);
    let int_fmt = Format::new().set_num_format("int");

    let mut sheet = wb.new_worksheet("Raw Data");
    let headers = ["Region", "Product", "Month", "Units", "Revenue"];
    for (i, name) in headers.iter().enumerate() {
        sheet.write_string_with_format(0, i as u32, name, &header_fmt);
        sheet.set_column_width(i as u32, 14.0);
    }
    sheet.set_freeze_panes(1);

    for (row, rec) in records.iter().enumerate() {
        let r = 1 + row as u32;
        sheet.write_string(r, 0, rec.region);
        sheet.write_string(r, 1, rec.product);
        sheet.write_string(r, 2, rec.month);
        sheet.write_number_with_format(r, 3, rec.units as f64, &int_fmt);
        sheet.write_number_with_format(r, 4, rec.revenue, &int_fmt);
    }

    wb.finish_worksheet(sheet).expect("write Raw Data sheet");
}

fn write_summary_sheet(wb: &mut StreamingWorkbook<File>, records: &[Record]) {
    // pivot[region][month] = total revenue
    let mut pivot: BTreeMap<&str, BTreeMap<&str, f64>> = BTreeMap::new();
    for rec in records {
        *pivot.entry(rec.region).or_default().entry(rec.month).or_insert(0.0) += rec.revenue;
    }

    let n_months = MONTHS.len();
    let col_total = 1 + n_months as u32;
    let col_var_u = col_total + 1;
    let col_var_pct = col_total + 2;
    let last_col = col_var_pct;

    let banner_fmt = Format::new()
        .set_bold()
        .set_font_size(12.0)
        .set_font_color(Color::WHITE)
        .set_background_color(HEADER_BLUE)
        .set_align(HAlign::Center, VAlign::Center);
    let header_fmt = Format::new()
        .set_bold()
        .set_font_color(Color::WHITE)
        .set_background_color(HEADER_BLUE)
        .set_align(HAlign::Center, VAlign::Center)
        .set_border(BorderStyle::Thin);
    let region_fmt = Format::new().set_border(BorderStyle::Thin);
    let value_fmt = Format::new().set_num_format("int").set_border(BorderStyle::Thin);
    let total_label_fmt = Format::new().set_bold().set_border(BorderStyle::Thin);
    let total_value_fmt = Format::new()
        .set_bold()
        .set_num_format("int")
        .set_border(BorderStyle::Thin);
    let pct_fmt = Format::new().set_num_format("pct2").set_border(BorderStyle::Thin);

    let mut sheet = wb.new_worksheet("Summary");
    sheet.set_row_height(0, 20.0);
    sheet.merge_range(0, 0, 0, last_col);
    sheet.write_string_with_format(0, 0, "Revenue by Region", &banner_fmt);

    sheet.write_string_with_format(1, 0, "Region", &header_fmt);
    sheet.set_column_width(0, 12.0);
    for (i, month) in MONTHS.iter().enumerate() {
        let col = 1 + i as u32;
        sheet.write_string_with_format(1, col, month, &header_fmt);
        sheet.set_column_width(col, 10.0);
    }
    sheet.write_string_with_format(1, col_total, "TOTAL", &header_fmt);
    sheet.write_string_with_format(1, col_var_u, "VAR (U)", &header_fmt);
    sheet.write_string_with_format(1, col_var_pct, "VAR (%)", &header_fmt);
    sheet.set_column_width(col_total, 12.0);
    sheet.set_column_width(col_var_u, 12.0);
    sheet.set_column_width(col_var_pct, 12.0);
    sheet.set_freeze_panes(2);
    sheet.set_freeze_panes_cols(1);

    let mut row = 2u32;
    let mut month_sums = vec![0.0f64; n_months];
    for &region in REGIONS {
        let by_month = &pivot[region];
        sheet.write_string_with_format(row, 0, region, &region_fmt);
        for (i, month) in MONTHS.iter().enumerate() {
            let v = by_month.get(month).copied().unwrap_or(0.0);
            sheet.write_number_with_format(row, 1 + i as u32, v, &value_fmt);
            month_sums[i] += v;
        }
        let total: f64 = MONTHS.iter().map(|m| by_month.get(m).copied().unwrap_or(0.0)).sum();
        sheet.write_formula_num_with_format(
            row,
            col_total,
            Formula::sum_range(row, 1, row, col_total - 1),
            total,
            &total_value_fmt,
        );

        let curr = *by_month.get(MONTHS[n_months - 1]).unwrap_or(&0.0);
        let prev = *by_month.get(MONTHS[n_months - 2]).unwrap_or(&0.0);
        let curr_col = col_total - 1;
        let prev_col = col_total - 2;
        // VAR (U) = IF(curr=0, "", curr-prev)
        let var_u = Formula::if_then_else(
            Formula::cell(row, curr_col).eq(Formula::num(0.0)),
            Formula::str(""),
            Formula::cell(row, curr_col).sub(Formula::cell(row, prev_col)),
        );
        match if curr == 0.0 { None } else { Some(curr - prev) } {
            Some(v) => sheet.write_formula_num_with_format(row, col_var_u, var_u, v, &value_fmt),
            None => sheet.write_formula_str_with_format(row, col_var_u, var_u, "", &value_fmt),
        };
        // VAR (%) = IFERROR(IF(curr=0,"",curr/prev-1), "")
        let inner_if = Formula::if_then_else(
            Formula::cell(row, curr_col).eq(Formula::num(0.0)),
            Formula::str(""),
            Formula::cell(row, curr_col)
                .div(Formula::cell(row, prev_col))
                .sub(Formula::num(1.0)),
        );
        let var_pct = Formula::iferror(inner_if, Formula::str(""));
        match if curr == 0.0 || prev == 0.0 {
            None
        } else {
            Some(curr / prev - 1.0)
        } {
            Some(v) => sheet.write_formula_num_with_format(row, col_var_pct, var_pct, v, &pct_fmt),
            None => sheet.write_formula_str_with_format(row, col_var_pct, var_pct, "", &pct_fmt),
        };
        row += 1;
    }

    let last_data_row = row - 1;
    let totals_row = row;
    sheet.write_string_with_format(totals_row, 0, "TOTAL", &total_label_fmt);
    for (i, &sum_value) in month_sums.iter().enumerate() {
        let col = 1 + i as u32;
        let sum = Formula::sum_range(2, col, last_data_row, col);
        sheet.write_formula_num_with_format(totals_row, col, sum, sum_value, &total_value_fmt);
    }
    let grand_total: f64 = month_sums.iter().sum();
    sheet.write_formula_num_with_format(
        totals_row,
        col_total,
        Formula::sum_range(2, col_total, last_data_row, col_total),
        grand_total,
        &total_value_fmt,
    );
    sheet.write_blank_with_format(totals_row, col_var_u, &total_label_fmt);
    sheet.write_blank_with_format(totals_row, col_var_pct, &total_label_fmt);

    wb.finish_worksheet(sheet).expect("write Summary sheet");
}

fn main() {
    let out_path = "sales_report.xlsb";
    let out_file = File::create(out_path).expect("create output file");
    let mut wb = StreamingWorkbook::create(out_file);

    let records = synthetic_records();
    write_raw_data_sheet(&mut wb, &records);
    write_summary_sheet(&mut wb, &records);

    wb.finish().expect("finish xlsb");
    println!("Wrote {out_path}");
}
