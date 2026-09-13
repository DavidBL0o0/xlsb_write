//! Randomized robustness suite: proves the writer works correctly across a
//! wide range of shapes, not just one hand-picked layout, by running every
//! existing feature (formulas, formatting, freeze panes, column
//! width/hidden, merges, both `Workbook` and `StreamingWorkbook`) against a
//! deliberately-chosen matrix of randomly-typed/shaped parquet fixtures
//! (`test_fixtures/random/`, produced by `cargo run --example
//! gen_random_datasets` with a fixed seed — regenerate them if this test
//! reports a missing-fixtures failure).
//!
//! Not a fuzzer: a fixed set of tricky shapes, run deterministically. Test
//! infrastructure only — no new public API is exercised beyond what
//! `tests/roundtrip.rs` already covers elsewhere.

use calamine::{open_workbook, Data, Reader, Xlsb};
use parquet::basic::Type as PhysicalType;
use parquet::file::reader::{FileReader, SerializedFileReader};
use parquet::record::Field;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use xlsb_write::biff12::try_parse_records;
use xlsb_write::{Color, Format, Formula, StreamingWorkbook, WriteError, Worksheet};
use zip::ZipArchive;

fn field_as_f64(f: &Field) -> Option<f64> {
    match f {
        Field::Byte(v) => Some(*v as f64),
        Field::Short(v) => Some(*v as f64),
        Field::Int(v) => Some(*v as f64),
        Field::Long(v) => Some(*v as f64),
        Field::UByte(v) => Some(*v as f64),
        Field::UShort(v) => Some(*v as f64),
        Field::UInt(v) => Some(*v as f64),
        Field::ULong(v) => Some(*v as f64),
        Field::Float(v) => Some(*v as f64),
        Field::Double(v) => Some(*v),
        _ => None,
    }
}

/// What `populate_generic_sheet` did, needed to check calamine's reopened
/// view actually matches what was written.
struct SheetShape {
    data_rows: u32,
    cols: u32,
    /// The first cell `write_blank` was called on, if any — used to check a
    /// null value really does read back blank, not just "some value".
    first_blank_cell: Option<(u32, u32)>,
    /// Whether a `SUM` row was written below the data (only when there's a
    /// numeric column and at least one data row) — extends the used range
    /// by one row beyond header + data, which the row-count check needs to
    /// know about to compute the right expectation.
    has_sum_row: bool,
}

/// Writes one sheet generically from `path`'s parquet schema/rows — a
/// header row (bold+filled, from column names), one row per parquet row
/// with each field mapped to the closest `CellValue`, plus every other
/// *existing* feature this suite is responsible for exercising: frozen
/// header, a column width, a hidden last column, a merged header pair, and
/// one `SUM` formula over the first numeric column.
fn populate_generic_sheet(sheet: &mut Worksheet, path: &Path) -> SheetShape {
    let file = std::fs::File::open(path).unwrap_or_else(|e| panic!("open {}: {e}", path.display()));
    let reader = SerializedFileReader::new(file).expect("open parquet reader");
    let schema = reader.metadata().file_metadata().schema_descr();
    let num_cols = schema.num_columns() as u32;
    let num_rows = reader.metadata().file_metadata().num_rows() as u32;

    let numeric_cols: Vec<u32> = (0..schema.num_columns())
        .filter(|&i| {
            matches!(
                schema.column(i).physical_type(),
                PhysicalType::INT32 | PhysicalType::INT64 | PhysicalType::FLOAT | PhysicalType::DOUBLE
            )
        })
        .map(|i| i as u32)
        .collect();
    let first_numeric_col = numeric_cols.first().copied();

    let header_fmt = Format::new().set_bold().set_background_color(Color::rgb(200, 200, 200));
    for i in 0..num_cols {
        sheet.write_string_with_format(0, i, &format!("col_{i}"), &header_fmt);
    }
    sheet.set_freeze_panes(1);
    sheet.set_column_width(0, 15.0);
    if num_cols >= 2 {
        sheet.set_column_hidden(num_cols - 1);
        sheet.merge_range(0, 0, 0, 1);
    }

    let mut row_idx = 1u32;
    let mut first_blank_cell = None;
    let mut running_sum = 0.0f64;
    let row_iter = reader.get_row_iter(None).expect("row iter");
    for row_result in row_iter {
        let row = row_result.expect("read row");
        for (c, (_name, field)) in row.get_column_iter().enumerate() {
            let col = c as u32;
            match field {
                Field::Null => {
                    sheet.write_blank(row_idx, col);
                    first_blank_cell.get_or_insert((row_idx, col));
                }
                Field::Bool(b) => {
                    sheet.write_boolean(row_idx, col, *b);
                }
                Field::Str(s) => {
                    sheet.write_string(row_idx, col, s);
                }
                other => {
                    if let Some(v) = field_as_f64(other) {
                        if Some(col) == first_numeric_col && v.is_finite() {
                            running_sum += v;
                        }
                        sheet.write_number(row_idx, col, v);
                    } else {
                        sheet.write_string(row_idx, col, &format!("{other}"));
                    }
                }
            }
        }
        row_idx += 1;
    }

    let has_sum_row = first_numeric_col.is_some() && num_rows >= 1;
    if let Some(fc) = first_numeric_col {
        if num_rows >= 1 {
            let sum_row = row_idx;
            let formula = Formula::sum_range(1, fc, row_idx - 1, fc);
            sheet.write_formula_num(sum_row, fc, formula, running_sum);
        }
    }

    SheetShape { data_rows: num_rows, cols: num_cols, first_blank_cell, has_sum_row }
}

fn is_blank(v: Option<&Data>) -> bool {
    matches!(v, None | Some(Data::Empty))
}

/// Builds one case's sheet, alternating `Workbook`/`StreamingWorkbook` by
/// `case_index` parity, writes the result to a temp file, and asserts every
/// structural/value-level check. Appends a labeled failure to `failures`
/// instead of panicking immediately, so one run reports every failing case,
/// not just the first.
fn check_case(path: &Path, case_index: usize, failures: &mut Vec<String>) {
    let label = path.file_name().unwrap().to_string_lossy().to_string();
    let bytes = if case_index % 2 == 0 {
        let mut wb = xlsb_write::Workbook::new();
        let sheet = wb.add_worksheet("Sheet1");
        let shape = populate_generic_sheet(sheet, path);
        let mut buf = Cursor::new(Vec::new());
        if let Err(e) = wb.write(&mut buf) {
            failures.push(format!("{label}: Workbook::write failed: {e}"));
            return;
        }
        (buf.into_inner(), shape)
    } else {
        let mut buf = Cursor::new(Vec::new());
        let result: Result<SheetShape, WriteError> = (|| {
            let mut wb = StreamingWorkbook::create(&mut buf);
            let mut sheet = wb.new_worksheet("Sheet1");
            let shape = populate_generic_sheet(&mut sheet, path);
            wb.finish_worksheet(sheet)?;
            wb.finish()?;
            Ok(shape)
        })();
        match result {
            Ok(shape) => (buf.into_inner(), shape),
            Err(e) => {
                failures.push(format!("{label}: StreamingWorkbook failed: {e}"));
                return;
            }
        }
    };
    let (bytes, shape) = bytes;

    // Structural check: every binary part is a well-formed BIFF12 record
    // stream (same method as examples/validate_xlsb.rs).
    match ZipArchive::new(Cursor::new(&bytes)) {
        Ok(mut zip) => {
            let names: Vec<String> = (0..zip.len()).map(|i| zip.by_index(i).unwrap().name().to_string()).collect();
            for name in &names {
                if !name.ends_with(".bin") {
                    continue;
                }
                let mut entry = zip.by_name(name).expect("reopen entry");
                let mut part_bytes = Vec::with_capacity(entry.size() as usize);
                std::io::Read::read_to_end(&mut entry, &mut part_bytes).expect("read entry");
                if let Err(e) = try_parse_records(&part_bytes) {
                    failures.push(format!("{label}: malformed BIFF12 in {name}: {e}"));
                }
            }
        }
        Err(e) => {
            failures.push(format!("{label}: not a valid zip: {e}"));
            return;
        }
    }

    // Value-level check: an independent reader (calamine) opens it and
    // reports a consistent shape.
    let tmp = std::env::temp_dir().join(format!("xlsb_write_robustness_{label}.xlsb"));
    std::fs::write(&tmp, &bytes).expect("write temp file");
    match open_workbook::<Xlsb<_>, _>(&tmp) {
        Ok(mut wbk) => match wbk.worksheet_range("Sheet1") {
            Ok(range) => {
                let (rows, cols) = range.get_size();
                if shape.data_rows > 0 {
                    let expected_rows = (shape.data_rows + 1 + shape.has_sum_row as u32) as usize; // + header (+ sum row)
                    if rows != expected_rows {
                        failures.push(format!(
                            "{label}: calamine reports {rows} rows, expected {expected_rows} (data_rows={} + header)",
                            shape.data_rows
                        ));
                    }
                }
                if cols as u32 != shape.cols {
                    failures.push(format!("{label}: calamine reports {cols} cols, expected {}", shape.cols));
                }
                if let Some((r, c)) = shape.first_blank_cell {
                    let v = range.get_value((r, c));
                    if !is_blank(v) {
                        failures.push(format!("{label}: cell ({r},{c}) written as blank, calamine read back {v:?}"));
                    }
                }
            }
            Err(e) => failures.push(format!("{label}: calamine could not read 'Sheet1': {e}")),
        },
        Err(e) => failures.push(format!("{label}: calamine could not open file: {e}")),
    }
}

fn random_case_files() -> Vec<PathBuf> {
    let dir = Path::new("test_fixtures/random");
    assert!(
        dir.is_dir(),
        "test_fixtures/random/ not found — run `cargo run --example gen_random_datasets` first to generate fixtures"
    );
    let mut files: Vec<PathBuf> = std::fs::read_dir(dir)
        .expect("read test_fixtures/random")
        .filter_map(|e| e.ok().map(|e| e.path()))
        .filter(|p| p.extension().is_some_and(|ext| ext == "parquet"))
        .collect();
    files.sort();
    assert!(!files.is_empty(), "test_fixtures/random/ exists but has no .parquet files — regenerate the fixtures");
    files
}

#[test]
fn random_matrix_shapes_survive_generic_write_and_reopen() {
    let files = random_case_files();
    let mut failures = Vec::new();
    for (i, path) in files.iter().enumerate() {
        check_case(path, i, &mut failures);
    }
    assert!(failures.is_empty(), "{} of {} random-shape cases failed:\n{}", failures.len(), files.len(), failures.join("\n"));
}
