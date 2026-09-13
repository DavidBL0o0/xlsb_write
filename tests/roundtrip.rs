//! End-to-end round-trip: write a .xlsb with `xlsb_write`, then read it back
//! with `calamine` (an independent, unrelated XLSB parser) and assert the
//! values match. This is the real proof the output is a valid Excel file,
//! not just "produces a zip".

use calamine::{Data, DataType, Reader, Xlsb, open_workbook};
use std::io::Cursor;
use xlsb_write::{BorderStyle, Color, Format, Formula, HAlign, VAlign, Workbook};

fn write_sample() -> Vec<u8> {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.write_string(0, 0, "Name");
    sheet.write_string(0, 1, "Qty");
    sheet.write_string(0, 2, "Price");
    sheet.write_string(0, 3, "InStock");

    sheet.write_string(1, 0, "Alpha");
    sheet.write_number(1, 1, 100.0);
    sheet.write_number(1, 2, 9.99);
    sheet.write_boolean(1, 3, true);

    sheet.write_string(2, 0, "Beta");
    sheet.write_number(2, 1, 200.0);
    sheet.write_number(2, 2, 1234.56);
    sheet.write_boolean(2, 3, false);

    sheet.write_blank(3, 0);

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    buf.into_inner()
}

#[test]
fn roundtrip_values_through_calamine() {
    let bytes = write_sample();
    let path = std::env::temp_dir().join("xlsb_write_roundtrip_test.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let mut wb: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let range = wb.worksheet_range("Sheet1").expect("Sheet1 not found");

    assert_eq!(range.get_value((0, 0)), Some(&Data::String("Name".into())));
    assert_eq!(range.get_value((0, 1)), Some(&Data::String("Qty".into())));
    assert_eq!(range.get_value((0, 2)), Some(&Data::String("Price".into())));
    assert_eq!(range.get_value((0, 3)), Some(&Data::String("InStock".into())));

    // calamine surfaces whole numbers stored via RK-integer encoding as
    // `Data::Int`, and non-integer numbers as `Data::Float` — compare via
    // `as_f64()` instead of matching the exact enum variant.
    let as_f64 = |r: &calamine::Range<Data>, pos: (u32, u32)| r.get_value(pos).and_then(Data::as_f64);

    assert_eq!(range.get_value((1, 0)), Some(&Data::String("Alpha".into())));
    assert_eq!(as_f64(&range, (1, 1)), Some(100.0));
    assert_eq!(as_f64(&range, (1, 2)), Some(9.99));
    assert_eq!(range.get_value((1, 3)), Some(&Data::Bool(true)));

    assert_eq!(range.get_value((2, 0)), Some(&Data::String("Beta".into())));
    assert_eq!(as_f64(&range, (2, 1)), Some(200.0));
    assert_eq!(as_f64(&range, (2, 2)), Some(1234.56));
    assert_eq!(range.get_value((2, 3)), Some(&Data::Bool(false)));

    println!("Round-trip OK. File written to: {}", path.display());
}

/// Pins the streaming-write contract end-to-end: several rows, written in
/// increasing row order, mixing numbers and strings, each row flushed to
/// `Worksheet::body` before the next one starts — decode via `calamine` and
/// confirm every value survives.
#[test]
fn roundtrip_multi_row_streaming_through_calamine() {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    for row in 0..5u32 {
        sheet.write_string(row, 0, &format!("row-{row}"));
        sheet.write_number(row, 1, row as f64 * 10.0);
    }

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let bytes = buf.into_inner();
    let path = std::env::temp_dir().join("xlsb_write_roundtrip_streaming.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let mut wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let range = wbk.worksheet_range("Sheet1").expect("Sheet1 not found");
    for row in 0..5u32 {
        assert_eq!(range.get_value((row, 0)), Some(&Data::String(format!("row-{row}"))));
        assert_eq!(
            range.get_value((row, 1)).and_then(Data::as_f64),
            Some(row as f64 * 10.0)
        );
    }
}

/// Same idea, but every cell carries real formatting: bold header, fill
/// color, border, alignment, custom number format. This proves the
/// styles.bin splice pipeline (Fonts/Fills/Borders/CellXfs) doesn't corrupt
/// the file or the values calamine reads back — it does NOT prove the
/// colors/borders render correctly in real Excel (no Excel on this
/// machine to check visually; flagged as a known gap in the project plan).
#[test]
fn roundtrip_with_formatting_through_calamine() {
    let header = Format::new()
        .set_bold()
        .set_font_color(Color::WHITE)
        .set_background_color(Color::rgb(31, 73, 125))
        .set_align(HAlign::Center, VAlign::Center)
        .set_border(BorderStyle::Thin)
        .set_border_color(Color::BLACK);
    let money = Format::new().set_num_format("float2");
    let pct = Format::new().set_num_format("pct2");

    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Kitchen Sink");
    sheet.set_freeze_panes(1);
    sheet.write_string_with_format(0, 0, "Client", &header);
    sheet.write_string_with_format(0, 1, "Revenue", &header);
    sheet.write_string_with_format(0, 2, "Growth", &header);

    sheet.write_string(1, 0, "Alpha");
    sheet.write_number_with_format(1, 1, 12345.678, &money);
    sheet.write_number_with_format(1, 2, 0.1534, &pct);

    sheet.write_string(2, 0, "Beta");
    sheet.write_number_with_format(2, 1, 999.5, &money);
    sheet.write_number_with_format(2, 2, -0.02, &pct);

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let bytes = buf.into_inner();
    assert_eq!(&bytes[..2], b"PK");

    let path = std::env::temp_dir().join("xlsb_write_roundtrip_formatted.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let mut wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the formatted file");
    let range = wbk.worksheet_range("Kitchen Sink").expect("sheet not found");

    let as_f64 = |r: &calamine::Range<Data>, pos: (u32, u32)| r.get_value(pos).and_then(Data::as_f64);

    assert_eq!(range.get_value((0, 0)), Some(&Data::String("Client".into())));
    assert_eq!(range.get_value((0, 1)), Some(&Data::String("Revenue".into())));
    assert_eq!(range.get_value((0, 2)), Some(&Data::String("Growth".into())));

    assert_eq!(range.get_value((1, 0)), Some(&Data::String("Alpha".into())));
    assert_eq!(as_f64(&range, (1, 1)), Some(12345.678));
    assert_eq!(as_f64(&range, (1, 2)), Some(0.1534));

    assert_eq!(range.get_value((2, 0)), Some(&Data::String("Beta".into())));
    assert_eq!(as_f64(&range, (2, 1)), Some(999.5));
    assert_eq!(as_f64(&range, (2, 2)), Some(-0.02));

    println!("Formatted round-trip OK. File written to: {}", path.display());
}

/// Exercises every Phase-3 sheet feature (column width, hidden column,
/// row height, hidden row, merged cells, frozen rows+cols) in one sheet.
/// calamine doesn't expose column width/hidden/merge metadata for XLSB, so
/// this can't verify those render correctly — but it DOES prove the new
/// BrtColInfo/BrtBeginColInfos/BrtMergeCell records are framed correctly
/// (correct rid+length), since a single off-by-one there would desync the
/// record stream and corrupt every cell value read after it.
#[test]
fn roundtrip_with_sheet_features_through_calamine() {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Layout");
    sheet.set_freeze_panes(1);
    sheet.set_freeze_panes_cols(1);
    sheet.set_column_width(0, 20.0);
    sheet.set_column_hidden(3);
    sheet.set_row_height(0, 25.0);
    sheet.set_row_hidden(5);
    sheet.merge_range(0, 1, 0, 2);

    sheet.write_string(0, 0, "ID");
    sheet.write_string(0, 1, "Merged Header");
    sheet.write_string(1, 0, "A1");
    sheet.write_number(1, 1, 1.0);
    sheet.write_string(5, 0, "Hidden row");
    sheet.write_number(5, 1, 42.0);

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let bytes = buf.into_inner();
    assert_eq!(&bytes[..2], b"PK");

    let path = std::env::temp_dir().join("xlsb_write_roundtrip_sheet_features.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let mut wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let range = wbk.worksheet_range("Layout").expect("sheet not found");

    assert_eq!(range.get_value((0, 0)), Some(&Data::String("ID".into())));
    assert_eq!(range.get_value((0, 1)), Some(&Data::String("Merged Header".into())));
    assert_eq!(range.get_value((1, 0)), Some(&Data::String("A1".into())));
    assert_eq!(range.get_value((1, 1)).and_then(Data::as_f64), Some(1.0));
    assert_eq!(range.get_value((5, 0)), Some(&Data::String("Hidden row".into())));
    assert_eq!(range.get_value((5, 1)).and_then(Data::as_f64), Some(42.0));

    println!("Sheet-features round-trip OK. File written to: {}", path.display());
}

/// Writes SUM/arithmetic formulas and checks that calamine reads back the
/// CACHED value correctly (calamine, like Excel before its own
/// recalculation, displays the cached value — it doesn't evaluate the
/// formula token stream itself). This is the real proof the Rgce/Ptg
/// encoding is byte-correct: if a single token or offset were wrong, the
/// record framing would desync and everything after it would fail to
/// parse, not just silently show a wrong number.
#[test]
fn roundtrip_formulas_through_calamine() {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Formulas");

    // Values written in row-major order — the streaming writer requires
    // non-decreasing row order; the formulas/assertions are unchanged from
    // before this reordering.
    let total = 10.0 + 20.0 + 30.0;
    // (A1 + 5) * 2
    let expr = Formula::cell(0, 0).add(Formula::num(5.0)).mul(Formula::num(2.0));

    sheet.write_number(0, 0, 10.0);
    sheet.write_formula_num(0, 1, expr, (10.0 + 5.0) * 2.0);

    // Isolation cases to narrow down any remaining operator/precedence bug.
    sheet.write_number(1, 0, 20.0);
    sheet.write_formula_num(1, 1, Formula::cell(0, 0).add(Formula::num(5.0)), 10.0 + 5.0); // A1+5 = 15

    sheet.write_number(2, 0, 30.0);
    sheet.write_formula_num(2, 1, Formula::num(5.0).mul(Formula::num(2.0)), 5.0 * 2.0); // 5*2 = 10

    sheet.write_formula_num(3, 0, Formula::sum_range(0, 0, 2, 0), total);
    sheet.write_formula_num(
        3,
        1,
        Formula::num(2.0).mul(Formula::cell(0, 0).add(Formula::num(5.0))),
        2.0 * (10.0 + 5.0),
    ); // 2*(A1+5) = 30, operand order swapped vs the first case

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let bytes = buf.into_inner();
    assert_eq!(&bytes[..2], b"PK");

    let path = std::env::temp_dir().join("xlsb_write_roundtrip_formulas.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let mut wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let range = wbk.worksheet_range("Formulas").expect("sheet not found");

    assert_eq!(range.get_value((0, 0)).and_then(Data::as_f64), Some(10.0));
    assert_eq!(range.get_value((1, 0)).and_then(Data::as_f64), Some(20.0));
    assert_eq!(range.get_value((2, 0)).and_then(Data::as_f64), Some(30.0));
    assert_eq!(range.get_value((3, 0)).and_then(Data::as_f64), Some(total));
    assert_eq!(range.get_value((0, 1)).and_then(Data::as_f64), Some(30.0));

    println!("Formula round-trip OK. File written to: {}", path.display());
}

/// `IF`/`IFERROR` use branch tokens (`PtgAttrIf`/`PtgAttrGoto`) with
/// precise byte-offset jumps — the riskiest formula encoding in this
/// crate. Verify with cases matching the legacy pipeline's actual
/// formulas: `VAR (U) = IF(curr=0,"",curr-prev)` and a `IFERROR` division.
#[test]
fn roundtrip_if_iferror_through_calamine() {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("IfTest");

    // Values written in row-major order — the streaming writer requires
    // non-decreasing row order; the formulas/assertions are unchanged from
    // before this reordering.

    // curr=0 in row0 -> the legacy pipeline's VAR(U) = IF(curr=0,"",curr-prev)
    // should take the "" branch.
    sheet.write_number(0, 0, 0.0); // curr
    sheet.write_number(0, 1, 100.0); // prev
    let var_u_zero = Formula::if_then_else(
        Formula::cell(0, 0).eq(Formula::num(0.0)),
        Formula::str(""),
        Formula::cell(0, 0).sub(Formula::cell(0, 1)),
    );
    sheet.write_formula_str(0, 2, var_u_zero, "");

    // curr=10, prev=0 -> IFERROR(curr/prev, "") should take the "" branch
    // (division by zero -> #DIV/0! -> ISERROR true -> default).
    sheet.write_number(1, 0, 10.0);
    sheet.write_number(1, 1, 0.0);
    let iferr_div0 = Formula::iferror(Formula::cell(1, 0).div(Formula::cell(1, 1)), Formula::str(""));
    sheet.write_formula_str(1, 2, iferr_div0, "");

    // curr=10, prev=2 -> IFERROR(curr/prev, "") should take the value branch (5.0).
    sheet.write_number(2, 0, 10.0);
    sheet.write_number(2, 1, 2.0);
    let iferr_ok = Formula::iferror(Formula::cell(2, 0).div(Formula::cell(2, 1)), Formula::str(""));
    sheet.write_formula_num(2, 2, iferr_ok, 5.0);

    // curr=50, prev=20 -> VAR(U) should take the value branch (30).
    sheet.write_number(3, 0, 50.0);
    sheet.write_number(3, 1, 20.0);
    let var_u_nonzero = Formula::if_then_else(
        Formula::cell(3, 0).eq(Formula::num(0.0)),
        Formula::str(""),
        Formula::cell(3, 0).sub(Formula::cell(3, 1)),
    );
    sheet.write_formula_num(3, 2, var_u_nonzero, 30.0);

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let bytes = buf.into_inner();
    assert_eq!(&bytes[..2], b"PK");

    let path = std::env::temp_dir().join("xlsb_write_roundtrip_if.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let mut wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let range = wbk.worksheet_range("IfTest").expect("sheet not found");
    assert_eq!(range.get_value((0, 2)), Some(&Data::String("".into())));
    assert_eq!(range.get_value((3, 2)).and_then(Data::as_f64), Some(30.0));
    assert_eq!(range.get_value((1, 2)), Some(&Data::String("".into())));
    assert_eq!(range.get_value((2, 2)).and_then(Data::as_f64), Some(5.0));

    println!("IF/IFERROR round-trip OK. File written to: {}", path.display());
}

/// `StreamingWorkbook` writes each sheet to the output the moment
/// `finish_worksheet` is called, instead of holding every sheet in memory
/// until one terminal `write` call (see `Workbook`) — this proves the
/// resulting file is just as valid: multiple sheets, mixed value types,
/// read back correctly and with the right active/selected sheet.
#[test]
fn roundtrip_streaming_workbook_through_calamine() {
    use xlsb_write::StreamingWorkbook;

    let mut buf = Cursor::new(Vec::new());
    let mut wb = StreamingWorkbook::create(&mut buf);

    let mut sheet1 = wb.new_worksheet("First");
    for row in 0..3u32 {
        sheet1.write_string(row, 0, &format!("first-{row}"));
        sheet1.write_number(row, 1, row as f64);
    }
    wb.finish_worksheet(sheet1).unwrap();

    let mut sheet2 = wb.new_worksheet("Second");
    sheet2.write_boolean(0, 0, true);
    sheet2.write_boolean(1, 0, false);
    wb.finish_worksheet(sheet2).unwrap();

    wb.finish().unwrap();
    let bytes = buf.into_inner();
    assert_eq!(&bytes[..2], b"PK");

    let path = std::env::temp_dir().join("xlsb_write_roundtrip_streaming_workbook.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let mut wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let first = wbk.worksheet_range("First").expect("First sheet not found");
    for row in 0..3u32 {
        assert_eq!(first.get_value((row, 0)), Some(&Data::String(format!("first-{row}"))));
        assert_eq!(first.get_value((row, 1)).and_then(Data::as_f64), Some(row as f64));
    }
    let second = wbk.worksheet_range("Second").expect("Second sheet not found");
    assert_eq!(second.get_value((0, 0)), Some(&Data::Bool(true)));
    assert_eq!(second.get_value((1, 0)), Some(&Data::Bool(false)));

    println!("StreamingWorkbook round-trip OK. File written to: {}", path.display());
}

/// `new_worksheet_sized` streams every row straight to the zip entry (no
/// whole-sheet `body` buffer at all — see
/// `2026-09-13-round6-streaming-and-parallel-write.md`), which means the
/// header (`BrtWsDim`) is written from the *declared* extent, not from
/// what's actually written. This proves the resulting file still reads
/// back correctly: values, mixed types, formatting, and (deliberately) a
/// declared extent bigger than the real content, to pin the "superset
/// `BrtWsDim` is safe" finding from that round's real-Excel-COM
/// verification.
#[test]
fn roundtrip_sized_streaming_workbook_through_calamine() {
    use xlsb_write::StreamingWorkbook;

    let mut buf = Cursor::new(Vec::new());
    let mut wb = StreamingWorkbook::create(&mut buf);

    // Declare a range larger than what's actually written (10 rows x 5
    // cols declared, only 4 rows x 2 cols ever touched) — the empirically
    // verified superset case.
    let mut sheet = wb.new_worksheet_sized("Sized", 9, 4);
    for row in 0..4u32 {
        sheet.write_string(row, 0, &format!("sized-{row}")).unwrap();
        sheet.write_number(row, 1, row as f64 * 2.5).unwrap();
    }
    sheet.finish().unwrap();

    wb.finish().unwrap();
    let bytes = buf.into_inner();
    assert_eq!(&bytes[..2], b"PK");

    let path = std::env::temp_dir().join("xlsb_write_roundtrip_sized_streaming.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let mut wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let range = wbk.worksheet_range("Sized").expect("Sized sheet not found");
    for row in 0..4u32 {
        assert_eq!(range.get_value((row, 0)), Some(&Data::String(format!("sized-{row}"))));
        assert_eq!(range.get_value((row, 1)).and_then(Data::as_f64), Some(row as f64 * 2.5));
    }

    println!("Sized-streaming round-trip OK. File written to: {}", path.display());
}

/// `Workbook::define_name` — unlike autofilter (see `tests/autofilter.rs`;
/// calamine has no concept of it at all), calamine DOES expose defined
/// names via `Reader::defined_names()`, returning `(name, formula_text)`
/// pairs it decodes from the same `BrtName`/`BrtExternSheet` records this
/// crate now writes — so this is a genuine independent-reader check, not
/// just a structural one. Two names: one on the first sheet (reuses XTI
/// entry 0), one on a later sheet (needs its own new XTI entry) — see
/// `wb_part.rs`'s module doc comment for why that distinction matters.
#[test]
fn roundtrip_defined_names_through_calamine() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_worksheet("Sheet1");
    sheet1.write_string(0, 0, "A1");
    sheet1.write_string(1, 1, "B2");
    let sheet2 = wb.add_worksheet("Sheet2");
    sheet2.write_string(2, 2, "C3");

    wb.define_name("FirstSheetRange", 0, 0, 0, 1, 1); // Sheet1!$A$1:$B$2
    wb.define_name("SecondSheetCell", 1, 2, 2, 2, 2); // Sheet2!$C$3

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let bytes = buf.into_inner();
    assert_eq!(&bytes[..2], b"PK");

    let path = std::env::temp_dir().join("xlsb_write_roundtrip_defined_names.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let names = wbk.defined_names();

    let first = names
        .iter()
        .find(|(n, _)| n == "FirstSheetRange")
        .expect("FirstSheetRange not found");
    assert_eq!(first.1, "Sheet1!$A$1:$B$2");

    let second = names
        .iter()
        .find(|(n, _)| n == "SecondSheetCell")
        .expect("SecondSheetCell not found");
    assert_eq!(second.1, "Sheet2!$C$3:$C$3");

    println!("Defined-names round-trip OK. File written to: {}", path.display());
}

/// Same as above, but via `StreamingWorkbook` — proves `define_name` works
/// identically regardless of which workbook type is used, same as every
/// other workbook-level/per-sheet feature in this crate.
#[test]
fn roundtrip_defined_names_streaming_workbook_through_calamine() {
    use xlsb_write::StreamingWorkbook;

    let mut buf = Cursor::new(Vec::new());
    let mut wb = StreamingWorkbook::create(&mut buf);

    let mut sheet1 = wb.new_worksheet("Sheet1");
    sheet1.write_string(0, 0, "A1");
    wb.finish_worksheet(sheet1).unwrap();

    let mut sheet2 = wb.new_worksheet("Sheet2");
    sheet2.write_string(0, 0, "A1-on-sheet2");
    wb.finish_worksheet(sheet2).unwrap();

    wb.define_name("OnSheet2", 1, 0, 0, 0, 0); // Sheet2!$A$1

    wb.finish().unwrap();
    let bytes = buf.into_inner();
    assert_eq!(&bytes[..2], b"PK");

    let path = std::env::temp_dir().join("xlsb_write_roundtrip_defined_names_streaming.xlsb");
    std::fs::write(&path, &bytes).unwrap();

    let wbk: Xlsb<_> = open_workbook(&path).expect("calamine failed to open the file");
    let names = wbk.defined_names();
    let entry = names.iter().find(|(n, _)| n == "OnSheet2").expect("OnSheet2 not found");
    assert_eq!(entry.1, "Sheet2!$A$1:$A$1");
}
