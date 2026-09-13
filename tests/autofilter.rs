//! Structural tests for `Worksheet::set_autofilter`.
//!
//! `calamine` (this crate's usual independent-reader check) doesn't surface
//! autofilter at all — it's a cell-data reader, same gap as embedded images
//! (see `tests/embed_image.rs`) — so these tests instead verify the actual
//! `BrtBeginAFilter`/`BrtEndAFilter` records this crate emits into the
//! sheet's own `.bin` (parsed back with `xlsb_write::biff12`, the same
//! structural check `examples/validate_xlsb.rs` runs), across all three
//! worksheet types. Real-Excel visual/COM verification for this feature is
//! documented in this session's report, since it needs Excel installed and
//! isn't something `cargo test` can assert on in CI.

use std::io::{Cursor, Read};
use xlsb_write::{StreamingWorkbook, Workbook};
use zip::ZipArchive;

const RID_BEGIN_AFILTER: u32 = 161;
const RID_END_AFILTER: u32 = 162;

fn extract_sheet_bin(bytes: &[u8], sheet_number: usize) -> Vec<u8> {
    let cursor = Cursor::new(bytes);
    let mut zip = ZipArchive::new(cursor).unwrap();
    let mut entry = zip.by_name(&format!("xl/worksheets/sheet{sheet_number}.bin")).unwrap();
    let mut buf = Vec::new();
    entry.read_to_end(&mut buf).unwrap();
    buf
}

fn expected_afilter_payload(first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> Vec<u8> {
    let mut p = [0u8; 16];
    p[0..4].copy_from_slice(&first_row.to_le_bytes());
    p[4..8].copy_from_slice(&last_row.to_le_bytes());
    p[8..12].copy_from_slice(&first_col.to_le_bytes());
    p[12..16].copy_from_slice(&last_col.to_le_bytes());
    p.to_vec()
}

#[test]
fn workbook_set_autofilter_produces_expected_record() {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.write_string(0, 0, "Name");
    sheet.write_string(0, 1, "Qty");
    sheet.write_string(1, 0, "Alpha");
    sheet.write_number(1, 1, 1.0);
    sheet.set_autofilter(0, 0, 1, 1);

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let sheet_bin = extract_sheet_bin(&buf.into_inner(), 1);

    let recs = xlsb_write::biff12::try_parse_records(&sheet_bin).expect("sheet1.bin must be well-formed BIFF12");
    let begin: Vec<&Vec<u8>> = recs
        .iter()
        .filter(|(rid, _)| *rid == RID_BEGIN_AFILTER)
        .map(|(_, p)| p)
        .collect();
    assert_eq!(begin.len(), 1, "exactly one BrtBeginAFilter expected");
    assert_eq!(*begin[0], expected_afilter_payload(0, 0, 1, 1));

    let end_count = recs.iter().filter(|(rid, _)| *rid == RID_END_AFILTER).count();
    assert_eq!(end_count, 1, "exactly one BrtEndAFilter expected");
}

/// A sheet with no `set_autofilter` call must get neither record — confirms
/// the feature is fully opt-in and doesn't change existing output.
#[test]
fn sheet_without_autofilter_gets_no_afilter_records() {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.write_string(0, 0, "No filter here");

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let sheet_bin = extract_sheet_bin(&buf.into_inner(), 1);

    let recs = xlsb_write::biff12::try_parse_records(&sheet_bin).unwrap();
    assert!(
        !recs
            .iter()
            .any(|(rid, _)| *rid == RID_BEGIN_AFILTER || *rid == RID_END_AFILTER)
    );
}

/// Same feature via `StreamingWorkbook` (the plain, whole-sheet-buffered
/// streaming type) — must produce equivalent output to the `Workbook` path.
#[test]
fn streaming_workbook_set_autofilter_produces_expected_record() {
    let mut buf = Cursor::new(Vec::new());
    let mut wb = StreamingWorkbook::create(&mut buf);
    let mut sheet = wb.new_worksheet("Sheet1");
    sheet.write_string(0, 0, "Name");
    sheet.write_string(1, 0, "Alpha");
    sheet.set_autofilter(0, 0, 3, 2);
    wb.finish_worksheet(sheet).unwrap();
    wb.finish().unwrap();

    let sheet_bin = extract_sheet_bin(&buf.into_inner(), 1);
    let recs = xlsb_write::biff12::try_parse_records(&sheet_bin).unwrap();
    let begin: Vec<&Vec<u8>> = recs
        .iter()
        .filter(|(rid, _)| *rid == RID_BEGIN_AFILTER)
        .map(|(_, p)| p)
        .collect();
    assert_eq!(begin.len(), 1);
    assert_eq!(*begin[0], expected_afilter_payload(0, 0, 3, 2));
}

/// Same feature via `SizedStreamingWorksheet` — `set_autofilter` must work
/// even though the header may already have been sent by the time it's
/// called (autofilter is written into the footer, not the header — same
/// "no ordering restriction" contract as `embed_image`).
#[test]
fn sized_streaming_worksheet_set_autofilter_produces_expected_record() {
    let mut buf = Cursor::new(Vec::new());
    let mut wb = StreamingWorkbook::create(&mut buf);
    let mut sheet = wb.new_worksheet_sized("Sheet1", 10, 5);
    sheet.write_string(0, 0, "row0").unwrap();
    sheet.write_string(1, 0, "row1").unwrap(); // flushes row 0, sends the header
    sheet.set_autofilter(0, 0, 1, 2); // called AFTER header already sent
    sheet.finish().unwrap();
    wb.finish().unwrap();

    let sheet_bin = extract_sheet_bin(&buf.into_inner(), 1);
    let recs = xlsb_write::biff12::try_parse_records(&sheet_bin).unwrap();
    let begin: Vec<&Vec<u8>> = recs
        .iter()
        .filter(|(rid, _)| *rid == RID_BEGIN_AFILTER)
        .map(|(_, p)| p)
        .collect();
    assert_eq!(begin.len(), 1);
    assert_eq!(*begin[0], expected_afilter_payload(0, 0, 1, 2));
}

#[test]
#[should_panic(expected = "inverted")]
fn set_autofilter_rejects_inverted_range() {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.set_autofilter(5, 5, 1, 1); // last < first
}
