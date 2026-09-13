//! Sheet (worksheet) binary encoder.
//!
//! Structural framing is adapted from the `xlsb-writer` crate (MIT License,
//! Copyright (c) 2026 kotucha), byte-verified against real Excel output.
//! The original wrote a duplicate of `ROW_PRE` inline in the header to
//! cover "row 0"; byte comparison shows that inline block is identical to
//! `ROW_PRE` itself, so here every row — including row 0 — uses the same
//! `ROW_PRE` prefix uniformly, which removes the "row 0 must always be
//! present" assumption.
//!
//! Record IDs below are cross-checked against `pyxlsb`'s `recordtypes.py`
//! (an independent, authoritative reader implementation) — this caught a
//! real mislabeling in the original: what it called "BrtColInfo" (rid
//! 0x0098) is actually `BrtSel` (selection state); the real `BrtColInfo`
//! is rid 60 and lives in its own `BrtBeginColInfos`(390)/`BrtEndColInfos`
//! (391) collection, which this module now emits for real. `BrtWsDim`
//! (148/0x94) and `BrtWsProp` (147/0x93) were also swapped in the
//! original's comments (bytes/order were already correct, only the labels
//! were wrong).

use crate::CellValue;
use crate::biff12::*;

const RID_BEGIN_COL_INFOS: u32 = 390;
const RID_END_COL_INFOS: u32 = 391;
const RID_COL_INFO: u32 = 60;
const RID_BEGIN_MERGE_CELLS: u32 = 177;
const RID_END_MERGE_CELLS: u32 = 178;
const RID_MERGE_CELL: u32 = 176;

/// One `set_column`-style entry: `(col, width_chars, hidden)`.
pub type ColSpec = (u32, f64, bool);

/// Used-range dimensions: `(first_row, last_row, first_col, last_col)`,
/// zero-based and inclusive. `None` for an empty sheet.
pub type Dimension = Option<(u32, u32, u32, u32)>;

#[allow(clippy::too_many_arguments)]
pub fn write_sheet_header(
    freeze_row: u32,
    freeze_col: u32,
    col_specs: &[ColSpec],
    dim: Dimension,
    active: bool,
    buf: &mut Vec<u8>,
) {
    write_r0(0x0081, buf); // BrtBeginSheet (129)
    // BrtWsProp (147/0x93) — static (scroll position / outline hint, same in all reference files)
    write_rec(
        0x0093,
        &[
            0xc9, 0x04, 0x02, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff,
            0xff, 0x00, 0x00, 0x00, 0x00,
        ],
        buf,
    );
    // BrtWsDim (148/0x94) — the sheet's real used range. This MUST reflect
    // actual content: a hardcoded/placeholder value here (the original
    // vendored crate had one baked in from whatever reference file it was
    // extracted from) makes Excel's strict validator discard the worksheet
    // part on open ("Replaced Part") the moment real content's extent
    // disagrees with the declared dimension — confirmed by comparing a
    // real empty-workbook reference (dimension all-zero) against a real
    // populated one (dimension matching its actual last row/col) byte-for-byte.
    let (rw_first, rw_last, col_first, col_last) = dim.unwrap_or((0, 0, 0, 0));
    let mut wsdim = [0u8; 16];
    wsdim[0..4].copy_from_slice(&rw_first.to_le_bytes());
    wsdim[4..8].copy_from_slice(&rw_last.to_le_bytes());
    wsdim[8..12].copy_from_slice(&col_first.to_le_bytes());
    wsdim[12..16].copy_from_slice(&col_last.to_le_bytes());
    write_rec(0x0094, &wsdim, buf);

    // BrtBeginWsViews(133)/BrtEndWsViews(134), wrapping BrtBeginWsView(137)
    // /BrtPane(151)/BrtSel(152)/BrtEndWsView(138) — sheet view + freeze state.
    write_r0(0x0085, buf);
    write_rec(0x0089, &bci_payload(active), buf);
    if freeze_row > 0 || freeze_col > 0 {
        let mut pane = [0u8; 29];
        pane[0..8].copy_from_slice(&(freeze_col as f64).to_le_bytes()); // xnumXSplit
        pane[8..16].copy_from_slice(&(freeze_row as f64).to_le_bytes()); // xnumYSplit
        pane[16..20].copy_from_slice(&freeze_row.to_le_bytes()); // rwTop
        pane[20..24].copy_from_slice(&freeze_col.to_le_bytes()); // colLeft
        pane[24..28].copy_from_slice(&2u32.to_le_bytes()); // pnnAct = bottomLeft
        pane[28] = 0x03; // fFrozen | fFrozenNoSplit
        write_rec(0x0097, &pane, buf);
        let (sel0, sel1) = sel_payload_frozen();
        write_rec(0x0098, &sel0, buf);
        write_rec(0x0098, &sel1, buf);
    } else {
        write_rec(0x0098, &sel_payload(), buf);
    }
    write_r0(0x008A, buf); // BrtEndWsView
    write_r0(0x0086, buf); // BrtEndWsViews

    // BrtBeginColInfos(390)/BrtColInfo(60)*/BrtEndColInfos(391) — real
    // per-column width/hidden state. Only emitted when the caller actually
    // set a column, so unmodified sheets stay byte-identical to before.
    if !col_specs.is_empty() {
        write_r0(RID_BEGIN_COL_INFOS, buf);
        for &(col, width_chars, hidden) in col_specs {
            write_rec(RID_COL_INFO, &encode_col_info(col, col, width_chars, hidden), buf);
        }
        write_r0(RID_END_COL_INFOS, buf);
    }

    // Static preamble (verified from reference) — BrtBeginList/0x0415/BrtEndList wrapper.
    write_rec(0x0025, &[0x01, 0x00, 0x02, 0x0e, 0x00, 0x80], buf);
    write_rec(0x0415, &[0x05, 0x00], buf);
    write_r0(0x0026, buf);
    write_rec(
        0x01E5,
        &[0xff, 0xff, 0xff, 0xff, 0x08, 0x00, 0x2c, 0x01, 0x00, 0x00, 0x00, 0x00],
        buf,
    );
    write_r0(0x0091, buf); // BrtBeginSheetData (145)
}

/// `BrtDrawing` (record 550) — [MS-XLSB] 2.4.354: a link to this sheet's
/// `Drawings` part. Payload is `stRelId`, a `RelID`
/// ([MS-XLSB] 2.5.115: an `XLNullableWideString`) naming the relationship
/// in this sheet's own `.bin.rels` file — always `"rId1"` here, since this
/// crate never emits `xl/worksheets/binaryIndexN.bin` (backlog item 11),
/// so the drawing relationship is always the sheet's only one (see
/// `drawing::sheet_rels`, which must agree with this constant).
const RID_DRAWING: u32 = 550;

/// `merges`: `(first_row, first_col, last_row, last_col)` per merged range.
/// `has_drawing`: whether this sheet has any embedded images — emits a
/// `BrtDrawing` record pointing at `xl/worksheets/_rels/sheetN.bin.rels`'s
/// `rId1` (the sheet's drawing relationship) when true.
pub fn write_sheet_footer(merges: &[(u32, u32, u32, u32)], has_drawing: bool, buf: &mut Vec<u8>) {
    buf.extend_from_slice(&[0x92, 0x01, 0x00]); // BrtEndSheetData (146)

    if !merges.is_empty() {
        // Unlike BrtBeginColInfos/BrtBeginSheetData, BrtBeginMergeCells DOES
        // carry a 4-byte count (confirmed against a real Excel-produced
        // file with an actual merge: payload was `01 00 00 00` for 1 merge,
        // not zero-length as originally assumed).
        write_rec(RID_BEGIN_MERGE_CELLS, &(merges.len() as u32).to_le_bytes(), buf);
        for &(r0, c0, r1, c1) in merges {
            let mut pay = [0u8; 16];
            pay[0..4].copy_from_slice(&r0.to_le_bytes());
            pay[4..8].copy_from_slice(&r1.to_le_bytes());
            pay[8..12].copy_from_slice(&c0.to_le_bytes());
            pay[12..16].copy_from_slice(&c1.to_le_bytes());
            write_rec(RID_MERGE_CELL, &pay, buf);
        }
        write_r0(RID_END_MERGE_CELLS, buf);
    }

    // 157 bytes verbatim from a real Excel-produced reference file (view/
    // pane state, default column-break etc.), split into two halves so a
    // `BrtDrawing` record can be spliced in between them when the sheet
    // has an image — re-extracted byte-for-byte from a fresh reference
    // after the previous transcription of this array (during the earlier
    // ROW_PRE refactor) turned out to be 8 bytes short, which was the
    // actual cause of Excel discarding the worksheet part on open
    // ("Replaced Part") — see write_sheet_header's BrtWsDim comment for
    // the (unrelated, also real) other half of that bug.
    //
    // Confirmed against a second, independently-produced real multi-sheet
    // Excel file (8 sheets vs. the original single-sheet reference) that
    // every byte here is identical across both files — EXCEPT the 16-byte
    // value embedded in `FOOTER_TAIL_B` (an `[MS-XLSB]` FRT-wrapped
    // record, rid 3072): a real Excel file stamps a fresh value on every
    // sheet, but this crate was copying the SAME 16 bytes (from whichever
    // one sheet the original reference happened to have) onto every sheet
    // it ever wrote. Fixed below: that slice is now generated per sheet
    // instead of baked into the static template.
    //
    // The split point (right after `BrtMargins`'s 48-byte payload, right
    // before the rid-37 wrapper around the per-sheet FRT identifier) is
    // exactly where a real Excel-authored `.xlsb` with an inserted picture
    // places its own `BrtDrawing` record — confirmed 2026-09-13 by
    // building one with Excel itself (COM automation) and dumping its
    // `sheet1.bin` with `examples/dump_sheet.rs`; this ordering isn't
    // documented anywhere the published MS-XLSB HTML pages render (the
    // actual ABNF grammar file they cite isn't included), so it's trusted
    // from that real reference file, not guessed from spec text alone.
    #[rustfmt::skip]
    const FOOTER_TAIL_A: &[u8] = &[
        0x97, 0x04, 0x42, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0xdd, 0x03, 0x02,
        0x10, 0x00, 0xdc, 0x03, 0x30, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0xe6, 0x3f, 0x66, 0x66, 0x66, 0x66, 0x66,
        0x66, 0xe6, 0x3f, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xe8, 0x3f, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xe8,
        0x3f, 0x33, 0x33, 0x33, 0x33, 0x33, 0x33, 0xd3, 0x3f, 0x33, 0x33, 0x33, 0x33, 0x33, 0x33, 0xd3, 0x3f,
    ];
    // Starts with the rid-37 wrapper (`0x25, 0x06, ...`), then the rid-3072
    // FRT-wrapped 16-byte identifier (placeholder here — replaced below,
    // per-sheet), then the fixed 5-byte `[BrtEndList, BrtEndSheet]` tail.
    #[rustfmt::skip]
    const FOOTER_TAIL_B: &[u8] = &[
        0x25,
        0x06, 0x01, 0x00, 0x00, 0x10, 0x00, 0x80, 0x80, 0x18, 0x10, 0xf0, 0xf8, 0x2a, 0xf4, 0xf1, 0x21, 0x6e, 0x47,
        0x9c, 0x59, 0xcf, 0x74, 0xd2, 0xaa, 0x32, 0x93, 0x26, 0x00, 0x82, 0x01, 0x00,
    ];
    const FOOTER_B_GUID_START: usize = 11;
    const FOOTER_B_GUID_LEN: usize = 16;

    buf.extend_from_slice(FOOTER_TAIL_A);

    if has_drawing {
        let mut pay = Vec::new();
        write_wstr("rId1", &mut pay);
        write_rec(RID_DRAWING, &pay, buf);
    }

    // The rid-3072 payload is the 16 bytes at a fixed offset into
    // `FOOTER_TAIL_B` (see the constants above) — sliced by fixed
    // start/length, not by a hand-counted offset from the end, so this
    // can't silently drift out of sync if `FOOTER_TAIL_B` is ever
    // re-transcribed.
    let guid_end = FOOTER_B_GUID_START + FOOTER_B_GUID_LEN;
    buf.extend_from_slice(&FOOTER_TAIL_B[..FOOTER_B_GUID_START]);
    buf.extend_from_slice(&crate::biff12::pseudo_unique_16_bytes());
    buf.extend_from_slice(&FOOTER_TAIL_B[guid_end..]);
}

/// `BrtColInfo` (rid 60): the first 18 bytes are the fields real readers
/// (verified against `pyxlsb`) actually rely on — colFirst(4)/colLast(4)/
/// coldx(4, width×256)/ixfe(4)/grbit(2, bit0=fHidden, bit1=fUserSet). No
/// trailing padding is required: BIFF12 records are length-prefixed, so a
/// reader stops at the declared length regardless of a type's "usual" size.
fn encode_col_info(col_first: u32, col_last: u32, width_chars: f64, hidden: bool) -> [u8; 18] {
    let mut p = [0u8; 18];
    p[0..4].copy_from_slice(&col_first.to_le_bytes());
    p[4..8].copy_from_slice(&col_last.to_le_bytes());
    let coldx = (width_chars * 256.0).round() as u32;
    p[8..12].copy_from_slice(&coldx.to_le_bytes());
    // p[12..16] = ixfe = 0 (default column format)
    let mut grbit: u16 = 0x0002; // fUserSet — this width was explicitly requested
    if hidden {
        grbit |= 0x0001;
    }
    p[16..18].copy_from_slice(&grbit.to_le_bytes());
    p
}

/// `BrtBeginWsView` payload. Byte 0's bit 6 is `fSelected` ([MS-XLSB]
/// 2.4.24) — whether this sheet's own tab is selected. The reference bytes
/// this was extracted from have it set (`0xDC`, from a single-sheet file
/// where that one sheet is naturally both active and selected); writing
/// that unconditionally on every sheet means every tab reports itself
/// selected, so Excel opens the workbook with all sheets grouped (the same
/// state you get from Shift-clicking every tab — edits/filters on one
/// sheet apply to all of them). Only the active sheet (`active: true`,
/// meaning the first sheet — see `Workbook::write`) should keep it set.
fn bci_payload(active: bool) -> [u8; 30] {
    let mut p = [0u8; 30];
    let byte0: u32 = if active { 0xDC } else { 0xDC & !0x40 };
    p[0..4].copy_from_slice(&(0x300 | byte0).to_le_bytes());
    p[14..18].copy_from_slice(&0x40u32.to_le_bytes());
    p[18..22].copy_from_slice(&0x64u32.to_le_bytes());
    p
}

fn sel_payload() -> [u8; 36] {
    let mut p = [0u8; 36];
    p[0..4].copy_from_slice(&3u32.to_le_bytes());
    p[16..20].copy_from_slice(&1u32.to_le_bytes());
    p
}

fn sel_payload_frozen() -> ([u8; 36], [u8; 36]) {
    let mut p0 = [0u8; 36];
    p0[0..4].copy_from_slice(&3u32.to_le_bytes());
    p0[16..20].copy_from_slice(&1u32.to_le_bytes());
    let mut p1 = [0u8; 36];
    p1[0..4].copy_from_slice(&2u32.to_le_bytes());
    p1[16..20].copy_from_slice(&1u32.to_le_bytes());
    (p0, p1)
}

/// Encode one row's cells. `cells` must be sorted by column ascending.
/// Each entry is `(col, value, xf_index)`.
pub fn encode_row(
    row_idx: u32,
    cells: &[(u32, &CellValue, u16)],
    height_twips: u16,
    hidden: bool,
    sst: &mut crate::sst::Sst,
    sheet_buf: &mut Vec<u8>,
) {
    sheet_buf.extend_from_slice(ROW_PRE);

    let mut cell_buf = Vec::with_capacity(cells.len() * 16);
    let mut last_col = 0u32;

    for &(col, value, ixfe) in cells {
        last_col = col;
        match value {
            CellValue::Blank => write_cell_blank(col, ixfe, &mut cell_buf),
            CellValue::Bool(b) => write_cell_bool(col, ixfe, *b, &mut cell_buf),
            CellValue::Number(v) => {
                if let Some(rk) = encode_rk(*v) {
                    write_cell_rk(col, ixfe, rk, &mut cell_buf);
                } else {
                    write_cell_real(col, ixfe, *v, &mut cell_buf);
                }
            }
            CellValue::String(s) => {
                if s.is_empty() {
                    write_cell_blank(col, ixfe, &mut cell_buf);
                } else {
                    let isst = sst.intern(s);
                    write_cell_isst(col, ixfe, isst, &mut cell_buf);
                }
            }
            CellValue::FormulaNum(formula, cached) => {
                crate::formula::write_fmla_num(col, ixfe, *cached, formula, &mut cell_buf);
            }
            CellValue::FormulaStr(formula, cached) => {
                crate::formula::write_fmla_string(col, ixfe, cached, formula, &mut cell_buf);
            }
        }
    }

    write_row_hdr(row_idx, last_col, height_twips, hidden, sheet_buf);
    sheet_buf.extend_from_slice(&cell_buf);
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Regression: every sheet used to hardcode `fSelected=1` in its own
    /// `BrtBeginWsView`, so a multi-sheet workbook opened in Excel with
    /// every tab grouped (as if the user had Shift-clicked all of them) —
    /// only the active sheet may have this bit set.
    #[test]
    fn only_active_sheet_has_f_selected_bit_set() {
        const F_SELECTED: u8 = 0x40;
        assert_ne!(
            bci_payload(true)[0] & F_SELECTED,
            0,
            "active sheet must have fSelected set"
        );
        assert_eq!(
            bci_payload(false)[0] & F_SELECTED,
            0,
            "inactive sheet must NOT have fSelected set"
        );
        // Every other bit/byte must be identical regardless of `active`.
        let (mut a, mut b) = (bci_payload(true), bci_payload(false));
        a[0] &= !F_SELECTED;
        b[0] &= !F_SELECTED;
        assert_eq!(a, b);
    }

    /// Regression: every sheet used to embed the exact same 16-byte value
    /// (copied from whichever one sheet the original reference file
    /// happened to have) in its footer — confirmed wrong against a real
    /// multi-sheet file, where this field is unique per sheet. Two
    /// `write_sheet_footer` calls must produce different bytes there,
    /// while everything else in the footer (same length, same bytes
    /// outside that one field) stays identical.
    #[test]
    fn footer_identifier_varies_between_sheets_rest_stays_fixed() {
        let mut buf1 = Vec::new();
        write_sheet_footer(&[], false, &mut buf1);
        let mut buf2 = Vec::new();
        write_sheet_footer(&[], false, &mut buf2);

        assert_eq!(
            buf1.len(),
            buf2.len(),
            "footer length must not depend on the generated identifier"
        );
        assert_ne!(buf1, buf2, "the per-sheet identifier must differ between calls");

        // Only the 16 bytes right before the fixed 5-byte tail may differ.
        let tail_start = buf1.len() - 5;
        let id_start = tail_start - 16;
        assert_eq!(
            buf1[..id_start],
            buf2[..id_start],
            "bytes before the identifier must be identical"
        );
        assert_eq!(
            buf1[tail_start..],
            buf2[tail_start..],
            "bytes after the identifier must be identical"
        );
        assert_ne!(
            buf1[id_start..tail_start],
            buf2[id_start..tail_start],
            "the identifier itself must differ"
        );
    }

    /// `has_drawing=false` (the default, for every sheet without an
    /// embedded image) must not change the footer at all — a `BrtDrawing`
    /// record referencing a `.rels` relationship that doesn't exist would
    /// be a real corruption bug for the overwhelming majority of sheets
    /// that never call `embed_image`.
    #[test]
    fn no_drawing_record_when_sheet_has_no_images() {
        let mut buf = Vec::new();
        write_sheet_footer(&[], false, &mut buf);
        let recs = crate::biff12::parse_records(&buf);
        assert!(
            !recs.iter().any(|(rid, _)| *rid == RID_DRAWING),
            "BrtDrawing must not appear when has_drawing is false"
        );
    }

    /// `has_drawing=true` must emit exactly one `BrtDrawing` (rid 550)
    /// record, with payload `stRelId = "rId1"` (an XLNullableWideString:
    /// cch=4 + UTF-16LE "rId1") — the exact shape and value confirmed
    /// against a real Excel-authored `.xlsb` with an inserted picture (see
    /// this module's `write_sheet_footer` doc comment). It must sit right
    /// after the margins/print-options records and before the per-sheet
    /// FRT identifier, i.e. its position within the parsed record stream
    /// must match where the real reference file put it.
    #[test]
    fn drawing_record_has_expected_payload_and_position() {
        let mut buf = Vec::new();
        write_sheet_footer(&[], true, &mut buf);
        let recs = crate::biff12::parse_records(&buf);

        let drawing_positions: Vec<usize> = recs
            .iter()
            .enumerate()
            .filter(|(_, (rid, _))| *rid == RID_DRAWING)
            .map(|(i, _)| i)
            .collect();
        assert_eq!(drawing_positions.len(), 1, "exactly one BrtDrawing record expected");

        let (_, payload) = &recs[drawing_positions[0]];
        let expected = {
            let mut p = Vec::new();
            write_wstr("rId1", &mut p);
            p
        };
        assert_eq!(
            payload, &expected,
            "BrtDrawing payload must be the XLNullableWideString \"rId1\""
        );

        // Immediately followed by the rid-37 wrapper that starts
        // FOOTER_TAIL_B (the per-sheet FRT identifier's envelope) —
        // pinning this catches an accidental reordering relative to the
        // real reference file's shape.
        let (next_rid, _) = &recs[drawing_positions[0] + 1];
        assert_eq!(
            *next_rid, 37,
            "BrtDrawing must be immediately followed by the rid-37 FRT wrapper"
        );

        // And the footer must still end in BrtEndSheet (130), same as
        // when there's no drawing.
        assert_eq!(recs.last().unwrap().0, 130, "footer must still end with BrtEndSheet");
    }

    /// A sheet with a drawing must produce a footer exactly
    /// `BrtDrawing`'s own encoded length longer than one without —
    /// nothing else should change size.
    #[test]
    fn drawing_record_adds_exactly_its_own_encoded_length() {
        let mut without = Vec::new();
        write_sheet_footer(&[], false, &mut without);
        let mut with = Vec::new();
        write_sheet_footer(&[], true, &mut with);

        let mut expected_record = Vec::new();
        let mut pay = Vec::new();
        write_wstr("rId1", &mut pay);
        write_rec(RID_DRAWING, &pay, &mut expected_record);

        assert_eq!(with.len(), without.len() + expected_record.len());
    }
}
