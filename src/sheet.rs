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

/// `BrtBeginAFilter`/`BrtEndAFilter` ([MS-XLSB] AUTOFILTER = BrtBeginAFilter
/// *FILTERCOLUMN [SORTSTATE] BrtEndAFilter — this crate only ever emits the
/// begin/end pair with no filter columns, i.e. "show the dropdowns" without
/// pre-set criteria). Like `BrtDrawing`, the published MS-XLSB HTML pages
/// don't render the ABNF grammar that would say exactly where this pair
/// belongs in the worksheet record stream or that `BrtBeginAFilter`'s own
/// payload carries the filtered range, so this was derived empirically
/// (2026-09-13): built two independent reference files with real Excel
/// (COM automation, `Range.AutoFilter`) — one with the filter starting at
/// row 0, one starting mid-sheet on a renamed/reordered multi-sheet
/// workbook — and inspected both with `examples/dump_sheet.rs`. Both agree:
/// `BrtBeginAFilter` (161) carries the filtered range as a plain
/// `UncheckedRfX` (rowFirst/rowLast/colFirst/colLast, 4×u32 LE — the exact
/// same shape as `BrtWsDim`'s payload above), immediately followed by a
/// zero-length `BrtEndAFilter` (162) since there are no filter-column
/// entries. Also confirmed: this pair sits between the sheet-protection
/// record (rid 535, `FOOTER_MARGINS_LEAD`) and `BrtMargins`
/// (`FOOTER_MARGINS_TAIL`) — i.e. *before* the margins, not after — and is
/// wrapped in its own FRT shell (rid 37/38) carrying a copy of the same
/// per-sheet pseudo-unique identifier the trailing FRT shell at the very
/// end of the footer also carries (verified: both copies are byte-identical
/// within one sheet in every real reference file built for this). A fourth
/// reference file combining autofilter AND an embedded image (see
/// `drawing.rs`) confirmed the two features' insertion points don't
/// interact: autofilter's block sits entirely before `BrtMargins`,
/// `BrtDrawing` sits entirely after it, same relative position as when only
/// one of the two features is present.
const RID_BEGIN_AFILTER: u32 = 161;
const RID_END_AFILTER: u32 = 162;
/// The FRT (Future Record Type) wrapper used both around the per-sheet
/// pseudo-unique identifier (rid 3072) at the very end of the footer, and —
/// when autofilter is present — a second time around a copy of that same
/// identifier just before `BrtBeginAFilter`. Payload is fixed (verified
/// identical across every reference file this crate has inspected).
const RID_FRT_BEGIN: u32 = 37;
const RID_FRT_END: u32 = 38;
const RID_FRT_IDENTIFIER: u32 = 3072;
const FRT_HEADER_PAYLOAD: &[u8] = &[0x01, 0x00, 0x00, 0x10, 0x00, 0x80];
const RID_END_SHEET: u32 = 130;

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

/// The sheet-protection record (rid 535) that leads the footer's fixed
/// tail — verbatim from a real Excel-produced reference file. Ends exactly
/// where `BrtBeginAFilter`/`BrtEndAFilter` needs to be spliced in when the
/// sheet has an autofilter (see `RID_BEGIN_AFILTER`'s doc comment) — this
/// split point (69 bytes: a 3-byte record header + 66-byte payload) was
/// found by parsing this crate's own (pre-split) footer output back with
/// `biff12::parse_records`/`read_vi` and locating the record boundary, not
/// by hand-counting the array, so it can't silently drift if this blob is
/// ever re-transcribed.
#[rustfmt::skip]
const FOOTER_MARGINS_LEAD: &[u8] = &[
    0x97, 0x04, 0x42, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00,
];
/// `BrtPrintOptions`(477, 2 bytes) + `BrtMargins`(476, 48 bytes) — the rest
/// of the original single-block footer tail, verbatim from the same
/// reference file as `FOOTER_MARGINS_LEAD`. `BrtDrawing` (when the sheet has
/// an image) is spliced in right after this, before the trailing per-sheet
/// FRT identifier — see `write_sheet_footer`.
#[rustfmt::skip]
const FOOTER_MARGINS_TAIL: &[u8] = &[
    0xdd, 0x03, 0x02,
    0x10, 0x00, 0xdc, 0x03, 0x30, 0x66, 0x66, 0x66, 0x66, 0x66, 0x66, 0xe6, 0x3f, 0x66, 0x66, 0x66, 0x66, 0x66,
    0x66, 0xe6, 0x3f, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xe8, 0x3f, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xe8,
    0x3f, 0x33, 0x33, 0x33, 0x33, 0x33, 0x33, 0xd3, 0x3f, 0x33, 0x33, 0x33, 0x33, 0x33, 0x33, 0xd3, 0x3f,
];

/// Write the FRT (Future Record Type) wrapper around a copy of `guid` —
/// used both for the trailing per-sheet identifier and, when autofilter is
/// present, a second time immediately before `BrtBeginAFilter` (see
/// `RID_BEGIN_AFILTER`'s doc comment: real Excel writes the *same* 16 bytes
/// in both places within one sheet).
fn write_frt_identifier(guid: &[u8; 16], buf: &mut Vec<u8>) {
    write_rec(RID_FRT_BEGIN, FRT_HEADER_PAYLOAD, buf);
    write_rec(RID_FRT_IDENTIFIER, guid, buf);
    write_r0(RID_FRT_END, buf);
}

/// `merges`: `(first_row, first_col, last_row, last_col)` per merged range.
/// `has_drawing`: whether this sheet has any embedded images — emits a
/// `BrtDrawing` record pointing at `xl/worksheets/_rels/sheetN.bin.rels`'s
/// `rId1` (the sheet's drawing relationship) when true.
/// `autofilter`: `Some((first_row, first_col, last_row, last_col))` when
/// `Worksheet::set_autofilter` was called — emits `BrtBeginAFilter`/
/// `BrtEndAFilter` (see `RID_BEGIN_AFILTER`'s doc comment for exact shape
/// and position, empirically derived).
pub fn write_sheet_footer(
    merges: &[(u32, u32, u32, u32)],
    has_drawing: bool,
    autofilter: Option<(u32, u32, u32, u32)>,
    buf: &mut Vec<u8>,
) {
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

    buf.extend_from_slice(FOOTER_MARGINS_LEAD);

    // The per-sheet pseudo-unique identifier is generated ONCE per footer
    // and reused for both FRT-wrapped copies real Excel writes when
    // autofilter is present (see `RID_BEGIN_AFILTER`'s doc comment) — a
    // sheet with no autofilter only ever uses it once, at the very end.
    let guid = crate::biff12::pseudo_unique_16_bytes();

    if let Some((r0, c0, r1, c1)) = autofilter {
        write_frt_identifier(&guid, buf);
        let mut pay = [0u8; 16];
        pay[0..4].copy_from_slice(&r0.to_le_bytes());
        pay[4..8].copy_from_slice(&r1.to_le_bytes());
        pay[8..12].copy_from_slice(&c0.to_le_bytes());
        pay[12..16].copy_from_slice(&c1.to_le_bytes());
        write_rec(RID_BEGIN_AFILTER, &pay, buf);
        write_r0(RID_END_AFILTER, buf);
    }

    buf.extend_from_slice(FOOTER_MARGINS_TAIL);

    if has_drawing {
        let mut pay = Vec::new();
        write_wstr("rId1", &mut pay);
        write_rec(RID_DRAWING, &pay, buf);
    }

    write_frt_identifier(&guid, buf);
    write_r0(RID_END_SHEET, buf);
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
        write_sheet_footer(&[], false, None, &mut buf1);
        let mut buf2 = Vec::new();
        write_sheet_footer(&[], false, None, &mut buf2);

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
        write_sheet_footer(&[], false, None, &mut buf);
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
        write_sheet_footer(&[], true, None, &mut buf);
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
        write_sheet_footer(&[], false, None, &mut without);
        let mut with = Vec::new();
        write_sheet_footer(&[], true, None, &mut with);

        let mut expected_record = Vec::new();
        let mut pay = Vec::new();
        write_wstr("rId1", &mut pay);
        write_rec(RID_DRAWING, &pay, &mut expected_record);

        assert_eq!(with.len(), without.len() + expected_record.len());
    }

    /// No autofilter (the default) must not change the footer at all —
    /// verified the same way as `no_drawing_record_when_sheet_has_no_images`.
    #[test]
    fn no_autofilter_records_when_autofilter_is_none() {
        let mut buf = Vec::new();
        write_sheet_footer(&[], false, None, &mut buf);
        let recs = crate::biff12::parse_records(&buf);
        assert!(
            !recs
                .iter()
                .any(|(rid, _)| *rid == RID_BEGIN_AFILTER || *rid == RID_END_AFILTER),
            "BrtBeginAFilter/BrtEndAFilter must not appear when autofilter is None"
        );
    }

    /// `autofilter = Some(range)` must emit `BrtBeginAFilter` (161, payload
    /// = the range as rowFirst/rowLast/colFirst/colLast, same shape as
    /// `BrtWsDim`) immediately followed by a zero-length `BrtEndAFilter`
    /// (162), sitting right after the sheet-protection record and its own
    /// FRT-wrapped identifier, and before `BrtPrintOptions`/`BrtMargins` —
    /// the exact position/shape confirmed against two independent real
    /// Excel-produced reference files (see `RID_BEGIN_AFILTER`'s doc
    /// comment).
    #[test]
    fn autofilter_record_has_expected_payload_and_position() {
        let mut buf = Vec::new();
        write_sheet_footer(&[], false, Some((1, 2, 5, 4)), &mut buf);
        let recs = crate::biff12::parse_records(&buf);

        let begin_positions: Vec<usize> = recs
            .iter()
            .enumerate()
            .filter(|(_, (rid, _))| *rid == RID_BEGIN_AFILTER)
            .map(|(i, _)| i)
            .collect();
        assert_eq!(begin_positions.len(), 1, "exactly one BrtBeginAFilter expected");
        let pos = begin_positions[0];

        let (_, payload) = &recs[pos];
        let mut expected = [0u8; 16];
        expected[0..4].copy_from_slice(&1u32.to_le_bytes()); // rowFirst
        expected[4..8].copy_from_slice(&5u32.to_le_bytes()); // rowLast
        expected[8..12].copy_from_slice(&2u32.to_le_bytes()); // colFirst
        expected[12..16].copy_from_slice(&4u32.to_le_bytes()); // colLast
        assert_eq!(payload, &expected);

        // Immediately followed by a zero-length BrtEndAFilter.
        let (next_rid, next_payload) = &recs[pos + 1];
        assert_eq!(*next_rid, RID_END_AFILTER);
        assert!(next_payload.is_empty());

        // Immediately preceded by the rid-37/rid-3072/rid-38 FRT wrapper
        // carrying the per-sheet identifier.
        assert_eq!(recs[pos - 1].0, RID_FRT_END);
        assert_eq!(recs[pos - 2].0, RID_FRT_IDENTIFIER);
        assert_eq!(recs[pos - 3].0, RID_FRT_BEGIN);

        // And BrtPrintOptions (477) follows right after BrtEndAFilter.
        assert_eq!(recs[pos + 2].0, 477);
    }

    /// The FRT-wrapped identifier written just before `BrtBeginAFilter` must
    /// be byte-identical to the one written at the very end of the footer —
    /// confirmed against real Excel output that both copies within one
    /// sheet always match (see `RID_BEGIN_AFILTER`'s doc comment).
    #[test]
    fn autofilter_frt_identifier_matches_trailing_identifier() {
        let mut buf = Vec::new();
        write_sheet_footer(&[], false, Some((0, 0, 0, 0)), &mut buf);
        let recs = crate::biff12::parse_records(&buf);
        let identifiers: Vec<&Vec<u8>> = recs
            .iter()
            .filter(|(rid, _)| *rid == RID_FRT_IDENTIFIER)
            .map(|(_, p)| p)
            .collect();
        assert_eq!(identifiers.len(), 2, "expected two FRT-identifier copies");
        assert_eq!(identifiers[0], identifiers[1]);
    }

    /// Autofilter and an embedded image can coexist without their insertion
    /// points colliding — confirmed against a fourth real Excel reference
    /// file combining both features (see `RID_BEGIN_AFILTER`'s doc comment).
    #[test]
    fn autofilter_and_drawing_coexist_in_expected_order() {
        let mut buf = Vec::new();
        write_sheet_footer(&[], true, Some((0, 0, 1, 1)), &mut buf);
        let recs = crate::biff12::parse_records(&buf);
        let order: Vec<u32> = recs.iter().map(|(rid, _)| *rid).collect();

        let af_pos = order.iter().position(|&r| r == RID_BEGIN_AFILTER).unwrap();
        let margins_pos = order.iter().position(|&r| r == 476).unwrap();
        let drawing_pos = order.iter().position(|&r| r == RID_DRAWING).unwrap();
        assert!(af_pos < margins_pos, "autofilter must come before BrtMargins");
        assert!(margins_pos < drawing_pos, "BrtDrawing must come after BrtMargins");
        assert_eq!(order.last(), Some(&RID_END_SHEET));
    }
}
