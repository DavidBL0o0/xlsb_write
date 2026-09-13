//! xl/workbook.bin builder.
//!
//! For any sheet configuration, patches a fixed prefix + N×BrtBundleSh + fixed
//! suffix. Prefix/suffix are exact bytes from a valid reference Excel xlsb.
//! Adapted from the `xlsb-writer` crate (MIT License, Copyright (c) 2026
//! kotucha). Template-splicing is fine here short-term; will move to fully
//! programmatic construction once print areas are needed.
//!
//! **Defined names (`Workbook::define_name`/`StreamingWorkbook::define_name`,
//! 2026-09-13):** `WB_SUFFIX` was already carrying a `BrtBeginSupBook`(353)/
//! `BrtSupSelf`(357)/`BrtExternSheet`(362)/`BrtEndSupBook`(354) block and one
//! `BrtName`(39) record — an artifact of whatever real workbook the
//! original `xlsb-writer` template was extracted from (its name field
//! decodes as the literal text "Sheet1", not a recognizable Excel built-in
//! name, and its range covers 100,001 rows × 120 columns — almost certainly
//! a leftover defined name from that source file, not something Excel
//! writes automatically: a fresh multi-sheet reference file with no defined
//! names has none of these four record types at all, confirmed via COM
//! automation). That pre-existing record is left untouched here (removing
//! it is out of scope for this session and risks an unrelated regression);
//! this module's own defined-name support only *adds* to what's already
//! there.
//!
//! **How the shape was derived** (same real-Excel-reference method as
//! `drawing.rs`/`sheet.rs`'s `RID_BEGIN_AFILTER`, not spec text alone —
//! `BrtName`'s exact field layout isn't fully spelled out on the rendered
//! MS-XLSB HTML pages either): built three reference workbooks with real
//! Excel (COM automation, `Workbook.Names.Add`) — one single-sheet workbook
//! with a name on that one sheet, a 3-sheet workbook with no names (to
//! confirm SupBook/ExternSheet/Name are genuinely absent without any
//! defined name), and a 3-sheet workbook with a name referring to the
//! *third* sheet — and inspected all three with `examples/dump_sheet.rs`.
//!
//! Finding one: `BrtName`'s range is a full Rgce token stream (`cce`=15
//! always for a single-area reference), not a direct row/col field —
//! specifically `PtgArea3d` (ptg `0x3b`, reference class): `ixti`(u16) +
//! `rowFirst`(u32) + `rowLast`(u32) + `colFirst`(u16) + `colLast`(u16). A
//! workbook-scoped name (`Workbook.Names.Add` with no `SheetName!` prefix
//! on the name itself) still always resolves to a *specific* sheet via
//! `ixti`, even though the name's own `itab` field is `-1` (0xFFFFFFFF,
//! global scope) in every reference file — confirming this crate's own
//! `define_name` (workbook-level, not per-`Worksheet`) should always emit
//! `itab = -1`.
//!
//! Finding two: `ixti` indexes into `BrtExternSheet`'s XTI table, which is
//! NOT one entry per workbook sheet — the 3-sheet/name-on-sheet-3
//! reference has exactly one XTI entry, `{itabFirst: 2, itabLast: 2}`
//! (0-based), and the name's `ixti = 0` (the first and only entry). This
//! crate's own `BrtExternSheet` (rid 362) therefore only ever grows to
//! cover sheets actually referenced by a defined name: entry 0 is always
//! `itab = 0` (required — the pre-existing leftover `BrtName` above
//! hardcodes `ixti = 0`), and one more entry is appended per *additional*
//! distinct sheet index referenced by a `define_name` call.

use crate::biff12::{write_rec, write_wstr};

/// One workbook-level defined name — see `Workbook::define_name`/
/// `StreamingWorkbook::define_name`. Always workbook-scoped (`itab = -1`
/// in the encoded `BrtName`), even though the range itself lives on a
/// specific sheet (`sheet_index`, via `PtgArea3d`'s `ixti`) — see this
/// module's doc comment for why that's the correct real-Excel shape, not a
/// simplification.
pub(crate) struct DefinedName {
    pub name: String,
    pub sheet_index: usize,
    pub first_row: u32,
    pub first_col: u32,
    pub last_row: u32,
    pub last_col: u32,
}

const RID_NAME: u32 = 39;
const RID_EXTERN_SHEET: u32 = 362;
/// `PtgArea3d`, reference class ([MS-XLS]/[MS-XLSB] Ptg grammar) — a
/// sheet-qualified rectangular range reference, confirmed byte-for-byte
/// against real Excel-authored `BrtName` records (see module doc comment).
const PTG_AREA_3D: u8 = 0x3b;
/// Fixed 8-byte tail after `BrtName`'s `rgce` — confirmed byte-identical
/// across every real reference file this was checked against (single name,
/// name on a non-first sheet, differing name lengths): likely
/// cchCustMenu/cchDescription/cchHelpTopic/cchStatusText (all absent), but
/// the exact field semantics don't matter for a writer that never sets any
/// of them — only that the bytes match what real Excel emits.
const NAME_TAIL: &[u8] = &[0x00, 0x00, 0x00, 0x00, 0xff, 0xff, 0xff, 0xff];

const WB_PREFIX: &[u8] = &[
    0x83, 0x01, 0x00, 0x80, 0x01, 0x32, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x78, 0x00, 0x6c, 0x00, 0x01, 0x00, 0x00, 0x00, 0x37, 0x00, 0x01, 0x00,
    0x00, 0x00, 0x37, 0x00, 0x05, 0x00, 0x00, 0x00, 0x32, 0x00, 0x39, 0x00, 0x38, 0x00, 0x32, 0x00, 0x32, 0x00, 0x99,
    0x01, 0x0c, 0x20, 0x00, 0x01, 0x00, 0x42, 0xe5, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x25, 0x06, 0x01, 0x00, 0x03,
    0x0f, 0x00, 0x80, 0x97, 0x10, 0x58, 0x2a, 0x00, 0x00, 0x00, 0x43, 0x00, 0x3a, 0x00, 0x5c, 0x00, 0x55, 0x00, 0x73,
    0x00, 0x65, 0x00, 0x72, 0x00, 0x73, 0x00, 0x5c, 0x00, 0x61, 0x00, 0x6e, 0x00, 0x64, 0x00, 0x33, 0x00, 0x5c, 0x00,
    0x41, 0x00, 0x70, 0x00, 0x70, 0x00, 0x44, 0x00, 0x61, 0x00, 0x74, 0x00, 0x61, 0x00, 0x5c, 0x00, 0x4c, 0x00, 0x6f,
    0x00, 0x63, 0x00, 0x61, 0x00, 0x6c, 0x00, 0x5c, 0x00, 0x50, 0x00, 0x72, 0x00, 0x6f, 0x00, 0x67, 0x00, 0x72, 0x00,
    0x61, 0x00, 0x6d, 0x00, 0x73, 0x00, 0x5c, 0x00, 0x57, 0x00, 0x61, 0x00, 0x72, 0x00, 0x70, 0x00, 0x5c, 0x00, 0x26,
    0x00, 0x25, 0x06, 0x01, 0x00, 0x00, 0x10, 0x00, 0x80, 0x81, 0x18, 0x74, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x28, 0x00, 0x00, 0x00, 0x38, 0x00, 0x5f, 0x00, 0x7b, 0x00, 0x32, 0x00, 0x39, 0x00, 0x42, 0x00, 0x46, 0x00,
    0x41, 0x00, 0x37, 0x00, 0x39, 0x00, 0x44, 0x00, 0x2d, 0x00, 0x32, 0x00, 0x37, 0x00, 0x31, 0x00, 0x43, 0x00, 0x2d,
    0x00, 0x34, 0x00, 0x42, 0x00, 0x42, 0x00, 0x46, 0x00, 0x2d, 0x00, 0x39, 0x00, 0x39, 0x00, 0x36, 0x00, 0x41, 0x00,
    0x2d, 0x00, 0x30, 0x00, 0x43, 0x00, 0x39, 0x00, 0x36, 0x00, 0x36, 0x00, 0x32, 0x00, 0x39, 0x00, 0x33, 0x00, 0x43,
    0x00, 0x41, 0x00, 0x39, 0x00, 0x31, 0x00, 0x7d, 0x00, 0x2f, 0x00, 0x00, 0x00, 0x2f, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x26, 0x00, 0x87, 0x01, 0x00,
    0x25, 0x06, 0x01, 0x00, 0x02, 0x10, 0x00, 0x80, 0x80, 0x18, 0x10, 0x00, 0x00, 0x00, 0x00, 0x0d, 0x00, 0x00, 0x00,
    0xff, 0xff, 0xff, 0xff, 0x00, 0x00, 0x00, 0x00, 0x26, 0x00, 0x9e, 0x01, 0x1d, 0x89, 0x1c, 0x00, 0x00, 0x56, 0x13,
    0x00, 0x00, 0xbf, 0x5e, 0x00, 0x00, 0x5b, 0x3b, 0x00, 0x00, 0x58, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x78, 0x88, 0x01, 0x00, 0x8f, 0x01, 0x00,
];
const WB_SUFFIX: &[u8] = &[
    0x90, 0x01, 0x00, 0xe1, 0x02, 0x00, 0xe5, 0x02, 0x00, 0xea, 0x02, 0x10, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xe2, 0x02, 0x00, 0x27, 0x34, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xff, 0xff, 0xff, 0xff, 0x06, 0x00, 0x00, 0x00, 0x53, 0x00, 0x68, 0x00, 0x65, 0x00, 0x65, 0x00, 0x74, 0x00, 0x31,
    0x00, 0x0f, 0x00, 0x00, 0x00, 0x3b, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xa0, 0x86, 0x01, 0x00, 0x00, 0x00, 0x77,
    0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xff, 0xff, 0xff, 0x9d, 0x01, 0x1a, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00,
    0x00, 0x64, 0x00, 0x00, 0x00, 0xfc, 0xa9, 0xf1, 0xd2, 0x4d, 0x62, 0x50, 0x3f, 0x01, 0x00, 0x00, 0x00, 0x6a, 0x00,
    0x9b, 0x01, 0x01, 0x00, 0x84, 0x01, 0x00,
];

// `WB_SUFFIX` byte offsets (found by parsing it back with
// `biff12::read_vi`/record boundaries, not by hand-counting the array — see
// this module's doc comment for what each named span is):
//   [0..9)   BrtEndBundleShs(144) + BrtBeginSupBook(353) + BrtSupSelf(357),
//            each zero-length — unconditional, unrelated to defined names.
//   [9..28)  the pre-existing BrtExternSheet(362) record (header+payload),
//            REPLACED at runtime by `build_externsheet` below rather than
//            copied verbatim, since its XTI table must grow to cover any
//            additional sheet a `define_name` call references.
//   [28..31) BrtEndSupBook(354), zero-length.
//   [31..85) the pre-existing leftover BrtName(39) record — copied verbatim,
//            unconditionally, unchanged (see module doc comment).
//   [85..)   BrtCalcProp(157) onward through BrtEndBook(132) — unconditional
//            tail, unrelated to defined names.
const WB_SUFFIX_EXTERNSHEET_START: usize = 9;
const WB_SUFFIX_EXTERNSHEET_END: usize = 28;
const WB_SUFFIX_LEGACY_NAME_END: usize = 85;

fn encode_bundle_sh(tab_id: u32, rel_id: &str, name: &str) -> Vec<u8> {
    let mut pay = Vec::new();
    pay.extend_from_slice(&0u32.to_le_bytes()); // hsState = visible
    pay.extend_from_slice(&tab_id.to_le_bytes());
    write_wstr(rel_id, &mut pay);
    write_wstr(name, &mut pay);
    pay
}

/// Build `BrtExternSheet`'s XTI table and, for each `defined_names` entry,
/// the `ixti` value its `PtgArea3d` should use. Entry 0 is always `itab = 0`
/// — required by the pre-existing leftover `BrtName` record baked into
/// `WB_SUFFIX`, which hardcodes `ixti = 0` (see module doc comment) — and
/// one more entry is appended per *additional* distinct sheet index a
/// `define_name` call references, in order of first appearance. With no
/// defined names at all, this reproduces the exact same single-entry
/// (`itab = 0`) table `WB_SUFFIX` already hardcoded, so output for a
/// workbook with no defined names is byte-identical to before this feature
/// existed.
fn build_externsheet(defined_names: &[DefinedName]) -> (Vec<u8>, Vec<u16>) {
    let mut itabs: Vec<u32> = vec![0];
    let mut ixtis = Vec::with_capacity(defined_names.len());
    for dn in defined_names {
        let sheet_index = dn.sheet_index as u32;
        let ixti = itabs.iter().position(|&t| t == sheet_index).unwrap_or_else(|| {
            itabs.push(sheet_index);
            itabs.len() - 1
        });
        ixtis.push(ixti as u16);
    }
    let mut pay = Vec::with_capacity(4 + itabs.len() * 12);
    pay.extend_from_slice(&(itabs.len() as u32).to_le_bytes()); // cXti
    for &itab in &itabs {
        pay.extend_from_slice(&0u32.to_le_bytes()); // iSupBook = 0 (this workbook)
        pay.extend_from_slice(&itab.to_le_bytes()); // itabFirst
        pay.extend_from_slice(&itab.to_le_bytes()); // itabLast
    }
    (pay, ixtis)
}

/// Encode one `BrtName` record's payload — see module doc comment for how
/// this shape (including the fixed `NAME_TAIL`) was derived.
fn encode_name(name: &str, ixti: u16, first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> Vec<u8> {
    let mut pay = Vec::new();
    pay.extend_from_slice(&0u32.to_le_bytes()); // grbit — no hidden/builtin/function flags
    pay.push(0); // chKey — no keyboard shortcut
    pay.extend_from_slice(&(-1i32).to_le_bytes()); // itab = -1: workbook-scoped name
    let utf16: Vec<u16> = name.encode_utf16().collect();
    pay.extend_from_slice(&(utf16.len() as u32).to_le_bytes()); // cch
    for ch in &utf16 {
        pay.extend_from_slice(&ch.to_le_bytes());
    }
    let mut rgce = Vec::with_capacity(15);
    rgce.push(PTG_AREA_3D);
    rgce.extend_from_slice(&ixti.to_le_bytes());
    rgce.extend_from_slice(&first_row.to_le_bytes());
    rgce.extend_from_slice(&last_row.to_le_bytes());
    rgce.extend_from_slice(&crate::formula::col_rel_short(first_col).to_le_bytes());
    rgce.extend_from_slice(&crate::formula::col_rel_short(last_col).to_le_bytes());
    pay.extend_from_slice(&(rgce.len() as u32).to_le_bytes()); // cce
    pay.extend_from_slice(&rgce);
    pay.extend_from_slice(NAME_TAIL);
    pay
}

/// Build a complete workbook.bin for the given sheet names and defined
/// names. `defined_names` is normally empty — most workbooks don't use
/// `define_name` at all, and this function reproduces byte-identical
/// output to before this feature existed in that case.
pub fn build_workbook(names: &[&str], defined_names: &[DefinedName]) -> Vec<u8> {
    let mut out = Vec::with_capacity(WB_PREFIX.len() + names.len() * 50 + WB_SUFFIX.len() + defined_names.len() * 80);
    out.extend_from_slice(WB_PREFIX);
    for (i, name) in names.iter().enumerate() {
        let tab_id = (i + 1) as u32;
        let rel_id = format!("rId{}", tab_id);
        let pay = encode_bundle_sh(tab_id, &rel_id, name);
        write_rec(0x009C, &pay, &mut out);
    }

    out.extend_from_slice(&WB_SUFFIX[..WB_SUFFIX_EXTERNSHEET_START]);

    let (externsheet_payload, ixtis) = build_externsheet(defined_names);
    write_rec(RID_EXTERN_SHEET, &externsheet_payload, &mut out);

    out.extend_from_slice(&WB_SUFFIX[WB_SUFFIX_EXTERNSHEET_END..WB_SUFFIX_LEGACY_NAME_END]);

    for (dn, &ixti) in defined_names.iter().zip(&ixtis) {
        let pay = encode_name(&dn.name, ixti, dn.first_row, dn.first_col, dn.last_row, dn.last_col);
        write_rec(RID_NAME, &pay, &mut out);
    }

    out.extend_from_slice(&WB_SUFFIX[WB_SUFFIX_LEGACY_NAME_END..]);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn single_sheet1_is_531_bytes() {
        assert_eq!(build_workbook(&["Sheet1"], &[]).len(), 531);
    }

    #[test]
    fn multi_sheet_grows() {
        let wb2 = build_workbook(&["Sales", "Finance"], &[]);
        let wb3 = build_workbook(&["A", "B", "C"], &[]);
        assert!(wb2.len() > 531);
        assert!(wb3.len() > wb2.len());
    }

    /// With no defined names, output must be byte-identical to the
    /// pre-existing hardcoded `WB_SUFFIX` slice this function now rebuilds
    /// dynamically — i.e. `build_externsheet(&[])` must reproduce exactly
    /// the same 19 bytes `WB_SUFFIX[9..28]` already had.
    #[test]
    fn no_defined_names_reproduces_original_externsheet_bytes() {
        let mut expected = Vec::new();
        expected.extend_from_slice(&WB_SUFFIX[..WB_SUFFIX_EXTERNSHEET_START]);
        expected.extend_from_slice(&WB_SUFFIX[WB_SUFFIX_EXTERNSHEET_START..WB_SUFFIX_EXTERNSHEET_END]);
        expected.extend_from_slice(&WB_SUFFIX[WB_SUFFIX_EXTERNSHEET_END..]);

        let mut out = Vec::new();
        out.extend_from_slice(&WB_SUFFIX[..WB_SUFFIX_EXTERNSHEET_START]);
        let (pay, ixtis) = build_externsheet(&[]);
        assert!(ixtis.is_empty());
        write_rec(RID_EXTERN_SHEET, &pay, &mut out);
        out.extend_from_slice(&WB_SUFFIX[WB_SUFFIX_EXTERNSHEET_END..]);

        assert_eq!(out, expected);
    }

    /// One defined name on sheet 0 reuses XTI entry 0 (`ixti = 0`) — no new
    /// XTI entry needed, matching the single-sheet real-Excel reference
    /// file this was derived from.
    #[test]
    fn defined_name_on_sheet_zero_reuses_first_xti_entry() {
        let names = [DefinedName {
            name: "MyRange".to_owned(),
            sheet_index: 0,
            first_row: 0,
            first_col: 0,
            last_row: 1,
            last_col: 1,
        }];
        let (pay, ixtis) = build_externsheet(&names);
        assert_eq!(ixtis, vec![0]);
        assert_eq!(u32::from_le_bytes(pay[0..4].try_into().unwrap()), 1, "cXti must stay 1");
    }

    /// A defined name on a later sheet gets its own new XTI entry (index 1),
    /// while entry 0 (`itab = 0`) is preserved — matching the 3-sheet,
    /// name-on-sheet-3 real-Excel reference file this was derived from.
    #[test]
    fn defined_name_on_later_sheet_gets_new_xti_entry() {
        let names = [DefinedName {
            name: "ThirdSheetRange".to_owned(),
            sheet_index: 2,
            first_row: 1,
            first_col: 1,
            last_row: 2,
            last_col: 2,
        }];
        let (pay, ixtis) = build_externsheet(&names);
        assert_eq!(ixtis, vec![1]);
        assert_eq!(
            u32::from_le_bytes(pay[0..4].try_into().unwrap()),
            2,
            "cXti must grow to 2"
        );
        // Entry 0: itabFirst=itabLast=0 (bytes 4..16).
        assert_eq!(u32::from_le_bytes(pay[8..12].try_into().unwrap()), 0);
        // Entry 1: itabFirst=itabLast=2 (bytes 16..28).
        assert_eq!(u32::from_le_bytes(pay[20..24].try_into().unwrap()), 2);
        assert_eq!(u32::from_le_bytes(pay[24..28].try_into().unwrap()), 2);
    }

    /// Two defined names on the SAME later sheet must share one XTI entry,
    /// not allocate a new one each — real Excel's own XTI table dedupes by
    /// sheet, not by name.
    #[test]
    fn two_names_on_same_sheet_share_one_xti_entry() {
        let names = [
            DefinedName {
                name: "A".to_owned(),
                sheet_index: 1,
                first_row: 0,
                first_col: 0,
                last_row: 0,
                last_col: 0,
            },
            DefinedName {
                name: "B".to_owned(),
                sheet_index: 1,
                first_row: 1,
                first_col: 1,
                last_row: 1,
                last_col: 1,
            },
        ];
        let (pay, ixtis) = build_externsheet(&names);
        assert_eq!(ixtis, vec![1, 1]);
        assert_eq!(
            u32::from_le_bytes(pay[0..4].try_into().unwrap()),
            2,
            "cXti must be 2, not 3"
        );
    }

    /// `encode_name`'s payload must decode exactly the way real Excel's own
    /// `BrtName` records do (see module doc comment): grbit=0, chKey=0,
    /// itab=-1 (workbook scope), cch/name, cce=15, then a `PtgArea3d` token
    /// stream, then the fixed 8-byte tail.
    #[test]
    fn encode_name_matches_real_excel_field_layout() {
        let pay = encode_name("MyRange", 0, 0, 0, 1, 1);
        assert_eq!(&pay[0..5], &[0, 0, 0, 0, 0], "grbit(4)+chKey(1) must be zero");
        assert_eq!(i32::from_le_bytes(pay[5..9].try_into().unwrap()), -1, "itab must be -1");
        assert_eq!(u32::from_le_bytes(pay[9..13].try_into().unwrap()), 7, "cch must be 7");
        let name_utf16: Vec<u8> = "MyRange".encode_utf16().flat_map(u16::to_le_bytes).collect();
        assert_eq!(&pay[13..27], name_utf16.as_slice());
        assert_eq!(
            u32::from_le_bytes(pay[27..31].try_into().unwrap()),
            15,
            "cce must be 15"
        );
        let rgce = &pay[31..46];
        assert_eq!(rgce[0], PTG_AREA_3D);
        assert_eq!(u16::from_le_bytes(rgce[1..3].try_into().unwrap()), 0, "ixti");
        assert_eq!(u32::from_le_bytes(rgce[3..7].try_into().unwrap()), 0, "rowFirst");
        assert_eq!(u32::from_le_bytes(rgce[7..11].try_into().unwrap()), 1, "rowLast");
        assert_eq!(u16::from_le_bytes(rgce[11..13].try_into().unwrap()), 0, "colFirst");
        assert_eq!(u16::from_le_bytes(rgce[13..15].try_into().unwrap()), 1, "colLast");
        assert_eq!(&pay[46..54], NAME_TAIL);
    }

    /// End-to-end: `build_workbook` with one defined name must contain a
    /// well-formed `BrtName` record parseable back with `biff12`, sitting
    /// after the (dynamically-rebuilt) `BrtExternSheet`/`BrtEndSupBook` and
    /// after the pre-existing leftover name, before the rest of the fixed
    /// tail.
    #[test]
    fn build_workbook_with_defined_name_contains_expected_records() {
        let names = [DefinedName {
            name: "Totals".to_owned(),
            sheet_index: 0,
            first_row: 0,
            first_col: 0,
            last_row: 3,
            last_col: 2,
        }];
        let bytes = build_workbook(&["Sheet1"], &names);
        let recs = crate::biff12::try_parse_records(&bytes).expect("workbook.bin must be well-formed BIFF12");
        let name_recs: Vec<&Vec<u8>> = recs
            .iter()
            .filter(|(rid, _)| *rid == RID_NAME)
            .map(|(_, p)| p)
            .collect();
        // The pre-existing leftover name plus our new one.
        assert_eq!(name_recs.len(), 2);
        assert_eq!(
            *name_recs[1],
            encode_name("Totals", 0, 0, 0, 3, 2),
            "our new BrtName must come after the pre-existing leftover one"
        );
    }
}
