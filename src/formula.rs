//! Formula support: a small `Formula` builder + an Rgce/Ptg token encoder.
//!
//! This is not a text-formula parser — formulas are built programmatically
//! (`Formula::sum_range(...)`, `Formula::add(...)`, etc.) and rendered to a
//! BIFF12 Rgce token stream. Every byte layout below (PtgRef/PtgArea/
//! PtgFuncVar/PtgNum/PtgStr/PtgAdd.../RgceLocRel/ColRelShort/BrtFmlaNum) is
//! pulled from the published MS-XLSB spec. PtgStr in particular is NOT the
//! classic BIFF8 `ShortXLUnicodeString` (cch:1 + flags:1 + chars) — XLSB's
//! own PtgStr (2.5.98.88) is cch:2 + chars, no flags byte at all. An earlier
//! version of this file used the BIFF8 shape by mistake, which silently
//! corrupted the token stream for any formula containing a string literal
//! (Excel would discard the whole formula on load) — see
//! `write_short_xlunicode_string`'s doc comment for how that was found.
//!
//! Scope: numbers, strings, cell refs/ranges, `+ - * /`, function calls
//! (`SUM`, `AVERAGE`, ...) via `PtgFuncVar`, and `IF`/`IFERROR` via
//! `PtgAttrIf`/`PtgAttrGoto` branch tokens. The branch-token byte layout
//! (`ptg=0x19` + a bit-flag byte + a 2-byte offset, for both tokens) came
//! from the MS-XLS spec pages for `PtgAttrIf`/`PtgAttrGoto`; the offset
//! *semantics* (`PtgAttrIf.offset` = size of the "then" branch plus its
//! trailing `PtgAttrGoto`; each `PtgAttrGoto.offset` = size of everything
//! remaining in the if-expression, minus 1) are as specified there too.
//! `IFERROR(x, default)` is NOT encoded via `PtgAttrIfError` (a newer,
//! less-documented single-branch construct tied to a "Cetab" function
//! table for post-2007 functions) — instead it's desugared to
//! `IF(ISERROR(x), default, x)`, reusing the same well-verified `IF`
//! encoding and the classic (pre-2007, stable, well-documented) `Ftab`
//! index for `ISERROR` (0x0003).

use crate::biff12::{write_rec, write_wstr};

/// Builtin function index (`Ftab`), from the published MS-XLS spec.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FnIndex(pub(crate) u16);

impl FnIndex {
    pub const COUNT: FnIndex = FnIndex(0x0000);
    pub const ISNA: FnIndex = FnIndex(0x0002);
    pub const ISERROR: FnIndex = FnIndex(0x0003);
    pub const SUM: FnIndex = FnIndex(0x0004);
    pub const AVERAGE: FnIndex = FnIndex(0x0005);
    pub const MIN: FnIndex = FnIndex(0x0006);
    pub const MAX: FnIndex = FnIndex(0x0007);
    pub const ROUND: FnIndex = FnIndex(0x001B);
}

#[derive(Debug, Clone, PartialEq)]
pub enum Formula {
    Num(f64),
    Str(String),
    /// A single-cell reference, zero-based `(row, col)`.
    Ref(u32, u32),
    /// A rectangular range, zero-based `(first_row, first_col, last_row, last_col)`.
    Range(u32, u32, u32, u32),
    Add(Box<Formula>, Box<Formula>),
    Sub(Box<Formula>, Box<Formula>),
    Mul(Box<Formula>, Box<Formula>),
    Div(Box<Formula>, Box<Formula>),
    Lt(Box<Formula>, Box<Formula>),
    Le(Box<Formula>, Box<Formula>),
    Eq(Box<Formula>, Box<Formula>),
    Ge(Box<Formula>, Box<Formula>),
    Gt(Box<Formula>, Box<Formula>),
    Ne(Box<Formula>, Box<Formula>),
    /// A function call: `Ftab` index + arguments, encoded as `PtgFuncVar`.
    Func(FnIndex, Vec<Formula>),
    /// `IF(cond, then, else)`, encoded with `PtgAttrIf`/`PtgAttrGoto`
    /// short-circuit branch tokens (only one branch actually evaluates).
    If(Box<Formula>, Box<Formula>, Box<Formula>),
}

impl Formula {
    pub fn cell(row: u32, col: u32) -> Self {
        Formula::Ref(row, col)
    }
    pub fn range(first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> Self {
        Formula::Range(first_row, first_col, last_row, last_col)
    }
    pub fn num(v: f64) -> Self {
        Formula::Num(v)
    }
    pub fn str(s: &str) -> Self {
        Formula::Str(s.to_owned())
    }
    pub fn sum_range(first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> Self {
        Formula::Func(FnIndex::SUM, vec![Formula::range(first_row, first_col, last_row, last_col)])
    }
    pub fn add(self, rhs: Formula) -> Self {
        Formula::Add(Box::new(self), Box::new(rhs))
    }
    pub fn sub(self, rhs: Formula) -> Self {
        Formula::Sub(Box::new(self), Box::new(rhs))
    }
    pub fn mul(self, rhs: Formula) -> Self {
        Formula::Mul(Box::new(self), Box::new(rhs))
    }
    pub fn div(self, rhs: Formula) -> Self {
        Formula::Div(Box::new(self), Box::new(rhs))
    }
    pub fn lt(self, rhs: Formula) -> Self {
        Formula::Lt(Box::new(self), Box::new(rhs))
    }
    pub fn le(self, rhs: Formula) -> Self {
        Formula::Le(Box::new(self), Box::new(rhs))
    }
    pub fn eq(self, rhs: Formula) -> Self {
        Formula::Eq(Box::new(self), Box::new(rhs))
    }
    pub fn ge(self, rhs: Formula) -> Self {
        Formula::Ge(Box::new(self), Box::new(rhs))
    }
    pub fn gt(self, rhs: Formula) -> Self {
        Formula::Gt(Box::new(self), Box::new(rhs))
    }
    pub fn ne(self, rhs: Formula) -> Self {
        Formula::Ne(Box::new(self), Box::new(rhs))
    }
    pub fn if_then_else(cond: Formula, then: Formula, else_: Formula) -> Self {
        Formula::If(Box::new(cond), Box::new(then), Box::new(else_))
    }
    /// `IFERROR(expr, default)`, desugared to `IF(ISERROR(expr), default, expr)`
    /// — see the module doc comment for why. Note this evaluates `expr`
    /// twice (once inside `ISERROR`, once as the else-branch); fine for
    /// pure arithmetic on cell values, not appropriate for anything with
    /// side effects (not a concern for formulas built by this crate).
    pub fn iferror(expr: Formula, default: Formula) -> Self {
        Formula::If(
            Box::new(Formula::Func(FnIndex::ISERROR, vec![expr.clone()])),
            Box::new(default),
            Box::new(expr),
        )
    }

    /// Encode this formula to an Rgce token stream (postfix/RPN).
    pub(crate) fn encode(&self) -> Vec<u8> {
        let mut out = Vec::new();
        self.encode_into(&mut out);
        out
    }

    fn encode_into(&self, out: &mut Vec<u8>) {
        match self {
            Formula::Num(v) => {
                out.push(0x1F); // PtgNum: ptg(7 bits)=0x1F, reserved0(1 bit)=0
                out.extend_from_slice(&v.to_le_bytes());
            }
            Formula::Str(s) => {
                out.push(0x17); // PtgStr: ptg(7 bits)=0x17, reserved0(1 bit)=0
                write_short_xlunicode_string(s, out);
            }
            Formula::Ref(row, col) => {
                out.push(0x44); // PtgRef, class=VALUE(0x2): 0x04 | (0x2<<5)
                write_rgce_loc_rel(*row, *col, out);
            }
            Formula::Range(r0, c0, r1, c1) => {
                // PtgArea, class=REFERENCE(0x0): 0x05 | (0x0<<5). A cell
                // range is naturally a reference, not a value — encoding it
                // as value-class here is spec-legal (Excel opens it fine,
                // no repair) but doesn't match what real Excel writes for
                // this exact shape (confirmed byte-for-byte against a real
                // `SUM(range)` formula in a reference file), and Excel's
                // dynamic-array engine reacts to that mismatch by inserting
                // a spurious `@` (implicit intersection) into the formula
                // when the file is opened.
                out.push(0x25);
                // The two corners are NOT two interleaved (row,col) RgceLocRel
                // pairs (that was tried first and is wrong — confirmed by a
                // real reader turning it into `SUM($A$1:$A$131073)` instead
                // of `$A$1:$A$3`). XLSB's PtgArea groups fields by type, like
                // the classic BIFF RgceArea, just with row widened to 4
                // bytes: rowFirst, rowLast, colFirst, colLast.
                out.extend_from_slice(&r0.to_le_bytes());
                out.extend_from_slice(&r1.to_le_bytes());
                out.extend_from_slice(&(col_rel_short(*c0)).to_le_bytes());
                out.extend_from_slice(&(col_rel_short(*c1)).to_le_bytes());
            }
            Formula::Add(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x03); // PtgAdd
            }
            Formula::Sub(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x04); // PtgSub
            }
            Formula::Mul(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x05); // PtgMul
            }
            Formula::Div(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x06); // PtgDiv
            }
            Formula::Lt(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x09); // PtgLt
            }
            Formula::Le(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x0A); // PtgLe
            }
            Formula::Eq(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x0B); // PtgEq
            }
            Formula::Ge(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x0C); // PtgGe
            }
            Formula::Gt(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x0D); // PtgGt
            }
            Formula::Ne(a, b) => {
                a.encode_into(out);
                b.encode_into(out);
                out.push(0x0E); // PtgNe
            }
            Formula::Func(fn_idx, args) if *fn_idx == FnIndex::SUM && args.len() == 1 => {
                // `SUM(single-range)` is the one case real Excel encodes as
                // the dedicated `PtgAttrSum` shortcut (a plain reference
                // followed by this 4-byte control token) instead of the
                // general PtgFuncVar function-call form — confirmed
                // byte-for-byte against a real Excel-produced file. Using
                // PtgFuncVar here instead is spec-legal but, combined with
                // the reference-vs-value class mismatch this shortcut sits
                // on top of, is what triggered Excel's dynamic-array engine
                // to insert a spurious `@` into these formulas on load.
                args[0].encode_into(out);
                out.push(0x19); // PtgAttrSum
                out.push(0x10); // bitSum
                out.extend_from_slice(&[0x00, 0x00]); // unused
            }
            Formula::Func(fn_idx, args) => {
                for arg in args {
                    arg.encode_into(out);
                }
                out.push(0x42); // PtgFuncVar, class=VALUE(0x2): 0x02 | (0x2<<5)
                out.push(args.len() as u8); // cparams
                out.extend_from_slice(&fn_idx.0.to_le_bytes()); // tab (fCeFunc=0 implied, top bit unset)
            }
            Formula::If(cond, then, else_) => {
                cond.encode_into(out);
                let then_bytes = then.encode();
                let else_bytes = else_.encode();
                const GOTO_SIZE: u16 = 4;
                const FUNCVAR_SIZE: u16 = 4; // opcode(1) + cparams(1) + tab(2)

                // PtgAttrIf: skip past [then, Goto1] to reach `else` if cond is false.
                let offset_if = then_bytes.len() as u16 + GOTO_SIZE;
                out.push(0x19);
                out.push(0x02); // bitIf
                out.extend_from_slice(&offset_if.to_le_bytes());
                out.extend_from_slice(&then_bytes);

                // PtgAttrGoto #1: after `then` runs, skip past [else, Goto2, PtgFuncVar].
                let offset_goto1 = else_bytes.len() as u16 + GOTO_SIZE + FUNCVAR_SIZE - 1;
                out.push(0x19);
                out.push(0x08); // bitGoto
                out.extend_from_slice(&offset_goto1.to_le_bytes());
                out.extend_from_slice(&else_bytes);

                // PtgAttrGoto #2: after `else` runs, skip past [PtgFuncVar] — the
                // branch's result is already on the stack; PtgFuncVar(IF) below
                // is a structural marker a linear (non-branching) reader would
                // need, not something the branching evaluator actually executes.
                let offset_goto2 = FUNCVAR_SIZE - 1;
                out.push(0x19);
                out.push(0x08);
                out.extend_from_slice(&offset_goto2.to_le_bytes());

                out.push(0x42); // PtgFuncVar, class=VALUE
                out.push(3); // cparams
                out.extend_from_slice(&1u16.to_le_bytes()); // tab = IF
            }
        }
    }
}

/// `PtgStr`'s string field, per [MS-XLSB] 2.5.98.88/PtgStr: `cch` (2 bytes,
/// MUST be <= 255) + `rgch`, a plain array of 16-bit Unicode characters —
/// there is NO separate compressed/uncompressed flag byte here (that's the
/// classic BIFF8 `ShortXLUnicodeString` shape this was previously modeled
/// on, which does not apply to XLSB's PtgStr). Writing that extra flag byte
/// shifted the high byte of `cch` into it, so any reader decoding `cch` as
/// a real 2-byte little-endian integer saw a huge garbage length and
/// overran the token stream — this is what made Excel's loader discard the
/// whole formula ("Removed Records: Formula") for every formula containing
/// a string literal (`""`, `IFERROR` defaults, etc), confirmed by comparing
/// byte-for-byte against a real Excel-produced PtgStr in a reference file.
fn write_short_xlunicode_string(s: &str, buf: &mut Vec<u8>) {
    let utf16: Vec<u16> = s.encode_utf16().collect();
    debug_assert!(utf16.len() <= 255, "formula string literal too long for PtgStr (cch MUST be <= 255)");
    buf.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
    for ch in &utf16 {
        buf.extend_from_slice(&ch.to_le_bytes());
    }
}

/// `RgceLocRel`: row(4 bytes, absolute) + column(2 bytes: col(14 bits) |
/// fColRel(1 bit)=0 | fRwRel(1 bit)=0). Always absolute (no $ semantics
/// needed for programmatically-generated formulas — same calculated result
/// either way, simpler and avoids the signed-offset/wraparound rules that
/// apply when fColRel/fRwRel=1).
fn write_rgce_loc_rel(row: u32, col: u32, buf: &mut Vec<u8>) {
    buf.extend_from_slice(&row.to_le_bytes());
    buf.extend_from_slice(&col_rel_short(col).to_le_bytes());
}

/// `ColRelShort`: col(14 bits) | fColRel(1 bit)=0 | fRwRel(1 bit)=0 — always
/// absolute (see `write_rgce_loc_rel`'s doc comment for why).
fn col_rel_short(col: u32) -> u16 {
    (col as u16) & 0x3FFF
}

/// Write a `BrtFmlaNum` record: a formula cell whose most recent evaluation
/// produced a numeric value. `cached_value` is what viewers (including
/// calamine, and Excel before its own recalculation) display immediately.
pub fn write_fmla_num(col: u32, ixfe: u16, cached_value: f64, formula: &Formula, buf: &mut Vec<u8>) {
    let rgce = formula.encode();
    let mut pay = Vec::with_capacity(26 + rgce.len());
    pay.extend_from_slice(&col.to_le_bytes());
    pay.extend_from_slice(&(ixfe as u32).to_le_bytes()); // iStyleRef (low 24 bits) | fPhShow=0 | reserved=0
    pay.extend_from_slice(&cached_value.to_le_bytes());
    pay.extend_from_slice(&0u16.to_le_bytes()); // grbitFlags: fAlwaysCalc=0
    pay.extend_from_slice(&(rgce.len() as u32).to_le_bytes()); // cce
    pay.extend_from_slice(&rgce);
    pay.extend_from_slice(&0u32.to_le_bytes()); // cb (rgcb length) = 0
    write_rec(crate::biff12::RID_FMLA_NUM, &pay, buf);
}

/// Write a `BrtFmlaString` record: a formula cell whose most recent
/// evaluation produced a string value.
pub fn write_fmla_string(col: u32, ixfe: u16, cached_value: &str, formula: &Formula, buf: &mut Vec<u8>) {
    let rgce = formula.encode();
    let mut pay = Vec::with_capacity(10 + rgce.len());
    pay.extend_from_slice(&col.to_le_bytes());
    pay.extend_from_slice(&(ixfe as u32).to_le_bytes());
    write_wstr(cached_value, &mut pay); // string cached value: XLWideString (cch:4 + utf16)
    pay.extend_from_slice(&0u16.to_le_bytes()); // grbitFlags
    pay.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
    pay.extend_from_slice(&rgce);
    pay.extend_from_slice(&0u32.to_le_bytes());
    write_rec(crate::biff12::RID_FMLA_STRING, &pay, buf);
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::biff12::parse_records;

    /// Regression: a single-range `SUM` must be encoded the way real Excel
    /// encodes it — `PtgArea` (reference class) + `PtgAttrSum` — not
    /// `PtgArea` (value class) + `PtgFuncVar(SUM)`. The latter is spec-legal
    /// and opens without repair, but the class mismatch makes Excel's
    /// dynamic-array engine insert a spurious `@` into the formula on load.
    #[test]
    fn sum_range_encodes_expected_tokens() {
        let f = Formula::sum_range(1, 0, 10, 0); // SUM(A2:A11)
        let rgce = f.encode();
        // PtgArea(1+12=13) + PtgAttrSum(1+1+2=4) = 17 bytes
        assert_eq!(rgce.len(), 17);
        assert_eq!(rgce[0], 0x25); // PtgArea, reference class
        assert_eq!(rgce[13], 0x19); // PtgAttrSum
        assert_eq!(rgce[14], 0x10); // bitSum
    }

    /// Regression: PtgArea's two corners are NOT two interleaved (row,col)
    /// pairs — a real independent reader turned that encoding into
    /// `SUM($A$1:$A$131073)` instead of `$A$1:$A$3`. Fields are grouped by
    /// type: rowFirst, rowLast, colFirst, colLast.
    #[test]
    fn area_groups_fields_by_type_not_by_corner() {
        let f = Formula::range(5, 2, 9, 4); // rows 5..=9, cols 2..=4
        let rgce = f.encode();
        assert_eq!(rgce[0], 0x25); // PtgArea, reference class
        assert_eq!(u32::from_le_bytes(rgce[1..5].try_into().unwrap()), 5); // rowFirst
        assert_eq!(u32::from_le_bytes(rgce[5..9].try_into().unwrap()), 9); // rowLast
        assert_eq!(u16::from_le_bytes(rgce[9..11].try_into().unwrap()), 2); // colFirst
        assert_eq!(u16::from_le_bytes(rgce[11..13].try_into().unwrap()), 4); // colLast
    }

    /// Regression: PtgStr's `cch` is a 2-byte field with no flags byte
    /// (unlike the classic BIFF8 `ShortXLUnicodeString`). Writing a 1-byte
    /// cch + a flags byte shifted every following byte, so a real reader
    /// decoded `cch` as a huge garbage value and overran the token stream —
    /// Excel silently discarded the whole formula on load. Verified
    /// byte-for-byte against a real Excel-produced PtgStr in a reference
    /// file (see formula.rs's module doc comment).
    #[test]
    fn str_encodes_as_ptgstr_with_2byte_cch_and_no_flags_byte() {
        let rgce = Formula::str("hello").encode();
        assert_eq!(rgce[0], 0x17); // PtgStr
        assert_eq!(u16::from_le_bytes(rgce[1..3].try_into().unwrap()), 5); // cch, 2 bytes
        assert_eq!(&rgce[3..13], "hello".encode_utf16().flat_map(u16::to_le_bytes).collect::<Vec<u8>>().as_slice());
        assert_eq!(rgce.len(), 1 + 2 + 2 * 5); // ptg + cch(2) + chars, no flags byte
    }

    #[test]
    fn arithmetic_is_postfix() {
        // (A1 + 5) * 2
        let f = Formula::cell(0, 0).add(Formula::num(5.0)).mul(Formula::num(2.0));
        let rgce = f.encode();
        // PtgRef(7) PtgNum(9) PtgAdd(1) PtgNum(9) PtgMul(1) = 27 bytes, ends in PtgMul
        assert_eq!(rgce.len(), 7 + 9 + 1 + 9 + 1);
        assert_eq!(*rgce.last().unwrap(), 0x05); // PtgMul last (outermost op)
    }

    #[test]
    fn fmla_num_record_is_well_formed() {
        let mut buf = Vec::new();
        write_fmla_num(1, 0, 55.0, &Formula::sum_range(1, 0, 10, 1), &mut buf);
        let recs = parse_records(&buf);
        assert_eq!(recs.len(), 1);
        assert_eq!(recs[0].0, crate::biff12::RID_FMLA_NUM);
    }
}
