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
//!
//! **Cross-sheet references (2026-09-13, `Formula::sheet_cell`/
//! `sheet_range`):** a sheet-qualified reference (`Sheet2!A1`,
//! `Sheet2!B1:B5`) is `PtgRef3d`/`PtgArea3d` — the same shape as
//! `PtgRef`/`PtgArea` plus a 2-byte `ixti` field indexing into the
//! workbook's `BrtExternSheet` (rid 362) XTI table (see `wb_part.rs`,
//! which grew this table for `Workbook::define_name`'s `BrtName` support
//! first). Byte shape confirmed against a real Excel-authored `.xlsb`
//! (COM automation, three sheets, `Range.Formula = "=Sheet2!A1"` /
//! `"=SUM(Sheet2!B1:B5)"` / cross-references in both directions,
//! inspected with `examples/dump_sheet.rs`):
//! - `PtgRef3d` (value class `0x5A`, reference class `0x3A`): `ixti`(u16) +
//!   `row`(u32) + `col`(`ColRelShort`, u16) — 9 bytes total, one more field
//!   (`ixti`) than plain `PtgRef`.
//! - `PtgArea3d` (value class `0x5B`, reference class `0x3B`): `ixti`(u16) +
//!   `rowFirst`(u32) + `rowLast`(u32) + `colFirst`(`ColRelShort`) +
//!   `colLast`(`ColRelShort`) — 15 bytes total. This is the exact same
//!   token `wb_part::encode_name` already builds for a `BrtName`'s range
//!   (`PtgArea3d`, cce=15) — see `encode_ptg_area_3d` below, extracted so
//!   both call sites share one verified encoder instead of drifting apart.
//! - A single-range `SUM` over a cross-sheet range (`SUM(Sheet2!B1:B5)`)
//!   gets the exact same `PtgAttrSum` shortcut as a same-sheet
//!   `SUM(range)` (see the `Formula::Func`/`SUM` match arm below) —
//!   confirmed against the same reference file.
//! - Real Excel sets `ColRelShort`'s `fColRel`/`fRwRel` bits for a
//!   `Range.Formula`-typed cross-sheet reference (since typing `A1` without
//!   `$` is a *relative* reference by default) — this crate's own
//!   `Formula::sheet_cell`/`sheet_range` deliberately do NOT reproduce that
//!   (same choice already made for same-sheet `Formula::Ref`/`Formula::Range`
//!   — see `write_rgce_loc_rel`'s doc comment): always absolute
//!   (`col_rel_short`, bits clear), which computes the identical result
//!   without the signed-offset/wraparound rules relative encoding implies,
//!   and was confirmed to still open with no repair dialog and compute the
//!   correct value via the same real-Excel-COM round-trip method as every
//!   other formula feature in this crate.
//! - `ixti` itself is workbook-wide shared, incrementally-growing state,
//!   NOT something `Formula::encode` can resolve on its own (a `Formula` is
//!   built independently of which `Workbook`/`Worksheet` it ends up written
//!   into) — see `SheetRegistry` below for how this crate threads that
//!   state through from `Worksheet::write_formula_num`/... down to
//!   `Formula::encode_into`, and `wb_part.rs`'s module doc comment for the
//!   companion piece (how `BrtExternSheet` itself is finally emitted).

use crate::biff12::{write_rec, write_wstr};

/// Builtin function index (`Ftab`), from the published MS-XLS spec.
///
/// The constants below (`COUNT` through `ROUND`) are the only functions
/// this crate's own test suite has verified byte-for-byte against real
/// Excel output — use `Formula::Func` with one of them whenever possible.
///
/// The inner field is public so a caller who needs a function not listed
/// here (`VLOOKUP`, `SUMIF`, string/date functions, ...) can construct
/// `FnIndex(raw_ftab_index)` directly rather than forking this crate —
/// but doing so is **unverified by this crate**: look the real index up
/// in the published MS-XLS `Ftab` enumeration yourself (don't guess), and
/// double-check `PtgFuncVar`'s `cparams` byte (written from
/// `args.len()` in `Formula::Func`'s encoding) actually matches the
/// function's real required argument count — a mismatch there is exactly
/// the class of subtle bug this crate has hit before with its own
/// built-in functions.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct FnIndex(pub u16);

impl FnIndex {
    /// Verified against real Excel output by this crate's own tests.
    pub const COUNT: FnIndex = FnIndex(0x0000);
    pub const ISNA: FnIndex = FnIndex(0x0002);
    pub const ISERROR: FnIndex = FnIndex(0x0003);
    pub const SUM: FnIndex = FnIndex(0x0004);
    pub const AVERAGE: FnIndex = FnIndex(0x0005);
    pub const MIN: FnIndex = FnIndex(0x0006);
    pub const MAX: FnIndex = FnIndex(0x0007);
    pub const ROUND: FnIndex = FnIndex(0x001B);

    // ── 2026-09-13 formula-coverage expansion ──────────────────────────
    //
    // Every index below was cross-checked against Apache POI's
    // `functionMetadata.txt` (itself sourced from the published BIFF Ftab
    // enumeration) AND confirmed byte-for-byte against a real
    // Excel-authored `.xlsb` built via COM automation (`Range.Formula =
    // "=AND(...)"` etc., `SaveAs` format 50, inspected with
    // `examples/dump_sheet.rs`) — see the doc comment on `Formula::Func`
    // vs `Formula::FuncFixed` below for why some of these are one or the
    // other. `AND`/`OR`/`SUMIF`/`LEFT`/`RIGHT`/`VLOOKUP`/`INDEX`/`MATCH`/
    // `CONCATENATE` are real Excel's *variable*-argument-count functions
    // (encoded via the general `PtgFuncVar` form, like `SUM`/`AVERAGE`
    // above); `NOT`/`COUNTIF`/`MID`/`LEN`/`TEXT`/`TODAY`/`NOW`/`DATE` are
    // real Excel's *fixed*-argument-count functions (encoded via the
    // narrower `PtgFunc` form — 3 bytes, no `cparams` byte at all) — this
    // was NOT guessable from the argument count alone (`COUNTIF` and
    // `SUMIF` both commonly take 2 args, but only `COUNTIF`'s is fixed;
    // real Excel's own encoder chose differently for each, confirmed by
    // direct byte inspection).
    pub const CONCATENATE: FnIndex = FnIndex(0x0150); // 336, variable (0-30 args)
    pub const AND: FnIndex = FnIndex(0x0024); // 36, variable (1-30 args)
    pub const OR: FnIndex = FnIndex(0x0025); // 37, variable (1-30 args)
    pub const NOT: FnIndex = FnIndex(0x0026); // 38, fixed (1 arg) — PtgFunc
    pub const SUMIF: FnIndex = FnIndex(0x0159); // 345, variable (2-3 args)
    pub const COUNTIF: FnIndex = FnIndex(0x015A); // 346, fixed (2 args) — PtgFunc
    pub const LEFT: FnIndex = FnIndex(0x0073); // 115, variable (1-2 args)
    pub const RIGHT: FnIndex = FnIndex(0x0074); // 116, variable (1-2 args)
    pub const MID: FnIndex = FnIndex(0x001F); // 31, fixed (3 args) — PtgFunc
    pub const LEN: FnIndex = FnIndex(0x0020); // 32, fixed (1 arg) — PtgFunc
    pub const TEXT: FnIndex = FnIndex(0x0030); // 48, fixed (2 args) — PtgFunc
    /// Fixed (0 args), `PtgFunc` — and, unlike every other function here,
    /// volatile: real Excel wraps the call in a `PtgAttrSemi` marker token
    /// and sets a bit in the cell record's `grbitFlags` so it's
    /// recalculated on every pass, not just when a dependency changes.
    /// `Formula::today()`/`Formula::now()` reproduce both; see
    /// `Formula::is_volatile`.
    pub const TODAY: FnIndex = FnIndex(0x00DD); // 221, fixed (0 args), volatile
    /// See `TODAY`'s doc comment — same volatile handling.
    pub const NOW: FnIndex = FnIndex(0x004A); // 74, fixed (0 args), volatile
    pub const DATE: FnIndex = FnIndex(0x0041); // 65, fixed (3 args) — PtgFunc
    pub const VLOOKUP: FnIndex = FnIndex(0x0066); // 102, variable (3-4 args)
    pub const INDEX: FnIndex = FnIndex(0x001D); // 29, variable (2-4 args)
    pub const MATCH: FnIndex = FnIndex(0x0040); // 64, variable (2-3 args)
}

#[derive(Debug, Clone, PartialEq)]
pub enum Formula {
    Num(f64),
    Str(String),
    /// A boolean literal, encoded as `PtgBool` (opcode `0x1D`, one byte
    /// `0`/`1` after it — confirmed byte-for-byte against real Excel's
    /// encoding of `VLOOKUP(...,FALSE)`'s 4th argument).
    Bool(bool),
    /// A single-cell reference, zero-based `(row, col)`.
    Ref(u32, u32),
    /// A rectangular range, zero-based `(first_row, first_col, last_row, last_col)`.
    Range(u32, u32, u32, u32),
    /// A single-cell reference on ANOTHER sheet, encoded as `PtgRef3d`
    /// (value class `0x5A`) — `(sheet_name, row, col)`. Constructed via
    /// `Formula::sheet_cell`; see the module doc comment's "Cross-sheet
    /// references" section for the byte shape and how `ixti` (the index
    /// into `ixti`'s field into the workbook's `BrtExternSheet` XTI table)
    /// gets resolved at encode time via `SheetRegistry`.
    Ref3d(String, u32, u32),
    /// A rectangular range on ANOTHER sheet, encoded as `PtgArea3d`
    /// (reference class `0x3B`) — `(sheet_name, first_row, first_col,
    /// last_row, last_col)`. Constructed via `Formula::sheet_range`. See
    /// `Ref3d`'s doc comment.
    Range3d(String, u32, u32, u32, u32),
    Add(Box<Formula>, Box<Formula>),
    Sub(Box<Formula>, Box<Formula>),
    Mul(Box<Formula>, Box<Formula>),
    Div(Box<Formula>, Box<Formula>),
    /// `&` string concatenation, encoded as the binary infix `PtgConcat`
    /// (opcode `0x08`) — confirmed byte-for-byte against real Excel's
    /// encoding of `=A1&B1` (same shape as `Add`/`Sub`/`Mul`/`Div` above,
    /// just a different trailing opcode byte).
    Concat(Box<Formula>, Box<Formula>),
    Lt(Box<Formula>, Box<Formula>),
    Le(Box<Formula>, Box<Formula>),
    Eq(Box<Formula>, Box<Formula>),
    Ge(Box<Formula>, Box<Formula>),
    Gt(Box<Formula>, Box<Formula>),
    Ne(Box<Formula>, Box<Formula>),
    /// A function call with a *variable* argument count: `Ftab` index +
    /// arguments, encoded as `PtgFuncVar` (opcode `0x42`, carries a
    /// `cparams` byte = `args.len()`). Use for `SUM`, `AVERAGE`,
    /// `CONCATENATE`, `AND`, `OR`, `SUMIF`, `LEFT`, `RIGHT`, `VLOOKUP`,
    /// `INDEX`, `MATCH` — every function real Excel itself encodes this
    /// way (confirmed byte-for-byte; see `FnIndex`'s doc comments).
    Func(FnIndex, Vec<Formula>),
    /// A function call with a real, *fixed* argument count, encoded as
    /// `PtgFunc` (opcode `0x41`, 3 bytes total — no `cparams` byte at
    /// all, since the reader looks the arg count up from the `Ftab` entry
    /// itself). Use for `NOT`, `COUNTIF`, `MID`, `LEN`, `TEXT`, `TODAY`,
    /// `NOW`, `DATE` — real Excel encodes these with `PtgFunc`, not
    /// `PtgFuncVar`, confirmed byte-for-byte; see `FnIndex`'s doc
    /// comments. Getting this distinction wrong (e.g. using `Func`
    /// instead of `FuncFixed` for `COUNTIF`) is spec-adjacent but not what
    /// real Excel emits — exactly the class of mismatch that turned out to
    /// matter for `SUM`'s `PtgAttrSum` shortcut, so this crate's own
    /// verified constructors always pick the one real Excel actually uses.
    FuncFixed(FnIndex, Vec<Formula>),
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
        Formula::Func(
            FnIndex::SUM,
            vec![Formula::range(first_row, first_col, last_row, last_col)],
        )
    }

    /// A single-cell reference on another sheet, e.g.
    /// `Formula::sheet_cell("Data", 0, 0)` for `=Data!A1`. `sheet_name` must
    /// name a sheet already added to the workbook (via `add_worksheet`/
    /// `new_worksheet`/`new_worksheet_sized`) by the time this formula is
    /// actually written to a cell — resolution happens then, not here (a
    /// `Formula` is built independently of any `Workbook`); see
    /// `SheetRegistry::resolve_by_name`'s panic message for what happens
    /// otherwise.
    pub fn sheet_cell(sheet_name: &str, row: u32, col: u32) -> Self {
        Formula::Ref3d(sheet_name.to_owned(), row, col)
    }
    /// A rectangular range on another sheet, e.g.
    /// `Formula::sheet_range("Data", 0, 0, 9, 0)` for `=Data!A1:A10`. Same
    /// sheet-must-already-exist contract as `Formula::sheet_cell`.
    pub fn sheet_range(sheet_name: &str, first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> Self {
        Formula::Range3d(sheet_name.to_owned(), first_row, first_col, last_row, last_col)
    }
    /// `SUM` over a range on another sheet, e.g. `=SUM(Data!B1:B10)` — gets
    /// the same `PtgAttrSum` shortcut as `Formula::sum_range` (see the
    /// `Formula::Func`/`SUM` match arm in `encode_into`), confirmed against
    /// the same real-Excel reference file as `Formula::sheet_cell`'s doc
    /// comment.
    pub fn sum_sheet_range(sheet_name: &str, first_row: u32, first_col: u32, last_row: u32, last_col: u32) -> Self {
        Formula::Func(
            FnIndex::SUM,
            vec![Formula::sheet_range(
                sheet_name, first_row, first_col, last_row, last_col,
            )],
        )
    }
    // `add`/`sub`/`mul`/`div` deliberately name-match `std::ops` (this is a
    // formula-expression builder, not an arithmetic type — `Formula` isn't
    // meant to implement `Add`/`Sub`/`Mul`/`Div` itself, so there's no
    // actual trait-confusion risk despite the lint).
    #[allow(clippy::should_implement_trait)]
    pub fn add(self, rhs: Formula) -> Self {
        Formula::Add(Box::new(self), Box::new(rhs))
    }
    #[allow(clippy::should_implement_trait)]
    pub fn sub(self, rhs: Formula) -> Self {
        Formula::Sub(Box::new(self), Box::new(rhs))
    }
    #[allow(clippy::should_implement_trait)]
    pub fn mul(self, rhs: Formula) -> Self {
        Formula::Mul(Box::new(self), Box::new(rhs))
    }
    #[allow(clippy::should_implement_trait)]
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

    pub fn boolean(b: bool) -> Self {
        Formula::Bool(b)
    }

    /// `A1&B1` — the `&` string-concatenation operator (`PtgConcat`), not
    /// the `CONCATENATE(...)` function. See `Formula::concatenate` for the
    /// variadic function form.
    pub fn concat(self, rhs: Formula) -> Self {
        Formula::Concat(Box::new(self), Box::new(rhs))
    }

    /// `CONCATENATE(args...)` — the variadic function form. For two plain
    /// values, prefer `a.concat(b)` (the `&` operator): it's what Excel's
    /// own UI produces for that case and is one token shorter.
    pub fn concatenate(args: Vec<Formula>) -> Self {
        Formula::Func(FnIndex::CONCATENATE, args)
    }

    /// `AND(args...)` — variadic (1-30 args in real Excel).
    pub fn and(args: Vec<Formula>) -> Self {
        Formula::Func(FnIndex::AND, args)
    }
    /// `OR(args...)` — variadic (1-30 args in real Excel).
    pub fn or(args: Vec<Formula>) -> Self {
        Formula::Func(FnIndex::OR, args)
    }
    /// `NOT(expr)` — always exactly 1 argument in real Excel.
    #[allow(clippy::should_implement_trait)]
    pub fn not(expr: Formula) -> Self {
        Formula::FuncFixed(FnIndex::NOT, vec![expr])
    }

    /// `SUMIF(range, criteria)`. `criteria` is typically
    /// `Formula::str(">10")`/`Formula::str("apples")` or `Formula::num(10.0)`
    /// — real Excel encodes a comparison-expression criteria as a plain
    /// string literal (`PtgStr`), confirmed byte-for-byte against
    /// `SUMIF(A1:A10,">3")`. Real Excel's `SUMIF` also accepts an optional
    /// 3rd `sum_range` argument (2-3 args total, variable); reach for it
    /// with `Formula::Func(FnIndex::SUMIF, vec![range, criteria, sum_range])`
    /// directly if you need it — not separately byte-verified by this
    /// crate (only the 2-arg form was checked against real Excel output),
    /// but it follows the exact same verified variable-argument
    /// `PtgFuncVar` shape as the 2-arg form checked here.
    pub fn sumif(range: Formula, criteria: Formula) -> Self {
        Formula::Func(FnIndex::SUMIF, vec![range, criteria])
    }
    /// `COUNTIF(range, criteria)` — always exactly 2 arguments in real
    /// Excel (unlike `SUMIF`, which optionally takes a 3rd).
    pub fn countif(range: Formula, criteria: Formula) -> Self {
        Formula::FuncFixed(FnIndex::COUNTIF, vec![range, criteria])
    }

    /// `LEFT(text, num_chars)` — real Excel allows `num_chars` to be
    /// omitted (defaulting to 1), making this a variable 1-2 arg function;
    /// this constructor always passes both.
    pub fn left(text: Formula, num_chars: Formula) -> Self {
        Formula::Func(FnIndex::LEFT, vec![text, num_chars])
    }
    /// `RIGHT(text, num_chars)` — see `Formula::left`'s note on arity.
    pub fn right(text: Formula, num_chars: Formula) -> Self {
        Formula::Func(FnIndex::RIGHT, vec![text, num_chars])
    }
    /// `MID(text, start_num, num_chars)` — always exactly 3 arguments in
    /// real Excel.
    pub fn mid(text: Formula, start_num: Formula, num_chars: Formula) -> Self {
        Formula::FuncFixed(FnIndex::MID, vec![text, start_num, num_chars])
    }
    /// `LEN(text)` — always exactly 1 argument in real Excel.
    pub fn len(text: Formula) -> Self {
        Formula::FuncFixed(FnIndex::LEN, vec![text])
    }
    /// `TEXT(value, format_text)` — always exactly 2 arguments in real
    /// Excel. `format_text` is a normal `Formula::str(...)` argument (a
    /// number-format code like `"0.00"`), encoded as a plain `PtgStr`.
    pub fn text(value: Formula, format_text: Formula) -> Self {
        Formula::FuncFixed(FnIndex::TEXT, vec![value, format_text])
    }

    /// `TODAY()` — takes no arguments. Volatile: see `FnIndex::TODAY`'s
    /// doc comment for the `PtgAttrSemi` wrapper and cell-level
    /// recalculation flag this crate emits so Excel actually keeps this
    /// up to date, rather than freezing at whatever `cached_value` was
    /// supplied at write time.
    pub fn today() -> Self {
        Formula::FuncFixed(FnIndex::TODAY, vec![])
    }
    /// `NOW()` — see `Formula::today`'s note; same volatile handling.
    pub fn now() -> Self {
        Formula::FuncFixed(FnIndex::NOW, vec![])
    }
    /// `DATE(year, month, day)` — always exactly 3 arguments, and (unlike
    /// `TODAY`/`NOW`) not volatile.
    pub fn date(year: Formula, month: Formula, day: Formula) -> Self {
        Formula::FuncFixed(FnIndex::DATE, vec![year, month, day])
    }

    /// `VLOOKUP(lookup_value, table_array, col_index_num, range_lookup)`.
    /// `range_lookup` is a plain Rust `bool` here (not a `Formula`) since
    /// real Excel's own 4th argument is just a boolean literal in the
    /// overwhelming common case (`TRUE`/`FALSE`) — pass it through
    /// `Formula::Bool` if you need a computed 4th argument instead, via
    /// `Formula::Func(FnIndex::VLOOKUP, vec![...])` directly.
    pub fn vlookup(lookup_value: Formula, table_array: Formula, col_index_num: Formula, range_lookup: bool) -> Self {
        Formula::Func(
            FnIndex::VLOOKUP,
            vec![lookup_value, table_array, col_index_num, Formula::Bool(range_lookup)],
        )
    }
    /// `INDEX(array, row_num, column_num)` — real Excel's `INDEX` also
    /// supports a 1-arg area-only form and a 4th `area_num` argument
    /// (2-4 args total, variable); this constructor covers the common
    /// 3-arg form, which is what was byte-verified against real Excel
    /// output.
    pub fn index(array: Formula, row_num: Formula, column_num: Formula) -> Self {
        Formula::Func(FnIndex::INDEX, vec![array, row_num, column_num])
    }
    /// `MATCH(lookup_value, lookup_array, match_type)`. Named `match_`
    /// (trailing underscore) since `match` is a Rust keyword.
    pub fn match_(lookup_value: Formula, lookup_array: Formula, match_type: Formula) -> Self {
        Formula::Func(FnIndex::MATCH, vec![lookup_value, lookup_array, match_type])
    }

    /// Whether this formula (or any subexpression) calls a volatile
    /// function (`TODAY`/`NOW` so far). Used by `write_fmla_num`/
    /// `write_fmla_string`/`write_fmla_bool` to set the cell record's
    /// recalculate-always bit — see `FnIndex::TODAY`'s doc comment.
    pub(crate) fn is_volatile(&self) -> bool {
        match self {
            Formula::FuncFixed(fn_idx, args) => {
                *fn_idx == FnIndex::TODAY || *fn_idx == FnIndex::NOW || args.iter().any(Formula::is_volatile)
            }
            Formula::Func(_, args) => args.iter().any(Formula::is_volatile),
            Formula::Add(a, b)
            | Formula::Sub(a, b)
            | Formula::Mul(a, b)
            | Formula::Div(a, b)
            | Formula::Concat(a, b)
            | Formula::Lt(a, b)
            | Formula::Le(a, b)
            | Formula::Eq(a, b)
            | Formula::Ge(a, b)
            | Formula::Gt(a, b)
            | Formula::Ne(a, b) => a.is_volatile() || b.is_volatile(),
            Formula::If(c, t, e) => c.is_volatile() || t.is_volatile() || e.is_volatile(),
            Formula::Num(_)
            | Formula::Str(_)
            | Formula::Bool(_)
            | Formula::Ref(_, _)
            | Formula::Range(_, _, _, _)
            | Formula::Ref3d(_, _, _)
            | Formula::Range3d(_, _, _, _, _) => false,
        }
    }

    /// Encode this formula to an Rgce token stream (postfix/RPN), with no
    /// workbook context. Safe for any formula that contains no
    /// `Formula::sheet_cell`/`sheet_range` (`Ref3d`/`Range3d`) node — i.e.
    /// every formula this crate's own test suite built before cross-sheet
    /// references existed. Real cell writes always go through
    /// `encode_with`, called from `write_fmla_num`/`write_fmla_string`/
    /// `write_fmla_bool` with the workbook's real, shared `SheetRegistry` —
    /// this zero-argument convenience exists only so tests that don't
    /// involve cross-sheet references don't need to construct a throwaway
    /// registry just to call `encode`. `#[cfg(test)]` since production code
    /// never calls it (a `Ref3d`/`Range3d` node encoded this way would get
    /// its `ixti` from a private, throwaway registry — always disjoint from
    /// whatever real `BrtExternSheet` table the actual workbook ends up
    /// building — so this must never reach an actual cell write).
    #[cfg(test)]
    pub(crate) fn encode(&self) -> Vec<u8> {
        self.encode_with(&mut SheetRegistry::new())
    }

    /// Encode this formula to an Rgce token stream (postfix/RPN), resolving
    /// any `Ref3d`/`Range3d` node's sheet name to an `ixti` via `xti` —
    /// registering a new `BrtExternSheet` entry the first time a given
    /// sheet is referenced, exactly like `Workbook::define_name`'s own
    /// `BrtName` range does (see `SheetRegistry`'s doc comment). This is
    /// the real, production encode path — `write_fmla_num`/
    /// `write_fmla_string`/`write_fmla_bool` all go through it.
    pub(crate) fn encode_with(&self, xti: &mut SheetRegistry) -> Vec<u8> {
        let mut out = Vec::new();
        self.encode_into(xti, &mut out);
        out
    }

    fn encode_into(&self, xti: &mut SheetRegistry, out: &mut Vec<u8>) {
        match self {
            Formula::Num(v) => {
                out.push(0x1F); // PtgNum: ptg(7 bits)=0x1F, reserved0(1 bit)=0
                out.extend_from_slice(&v.to_le_bytes());
            }
            Formula::Str(s) => {
                out.push(0x17); // PtgStr: ptg(7 bits)=0x17, reserved0(1 bit)=0
                write_short_xlunicode_string(s, out);
            }
            Formula::Bool(b) => {
                // PtgBool: ptg=0x1D + one byte (0/1). Confirmed byte-for-byte
                // against real Excel's `VLOOKUP(...,FALSE)` 4th argument.
                out.push(0x1D);
                out.push(*b as u8);
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
            Formula::Ref3d(sheet_name, row, col) => {
                // PtgRef3d, value class (0x1A base | (0x2<<5) = 0x5A) —
                // confirmed byte-for-byte against real Excel's encoding of
                // a whole-formula cross-sheet cell reference (`=Sheet2!A1`)
                // in `examples/dump_sheet.rs`-inspected COM-automation
                // reference file (see module doc comment). Same
                // `ixti`+`row`+`col` field order/sizes as `PtgRef`, with
                // `ixti` (resolved via the shared `SheetRegistry`) inserted
                // right after the opcode.
                let ixti = xti.resolve_by_name(sheet_name);
                out.push(0x5A);
                out.extend_from_slice(&ixti.to_le_bytes());
                write_rgce_loc_rel(*row, *col, out);
            }
            Formula::Range3d(sheet_name, r0, c0, r1, c1) => {
                // PtgArea3d, reference class (0x1B base | (0x0<<5) = 0x3B) —
                // the exact same token `wb_part::encode_name` already builds
                // for a `BrtName`'s range; see `encode_ptg_area_3d` (shared
                // by both call sites) and the module doc comment.
                let ixti = xti.resolve_by_name(sheet_name);
                encode_ptg_area_3d(ixti, *r0, *c0, *r1, *c1, out);
            }
            Formula::Add(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x03); // PtgAdd
            }
            Formula::Sub(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x04); // PtgSub
            }
            Formula::Mul(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x05); // PtgMul
            }
            Formula::Div(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x06); // PtgDiv
            }
            Formula::Concat(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x08); // PtgConcat — confirmed byte-for-byte against real Excel's `=A1&B1`.
            }
            Formula::Lt(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x09); // PtgLt
            }
            Formula::Le(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x0A); // PtgLe
            }
            Formula::Eq(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x0B); // PtgEq
            }
            Formula::Ge(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x0C); // PtgGe
            }
            Formula::Gt(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
                out.push(0x0D); // PtgGt
            }
            Formula::Ne(a, b) => {
                a.encode_into(xti, out);
                b.encode_into(xti, out);
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
                // to insert a spurious `@` into these formulas on load. This
                // arm matches regardless of whether `args[0]` is a same-sheet
                // `Range` or a cross-sheet `Range3d` — `SUM(Sheet2!B1:B5)`
                // gets the identical shortcut, confirmed against the same
                // reference file (see module doc comment).
                args[0].encode_into(xti, out);
                out.push(0x19); // PtgAttrSum
                out.push(0x10); // bitSum
                out.extend_from_slice(&[0x00, 0x00]); // unused
            }
            Formula::Func(fn_idx, args) => {
                for arg in args {
                    arg.encode_into(xti, out);
                }
                out.push(0x42); // PtgFuncVar, class=VALUE(0x2): 0x02 | (0x2<<5)
                out.push(args.len() as u8); // cparams
                out.extend_from_slice(&fn_idx.0.to_le_bytes()); // tab (fCeFunc=0 implied, top bit unset)
            }
            Formula::FuncFixed(fn_idx, args) if *fn_idx == FnIndex::TODAY || *fn_idx == FnIndex::NOW => {
                // Volatile functions: real Excel prepends a PtgAttrSemi
                // marker (flags=bitSemi=0x01, offset=0) before the call —
                // confirmed byte-for-byte against real Excel's encoding of
                // `=TODAY()`/`=NOW()`. `write_fmla_num`/`write_fmla_string`
                // separately set a matching bit in the cell record's
                // `grbitFlags` (via `Formula::is_volatile`); both were
                // needed for Excel to actually recalculate these on open
                // rather than freezing at the supplied cached value forever
                // (no cell dependency would otherwise ever trigger a
                // recalc).
                out.push(0x19); // PtgAttrSemi
                out.push(0x01); // bitSemi
                out.extend_from_slice(&[0x00, 0x00]); // offset = 0 (no args to skip)
                for arg in args {
                    arg.encode_into(xti, out);
                }
                out.push(0x41); // PtgFunc, class=VALUE — confirmed byte-for-byte
                out.extend_from_slice(&fn_idx.0.to_le_bytes()); // iftab, no cparams byte
            }
            Formula::FuncFixed(fn_idx, args) => {
                // PtgFunc: a fixed-argument-count function call — real
                // Excel's encoder for `NOT`/`COUNTIF`/`MID`/`LEN`/`TEXT`/
                // `DATE` (confirmed byte-for-byte against each). Unlike
                // `PtgFuncVar` above, there is no `cparams` byte at all:
                // the reader determines the argument count from the
                // `Ftab`/`iftab` entry itself, not from anything in the
                // token stream.
                for arg in args {
                    arg.encode_into(xti, out);
                }
                out.push(0x41); // PtgFunc, class=VALUE
                out.extend_from_slice(&fn_idx.0.to_le_bytes()); // iftab
            }
            Formula::If(cond, then, else_) => {
                cond.encode_into(xti, out);
                let then_bytes = then.encode_with(xti);
                let else_bytes = else_.encode_with(xti);
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
    debug_assert!(
        utf16.len() <= 255,
        "formula string literal too long for PtgStr (cch MUST be <= 255)"
    );
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
/// absolute (see `write_rgce_loc_rel`'s doc comment for why). `pub(crate)`
/// so `wb_part.rs` can reuse it for `BrtName`'s `PtgArea3d` encoding, which
/// uses the exact same column field shape.
pub(crate) fn col_rel_short(col: u32) -> u16 {
    (col as u16) & 0x3FFF
}

/// `PtgArea3d`, reference class ([MS-XLS]/[MS-XLSB] Ptg grammar) — a
/// sheet-qualified rectangular range reference, confirmed byte-for-byte
/// against real Excel-authored `BrtName` records (see `wb_part.rs`'s module
/// doc comment) AND a real Excel-authored cross-sheet cell-formula
/// `SUM(Sheet2!B1:B5)` (see this module's doc comment) — both encode the
/// identical 15-byte token.
pub(crate) const PTG_AREA_3D: u8 = 0x3B;

/// Encode a complete `PtgArea3d` token (opcode + `ixti` + sheet-qualified
/// range) into `out`. Shared by `Formula::Range3d`'s own encoder and
/// `wb_part::encode_name` (a workbook `BrtName`'s range field is ALSO
/// always a `PtgArea3d` token stream, `cce`=15) — one encoder for both call
/// sites instead of two copies that could silently drift apart. `ixti` must
/// already be resolved (via `SheetRegistry` for a formula, or
/// `SheetRegistry::resolve_by_index` for a defined name) — this function
/// itself has no workbook context.
pub(crate) fn encode_ptg_area_3d(
    ixti: u16,
    row_first: u32,
    col_first: u32,
    row_last: u32,
    col_last: u32,
    out: &mut Vec<u8>,
) {
    out.push(PTG_AREA_3D);
    out.extend_from_slice(&ixti.to_le_bytes());
    out.extend_from_slice(&row_first.to_le_bytes());
    out.extend_from_slice(&row_last.to_le_bytes());
    out.extend_from_slice(&col_rel_short(col_first).to_le_bytes());
    out.extend_from_slice(&col_rel_short(col_last).to_le_bytes());
}

/// Workbook-wide, incrementally-growing registry backing cross-sheet
/// references — [MS-XLSB]'s `BrtExternSheet` (rid 362) XTI table, shared
/// (`Rc<RefCell<_>>`-wrapped, one per workbook — see `Worksheet`'s/
/// `Workbook`'s `sheet_registry` field in `lib.rs`) between:
/// - `Formula::sheet_cell`/`sheet_range` (`Ref3d`/`Range3d`), whose
///   `PtgRef3d`/`PtgArea3d` bytes are committed to a sheet's output the
///   moment that cell's row is flushed — for a streaming/sized-streaming
///   sheet, that can be well before a *later* sheet even exists as a
///   `Worksheet` object — so `ixti` has to be resolved incrementally, at
///   encode time, not deferred to workbook-finish time the way `BrtName`'s
///   `PtgArea3d` is (see below); and
/// - `Workbook::define_name`/`StreamingWorkbook::define_name`, whose
///   `BrtName` record is built once, at the very end, by
///   `wb_part::build_workbook` — using this SAME table so a formula cell's
///   already-written bytes and the final `BrtExternSheet` record this
///   registry backs always agree on the same `ixti` numbering, regardless
///   of which of the two registered a given sheet first.
///
/// Entry 0 is always `itab = 0`, unconditionally — required by the
/// pre-existing leftover `BrtName` record baked into `wb_part::WB_SUFFIX`
/// (hardcodes `ixti = 0`; see that module's doc comment) — so a workbook
/// using neither `define_name` nor a cross-sheet formula still produces
/// byte-identical `BrtExternSheet` output to before this feature existed
/// (confirmed: real Excel's own XTI table, with no defined names and no
/// pre-existing leftover record to accommodate, has NO such reserved first
/// entry — entries there are ordered purely by first-reference order,
/// confirmed against the same COM-automation reference file — so this
/// reserved slot is specifically a compatibility artifact of this crate's
/// own vendored `WB_SUFFIX` template, not a general BIFF12/XLSB rule).
pub(crate) struct SheetRegistry {
    /// Every sheet's name, in creation order — position IS that sheet's
    /// 0-based sheet index (matches the `sheet_names` slice passed to
    /// `wb_part::build_workbook`). Appended the moment a sheet is created
    /// (`Workbook::add_worksheet`/`StreamingWorkbook::new_worksheet`/
    /// `new_worksheet_sized`) — `resolve_by_name` can only find a sheet
    /// already in this list (see its own doc comment: no forward
    /// references to a not-yet-created sheet).
    names: Vec<String>,
    /// `BrtExternSheet`'s XTI dedup table: `itabs[ixti] == sheet_index`.
    itabs: Vec<u32>,
}

impl SheetRegistry {
    pub(crate) fn new() -> Self {
        Self {
            names: Vec::new(),
            itabs: vec![0],
        }
    }

    /// Record a newly-created sheet's name at its (already-known) index —
    /// called once, immediately, by `add_worksheet`/`new_worksheet`/
    /// `new_worksheet_sized`, so any *later* sheet's cross-sheet formula
    /// can resolve this name via `resolve_by_name`.
    pub(crate) fn register_sheet(&mut self, name: &str) {
        self.names.push(name.to_owned());
    }

    /// Resolve an already-known 0-based sheet index (used by
    /// `Workbook::define_name`, which takes an index directly, not a name)
    /// to its `ixti`, registering a new `BrtExternSheet` entry the first
    /// time this sheet index is referenced by ANYTHING — a formula or a
    /// defined name, whichever happens first.
    pub(crate) fn resolve_by_index(&mut self, sheet_index: u32) -> u16 {
        self.itabs.iter().position(|&t| t == sheet_index).unwrap_or_else(|| {
            self.itabs.push(sheet_index);
            self.itabs.len() - 1
        }) as u16
    }

    /// Resolve a sheet NAME (used by `Formula::sheet_cell`/`sheet_range` —
    /// a formula only knows the target sheet's name, not its index) to its
    /// `ixti`.
    ///
    /// # Panics
    /// If `name` doesn't match any sheet created so far — see this type's
    /// doc comment: a cross-sheet formula's bytes are committed to output
    /// the moment its row is flushed, which can happen before a sheet
    /// created *later* in the same program even exists, so forward
    /// references aren't supported. Add the target sheet (`add_worksheet`/
    /// `new_worksheet`/`new_worksheet_sized`) before writing a formula that
    /// references it.
    fn resolve_by_name(&mut self, name: &str) -> u16 {
        let idx = self.names.iter().position(|n| n == name).unwrap_or_else(|| {
            panic!(
                "xlsb_write: formula references sheet \"{name}\", which doesn't exist (yet) in \
                 this workbook — add_worksheet/new_worksheet/new_worksheet_sized must be called \
                 for the target sheet BEFORE writing a formula (Formula::sheet_cell/sheet_range) \
                 that references it; forward references to a not-yet-created sheet aren't \
                 supported."
            )
        });
        self.resolve_by_index(idx as u32)
    }

    /// The final `BrtExternSheet` payload (`cXti` + one `iSupBook`(0)/
    /// `itabFirst`/`itabLast` entry per registered itab) — built once, by
    /// `wb_part::build_workbook`, after every formula in the workbook has
    /// already been encoded (and thus every entry it could possibly need
    /// already registered) and every defined name has been resolved via
    /// `resolve_by_index`.
    pub(crate) fn build_externsheet_payload(&self) -> Vec<u8> {
        let mut pay = Vec::with_capacity(4 + self.itabs.len() * 12);
        pay.extend_from_slice(&(self.itabs.len() as u32).to_le_bytes()); // cXti
        for &itab in &self.itabs {
            pay.extend_from_slice(&0u32.to_le_bytes()); // iSupBook = 0 (this workbook)
            pay.extend_from_slice(&itab.to_le_bytes()); // itabFirst
            pay.extend_from_slice(&itab.to_le_bytes()); // itabLast
        }
        pay
    }
}

impl Default for SheetRegistry {
    fn default() -> Self {
        Self::new()
    }
}

/// The cell-record `grbitFlags` value for a volatile formula (`TODAY`/
/// `NOW`), confirmed byte-for-byte against real Excel's own `BrtFmlaNum`
/// output for `=TODAY()`/`=NOW()` (both had `grbitFlags=0x0002`; a
/// structurally-identical non-volatile formula, `=DATE(...)`, had
/// `grbitFlags=0x0000` — isolating this bit as the volatility marker).
/// Without it, Excel's dependency-based recalculation engine has no
/// reason to ever re-evaluate a formula with no cell dependencies, so it
/// would keep displaying the supplied `cached_value` forever instead of
/// updating on open/recalc.
const GRBIT_FLAGS_VOLATILE: u16 = 0x0002;

fn fmla_grbit_flags(formula: &Formula) -> u16 {
    if formula.is_volatile() { GRBIT_FLAGS_VOLATILE } else { 0 }
}

/// Write a `BrtFmlaNum` record: a formula cell whose most recent evaluation
/// produced a numeric value. `cached_value` is what viewers (including
/// calamine, and Excel before its own recalculation) display immediately.
/// `xti` resolves any `Formula::sheet_cell`/`sheet_range` node in `formula`
/// to its `ixti` — see `SheetRegistry`'s doc comment.
pub(crate) fn write_fmla_num(
    col: u32,
    ixfe: u16,
    cached_value: f64,
    formula: &Formula,
    xti: &mut SheetRegistry,
    buf: &mut Vec<u8>,
) {
    let rgce = formula.encode_with(xti);
    let mut pay = Vec::with_capacity(26 + rgce.len());
    pay.extend_from_slice(&col.to_le_bytes());
    pay.extend_from_slice(&(ixfe as u32).to_le_bytes()); // iStyleRef (low 24 bits) | fPhShow=0 | reserved=0
    pay.extend_from_slice(&cached_value.to_le_bytes());
    pay.extend_from_slice(&fmla_grbit_flags(formula).to_le_bytes());
    pay.extend_from_slice(&(rgce.len() as u32).to_le_bytes()); // cce
    pay.extend_from_slice(&rgce);
    pay.extend_from_slice(&0u32.to_le_bytes()); // cb (rgcb length) = 0
    write_rec(crate::biff12::RID_FMLA_NUM, &pay, buf);
}

/// Write a `BrtFmlaString` record: a formula cell whose most recent
/// evaluation produced a string value. See `write_fmla_num`'s doc comment
/// for `xti`.
pub(crate) fn write_fmla_string(
    col: u32,
    ixfe: u16,
    cached_value: &str,
    formula: &Formula,
    xti: &mut SheetRegistry,
    buf: &mut Vec<u8>,
) {
    let rgce = formula.encode_with(xti);
    let mut pay = Vec::with_capacity(10 + rgce.len());
    pay.extend_from_slice(&col.to_le_bytes());
    pay.extend_from_slice(&(ixfe as u32).to_le_bytes());
    write_wstr(cached_value, &mut pay); // string cached value: XLWideString (cch:4 + utf16)
    pay.extend_from_slice(&fmla_grbit_flags(formula).to_le_bytes());
    pay.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
    pay.extend_from_slice(&rgce);
    pay.extend_from_slice(&0u32.to_le_bytes());
    write_rec(crate::biff12::RID_FMLA_STRING, &pay, buf);
}

/// Write a `BrtFmlaBool` record: a formula cell whose most recent
/// evaluation produced a boolean value (`AND`/`OR`/`NOT`, or a bare
/// comparison like `A1=B1` used as a whole formula). Payload shape
/// (`col`(4) + `ixfe`(4) + `fBool`(1 byte, not the 8-byte `f64` field
/// `BrtFmlaNum` uses) + `grbitFlags`(2) + `cce`(4) + `rgce` + `cb`(4))
/// confirmed byte-for-byte against real Excel's own encoding of
/// `=AND(A1>0,A2>0)`.
pub(crate) fn write_fmla_bool(
    col: u32,
    ixfe: u16,
    cached_value: bool,
    formula: &Formula,
    xti: &mut SheetRegistry,
    buf: &mut Vec<u8>,
) {
    let rgce = formula.encode_with(xti);
    let mut pay = Vec::with_capacity(19 + rgce.len());
    pay.extend_from_slice(&col.to_le_bytes());
    pay.extend_from_slice(&(ixfe as u32).to_le_bytes());
    pay.push(cached_value as u8);
    pay.extend_from_slice(&fmla_grbit_flags(formula).to_le_bytes());
    pay.extend_from_slice(&(rgce.len() as u32).to_le_bytes());
    pay.extend_from_slice(&rgce);
    pay.extend_from_slice(&0u32.to_le_bytes());
    write_rec(crate::biff12::RID_FMLA_BOOL, &pay, buf);
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
        assert_eq!(
            &rgce[3..13],
            "hello"
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<u8>>()
                .as_slice()
        );
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
        write_fmla_num(
            1,
            0,
            55.0,
            &Formula::sum_range(1, 0, 10, 1),
            &mut SheetRegistry::new(),
            &mut buf,
        );
        let recs = parse_records(&buf);
        assert_eq!(recs.len(), 1);
        assert_eq!(recs[0].0, crate::biff12::RID_FMLA_NUM);
    }

    // ── 2026-09-13 formula-coverage expansion ──────────────────────────
    //
    // Every byte-shape asserted below (opcode, cparams-vs-no-cparams,
    // PtgAttrSemi wrapper) was cross-checked against a real Excel-authored
    // `.xlsb` built via COM automation — see the doc comments on
    // `FnIndex`, `Formula::Func`/`Formula::FuncFixed`, and
    // `GRBIT_FLAGS_VOLATILE` above for exactly what was compared.

    /// `&` must encode as the binary infix `PtgConcat` (0x08) — the same
    /// postfix shape as `Add`/`Sub`/`Mul`/`Div`, not a `PtgFuncVar` call.
    #[test]
    fn concat_operator_encodes_as_ptgconcat() {
        let f = Formula::cell(0, 0).concat(Formula::cell(0, 1)); // A1&B1
        let rgce = f.encode();
        assert_eq!(rgce.len(), 7 + 7 + 1); // PtgRef + PtgRef + PtgConcat
        assert_eq!(*rgce.last().unwrap(), 0x08); // PtgConcat last (outermost op)
    }

    /// `CONCATENATE` is variable-argument in real Excel: `PtgFuncVar` with
    /// `cparams` = the actual argument count and `tab` = 336 (0x0150).
    #[test]
    fn concatenate_encodes_as_funcvar_with_verified_index() {
        let f = Formula::concatenate(vec![Formula::str("a"), Formula::str("b")]);
        let rgce = f.encode();
        let tail = &rgce[rgce.len() - 4..];
        assert_eq!(tail[0], 0x42); // PtgFuncVar
        assert_eq!(tail[1], 2); // cparams
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 0x0150);
    }

    /// `AND`/`OR` are variable-argument: `PtgFuncVar`, tab = 36/37.
    #[test]
    fn and_or_encode_as_funcvar_with_verified_indices() {
        let and = Formula::and(vec![
            Formula::cell(0, 0).gt(Formula::num(0.0)),
            Formula::cell(1, 0).gt(Formula::num(0.0)),
        ]);
        let rgce = and.encode();
        let tail = &rgce[rgce.len() - 4..];
        assert_eq!(tail[0], 0x42);
        assert_eq!(tail[1], 2);
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 36);

        let or = Formula::or(vec![
            Formula::cell(0, 0).gt(Formula::num(0.0)),
            Formula::cell(1, 0).gt(Formula::num(0.0)),
        ]);
        let rgce = or.encode();
        let tail = &rgce[rgce.len() - 4..];
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 37);
    }

    /// `NOT` is fixed-argument in real Excel: `PtgFunc` (3 bytes, no
    /// `cparams` byte at all), tab = 38 — NOT `PtgFuncVar`.
    #[test]
    fn not_encodes_as_fixed_arg_ptgfunc() {
        let f = Formula::not(Formula::cell(0, 0).gt(Formula::num(0.0)));
        let rgce = f.encode();
        let tail = &rgce[rgce.len() - 3..];
        assert_eq!(tail[0], 0x41); // PtgFunc, not PtgFuncVar
        assert_eq!(u16::from_le_bytes(tail[1..3].try_into().unwrap()), 38);
    }

    /// `SUMIF` is variable-argument (2-3 args) — `PtgFuncVar`, tab = 345 —
    /// and its criteria argument is a plain `PtgStr`, not a special token.
    #[test]
    fn sumif_encodes_as_funcvar_with_string_criteria() {
        let f = Formula::sumif(Formula::range(0, 0, 9, 0), Formula::str(">3"));
        let rgce = f.encode();
        // PtgArea(13) + PtgStr(1+2+2*2=7) + PtgFuncVar(4) = 24
        assert_eq!(rgce.len(), 13 + 7 + 4);
        assert_eq!(rgce[13], 0x17); // PtgStr
        let tail = &rgce[rgce.len() - 4..];
        assert_eq!(tail[0], 0x42);
        assert_eq!(tail[1], 2);
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 345);
    }

    /// The 3-arg `SUMIF(range, criteria, sum_range)` form — reached via
    /// `Formula::Func(FnIndex::SUMIF, vec![range, criteria, sum_range])`
    /// directly, since `Formula::sumif` only covers the 2-arg form —
    /// separately confirmed byte-for-byte against real Excel's own
    /// `=SUMIF(A1:A5,"Hardware",B1:B5)`: still `PtgFuncVar`/tab=345, just
    /// with `cparams=3` and a 3rd `PtgArea` operand.
    #[test]
    fn sumif_3arg_form_encodes_with_cparams_3() {
        let f = Formula::Func(
            FnIndex::SUMIF,
            vec![
                Formula::range(0, 0, 4, 0),
                Formula::str("Hardware"),
                Formula::range(0, 1, 4, 1),
            ],
        );
        let rgce = f.encode();
        // PtgArea(13) + PtgStr(1+2+2*8=19) + PtgArea(13) + PtgFuncVar(4) = 49
        assert_eq!(rgce.len(), 13 + 19 + 13 + 4);
        let tail = &rgce[rgce.len() - 4..];
        assert_eq!(tail[0], 0x42);
        assert_eq!(tail[1], 3);
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 345);
    }

    /// `COUNTIF` is fixed-argument (always 2) — `PtgFunc`, tab = 346 — even
    /// though it takes the same shape of arguments as `SUMIF`, which is
    /// variable. This is exactly the "argument count alone doesn't tell
    /// you which Ptg family real Excel uses" trap this crate hit before.
    #[test]
    fn countif_encodes_as_fixed_arg_ptgfunc_unlike_sumif() {
        let f = Formula::countif(Formula::range(0, 0, 9, 0), Formula::str(">3"));
        let rgce = f.encode();
        // PtgArea(13) + PtgStr(7) + PtgFunc(3, no cparams) = 23
        assert_eq!(rgce.len(), 13 + 7 + 3);
        let tail = &rgce[rgce.len() - 3..];
        assert_eq!(tail[0], 0x41); // PtgFunc
        assert_eq!(u16::from_le_bytes(tail[1..3].try_into().unwrap()), 346);
    }

    /// `LEFT`/`RIGHT` are variable-argument (1-2 args) — `PtgFuncVar`.
    /// `MID`/`LEN` are fixed-argument — `PtgFunc`.
    #[test]
    fn text_functions_use_the_verified_ptg_family_per_function() {
        let left = Formula::left(Formula::cell(0, 1), Formula::num(3.0)).encode();
        let tail = &left[left.len() - 4..];
        assert_eq!(tail[0], 0x42);
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 115);

        let right = Formula::right(Formula::cell(0, 1), Formula::num(3.0)).encode();
        let tail = &right[right.len() - 4..];
        assert_eq!(tail[0], 0x42);
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 116);

        let mid = Formula::mid(Formula::cell(0, 1), Formula::num(2.0), Formula::num(3.0)).encode();
        let tail = &mid[mid.len() - 3..];
        assert_eq!(tail[0], 0x41);
        assert_eq!(u16::from_le_bytes(tail[1..3].try_into().unwrap()), 31);

        let len = Formula::len(Formula::cell(0, 1)).encode();
        let tail = &len[len.len() - 3..];
        assert_eq!(tail[0], 0x41);
        assert_eq!(u16::from_le_bytes(tail[1..3].try_into().unwrap()), 32);
    }

    /// `TEXT(value, format)` is fixed-argument — `PtgFunc`, tab = 48 — and
    /// its format-code argument is a plain `PtgStr`.
    #[test]
    fn text_fn_encodes_format_code_as_plain_string() {
        let f = Formula::text(Formula::cell(0, 0), Formula::str("0.00"));
        let rgce = f.encode();
        // PtgRef(7) + PtgStr(1+2+2*4=11) + PtgFunc(3) = 21
        assert_eq!(rgce.len(), 7 + 11 + 3);
        let tail = &rgce[rgce.len() - 3..];
        assert_eq!(tail[0], 0x41);
        assert_eq!(u16::from_le_bytes(tail[1..3].try_into().unwrap()), 48);
    }

    /// `DATE(y,m,d)` is fixed-argument, non-volatile — `PtgFunc`, tab = 65,
    /// and (unlike `TODAY`/`NOW`) NOT wrapped in `PtgAttrSemi`.
    #[test]
    fn date_encodes_as_fixed_arg_and_is_not_volatile() {
        let f = Formula::date(Formula::num(2026.0), Formula::num(9.0), Formula::num(13.0));
        assert!(!f.is_volatile());
        let rgce = f.encode();
        assert_eq!(rgce[0], 0x1F); // starts with a plain PtgNum, no PtgAttrSemi wrapper
        let tail = &rgce[rgce.len() - 3..];
        assert_eq!(tail[0], 0x41);
        assert_eq!(u16::from_le_bytes(tail[1..3].try_into().unwrap()), 65);
    }

    /// `TODAY`/`NOW` are volatile: wrapped in a leading `PtgAttrSemi`
    /// (flags=bitSemi=0x01) and `Formula::is_volatile()` reports `true` —
    /// both confirmed against real Excel's own `=TODAY()`/`=NOW()` output.
    #[test]
    fn today_and_now_are_volatile_and_wrapped_in_attr_semi() {
        let today = Formula::today();
        assert!(today.is_volatile());
        let rgce = today.encode();
        // PtgAttrSemi(4) + PtgFunc(3) = 7
        assert_eq!(rgce.len(), 7);
        assert_eq!(rgce[0], 0x19); // PtgAttrSemi
        assert_eq!(rgce[1], 0x01); // bitSemi
        assert_eq!(u16::from_le_bytes(rgce[2..4].try_into().unwrap()), 0); // offset=0
        assert_eq!(rgce[4], 0x41); // PtgFunc
        assert_eq!(u16::from_le_bytes(rgce[5..7].try_into().unwrap()), 221);

        let now = Formula::now();
        assert!(now.is_volatile());
        let rgce = now.encode();
        assert_eq!(rgce[0], 0x19);
        assert_eq!(u16::from_le_bytes(rgce[5..7].try_into().unwrap()), 74);
    }

    /// A formula that merely *contains* a volatile call (nested inside
    /// arithmetic) must still be reported volatile, so the cell-record
    /// `grbitFlags` bit gets set correctly.
    #[test]
    fn is_volatile_propagates_through_nesting() {
        let f = Formula::today().add(Formula::num(1.0));
        assert!(f.is_volatile());
        assert!(!Formula::num(1.0).add(Formula::num(2.0)).is_volatile());
    }

    /// `write_fmla_num`/`write_fmla_string` must set the cell record's
    /// volatile-recalculation bit (`grbitFlags = 0x0002`) for a formula
    /// containing `TODAY`/`NOW`, and leave it `0` otherwise — confirmed
    /// against real Excel's own `BrtFmlaNum` output (see
    /// `GRBIT_FLAGS_VOLATILE`'s doc comment).
    #[test]
    fn fmla_num_sets_volatile_grbit_flags_only_when_needed() {
        let mut buf = Vec::new();
        write_fmla_num(0, 0, 46000.0, &Formula::today(), &mut SheetRegistry::new(), &mut buf);
        let recs = parse_records(&buf);
        let payload = &recs[0].1;
        // col(4) + ixfe(4) + f64(8) = 16, grbitFlags at [16..18]
        assert_eq!(u16::from_le_bytes(payload[16..18].try_into().unwrap()), 0x0002);

        let mut buf2 = Vec::new();
        write_fmla_num(
            0,
            0,
            46000.0,
            &Formula::date(Formula::num(2026.0), Formula::num(9.0), Formula::num(13.0)),
            &mut SheetRegistry::new(),
            &mut buf2,
        );
        let recs2 = parse_records(&buf2);
        let payload2 = &recs2[0].1;
        assert_eq!(u16::from_le_bytes(payload2[16..18].try_into().unwrap()), 0x0000);
    }

    /// `PtgBool` (a boolean literal) is a single extra byte after the
    /// opcode — confirmed against real Excel's encoding of `VLOOKUP`'s
    /// `FALSE` 4th argument.
    #[test]
    fn bool_literal_encodes_as_ptgbool() {
        assert_eq!(Formula::boolean(true).encode(), vec![0x1D, 0x01]);
        assert_eq!(Formula::boolean(false).encode(), vec![0x1D, 0x00]);
    }

    /// `VLOOKUP` is variable-argument (3-4 args) — `PtgFuncVar`, tab = 102
    /// — and its exact-match flag is a `PtgBool`, not a `PtgInt`/`PtgNum`.
    #[test]
    fn vlookup_encodes_bool_arg_and_verified_index() {
        let f = Formula::vlookup(Formula::num(3.0), Formula::range(0, 0, 9, 1), Formula::num(2.0), false);
        let rgce = f.encode();
        let tail = &rgce[rgce.len() - 4..];
        assert_eq!(tail[0], 0x42);
        assert_eq!(tail[1], 4); // cparams
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 102);
        // The PtgBool(FALSE) sits right before the PtgFuncVar tail.
        assert_eq!(rgce[rgce.len() - 6], 0x1D);
        assert_eq!(rgce[rgce.len() - 5], 0x00);
    }

    /// `INDEX`/`MATCH` are variable-argument — `PtgFuncVar`, tab = 29/64.
    #[test]
    fn index_match_encode_with_verified_indices() {
        let index = Formula::index(Formula::range(0, 0, 9, 1), Formula::num(1.0), Formula::num(2.0)).encode();
        let tail = &index[index.len() - 4..];
        assert_eq!(tail[0], 0x42);
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 29);

        let match_ = Formula::match_(Formula::num(3.0), Formula::range(0, 0, 9, 0), Formula::num(0.0)).encode();
        let tail = &match_[match_.len() - 4..];
        assert_eq!(tail[0], 0x42);
        assert_eq!(u16::from_le_bytes(tail[2..4].try_into().unwrap()), 64);
    }

    #[test]
    fn fmla_bool_record_is_well_formed() {
        let mut buf = Vec::new();
        write_fmla_bool(
            0,
            0,
            true,
            &Formula::and(vec![Formula::cell(0, 0).gt(Formula::num(0.0))]),
            &mut SheetRegistry::new(),
            &mut buf,
        );
        let recs = parse_records(&buf);
        assert_eq!(recs.len(), 1);
        assert_eq!(recs[0].0, crate::biff12::RID_FMLA_BOOL);
    }
}
