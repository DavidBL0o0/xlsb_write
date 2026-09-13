//! BIFF12 encoding primitives — varints, records, cell types.
//!
//! All multi-byte integers are little-endian (LE) as per the MS-XLSB spec.
//!
//! Originally adapted from the `xlsb-writer` crate (MIT License,
//! Copyright (c) 2026 kotucha, <https://github.com/kotucha/xlsb-writer>),
//! byte-verified against real Excel output. Extended here with formula,
//! blank/error and additional record helpers as the library grows.

// ── Varint encoding ──────────────────────────────────────────────────────────

/// Encode a BIFF12 variable-length unsigned integer into `buf`.
/// Returns the number of bytes written (1–4 for the values we emit).
#[inline]
pub fn write_vi(val: u32, buf: &mut Vec<u8>) {
    let mut v = val;
    loop {
        let b = (v & 0x7F) as u8;
        v >>= 7;
        if v == 0 {
            buf.push(b);
            break;
        } else {
            buf.push(b | 0x80);
        }
    }
}

/// Decode a BIFF12 varint starting at `pos`. Returns (value, bytes consumed).
#[inline]
pub fn read_vi(data: &[u8], mut pos: usize) -> (u32, usize) {
    let mut val = 0u32;
    let mut shift = 0;
    let start = pos;
    loop {
        let b = data[pos];
        pos += 1;
        val |= ((b & 0x7F) as u32) << shift;
        shift += 7;
        if b & 0x80 == 0 {
            break;
        }
    }
    (val, pos - start)
}

/// Encode a BIFF12 record: varint(rid) + varint(payload.len()) + payload.
#[inline]
pub fn write_rec(rid: u32, payload: &[u8], buf: &mut Vec<u8>) {
    write_vi(rid, buf);
    write_vi(payload.len() as u32, buf);
    buf.extend_from_slice(payload);
}

/// Encode a zero-length record.
#[inline]
pub fn write_r0(rid: u32, buf: &mut Vec<u8>) {
    write_vi(rid, buf);
    buf.push(0x00);
}

/// Parse a stream of BIFF12 records. Used by tests and debugging tools.
/// Panics on malformed input — prefer `try_parse_records` for untrusted
/// or possibly-corrupt data (e.g. a validation tool).
pub fn parse_records(data: &[u8]) -> Vec<(u32, Vec<u8>)> {
    try_parse_records(data).expect("malformed BIFF12 record stream")
}

/// Like `parse_records`, but returns an error describing exactly where and
/// why parsing failed instead of panicking — a length that overruns the
/// buffer, a truncated varint at EOF, or a trailing partial record.
pub fn try_parse_records(data: &[u8]) -> Result<Vec<(u32, Vec<u8>)>, String> {
    let mut recs = Vec::new();
    let mut pos = 0;
    while pos < data.len() {
        let start = pos;
        let Some((rid, n)) = try_read_vi(data, pos) else {
            return Err(format!("truncated rid varint at byte {start} (stream len {})", data.len()));
        };
        pos += n;
        let Some((rlen, n)) = try_read_vi(data, pos) else {
            return Err(format!("truncated length varint at byte {pos} (record rid={rid} started at {start})"));
        };
        pos += n;
        let end = pos + rlen as usize;
        if end > data.len() {
            return Err(format!(
                "record rid={rid} at byte {start} declares length {rlen}, but only {} bytes remain",
                data.len() - pos
            ));
        }
        recs.push((rid, data[pos..end].to_vec()));
        pos = end;
    }
    Ok(recs)
}

fn try_read_vi(data: &[u8], mut pos: usize) -> Option<(u32, usize)> {
    let mut val = 0u32;
    let mut shift = 0;
    let start = pos;
    loop {
        let b = *data.get(pos)?;
        pos += 1;
        val |= ((b & 0x7F) as u32) << shift;
        shift += 7;
        if b & 0x80 == 0 {
            break;
        }
    }
    Some((val, pos - start))
}

// ── XLWideString ─────────────────────────────────────────────────────────────

/// Encode an XLWideString: cch (u32 LE) + UTF-16LE bytes.
pub fn write_wstr(s: &str, buf: &mut Vec<u8>) {
    let encoded: Vec<u16> = s.encode_utf16().collect();
    buf.extend_from_slice(&(encoded.len() as u32).to_le_bytes());
    for ch in &encoded {
        buf.extend_from_slice(&ch.to_le_bytes());
    }
}

// ── Per-instance unique identifiers ─────────────────────────────────────────

/// Generates 16 bytes that are, with overwhelming probability, unique
/// across calls within a process — needed for the handful of `[MS-XLSB]`
/// fields real Excel stamps freshly on every save (e.g. the FRT-wrapped
/// per-sheet identifier in `sheet::write_sheet_footer`) that this crate
/// previously baked in as one static value shared by every sheet of every
/// file it ever wrote (confirmed a real bug 2026-09-13 by diffing this
/// crate's output against two independently-produced real files — see
/// `docs/superpowers/specs/2026-09-13-byte-block-hardening-design.md`).
///
/// Not cryptographically random and not a spec-conformant GUID: these
/// fields aren't security-sensitive or required to be well-formed UUIDs,
/// they just need to not collide within a single run, so hashing wall-clock
/// time against a call counter is enough and needs no new dependency.
pub fn pseudo_unique_16_bytes() -> [u8; 16] {
    use std::hash::{Hash, Hasher};
    use std::sync::atomic::{AtomicU64, Ordering};
    use std::time::{SystemTime, UNIX_EPOCH};

    static COUNTER: AtomicU64 = AtomicU64::new(0);
    let count = COUNTER.fetch_add(1, Ordering::Relaxed);
    let nanos = SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_nanos()).unwrap_or(0);

    let mut out = [0u8; 16];
    for (i, salt) in [0xA5u8, 0x5Au8].into_iter().enumerate() {
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        (nanos, count, salt).hash(&mut hasher);
        out[i * 8..i * 8 + 8].copy_from_slice(&hasher.finish().to_le_bytes());
    }
    out
}

// ── RK number encoding ───────────────────────────────────────────────────────

/// Try to encode `v` as an RK (compact number). Returns `None` if the value
/// cannot be represented (e.g. NaN, Inf, or out of RK range).
pub fn encode_rk(v: f64) -> Option<u32> {
    if !v.is_finite() {
        return None;
    }
    // Integer form: fInt=1 (bit1), fX100=0 (bit0)
    let iv = v as i64;
    if iv as f64 == v && iv >= -(1 << 29) && iv < (1 << 29) {
        return Some(((iv << 2) | 2) as u32);
    }
    // ×100 integer form: fInt=1, fX100=1
    let v100 = v * 100.0;
    let iv100 = v100.round() as i64;
    if (v100 - iv100 as f64).abs() < 1e-6 && iv100 >= -(1 << 29) && iv100 < (1 << 29) {
        return Some(((iv100 << 2) | 3) as u32);
    }
    // Double-top form: fInt=0, fX100=0. This form stores only the high 32
    // bits of the f64 and drops the low 32 — only lossless when those low
    // bits are actually zero (e.g. 0.5, 0.25, 2.0). For anything else this
    // must NOT be used, or the value silently loses precision on save.
    let bits = v.to_bits();
    let hi = (bits >> 32) as u32;
    let lo = bits as u32;
    // lower 2 bits of hi word must be 0 (they become the fInt/fX100 flags)
    if lo == 0 && hi & 3 == 0 {
        return Some(hi);
    }
    None
}

// ── Cell record builders ──────────────────────────────────────────────────────
//
// Each cell record layout:
//   col(4) ixfe(2) grbitFmt(2) [value...]
//
// Note: grbitFmt is always 0x0000 for data cells (not 0x0100 like blank cells).
// BrtCellBlank uses 0x0100 for grbitFmt to match reference files.

pub const RID_CELL_BLANK: u32 = 0x0001;
pub const RID_CELL_RK: u32 = 0x0002;
pub const RID_CELL_ERROR: u32 = 0x0003;
pub const RID_CELL_BOOL: u32 = 0x0004;
pub const RID_CELL_REAL: u32 = 0x0005;
pub const RID_CELL_ISST: u32 = 0x0007;
// Verified against pyxlsb's recordtypes.py (an independent, authoritative
// reader implementation) — the original crate's own numbering for these was
// never actually used/tested and turned out to be off by 2.
pub const RID_FMLA_STRING: u32 = 8;
pub const RID_FMLA_NUM: u32 = 9;
pub const RID_FMLA_BOOL: u32 = 10;
pub const RID_FMLA_ERROR: u32 = 11;

/// Write a BrtCellBlank record (col, ixfe, grbitFmt=0x0100).
pub fn write_cell_blank(col: u32, ixfe: u16, buf: &mut Vec<u8>) {
    let mut pay = [0u8; 8];
    pay[0..4].copy_from_slice(&col.to_le_bytes());
    pay[4..6].copy_from_slice(&ixfe.to_le_bytes());
    pay[6..8].copy_from_slice(&0x0100u16.to_le_bytes());
    write_rec(RID_CELL_BLANK, &pay, buf);
}

/// Write a BrtCellRk record.
pub fn write_cell_rk(col: u32, ixfe: u16, rk: u32, buf: &mut Vec<u8>) {
    let mut pay = [0u8; 12];
    pay[0..4].copy_from_slice(&col.to_le_bytes());
    pay[4..6].copy_from_slice(&ixfe.to_le_bytes());
    // pay[6..8] = grbitFmt = 0x0000
    pay[8..12].copy_from_slice(&rk.to_le_bytes());
    write_rec(RID_CELL_RK, &pay, buf);
}

/// Write a BrtCellReal record (full f64).
pub fn write_cell_real(col: u32, ixfe: u16, v: f64, buf: &mut Vec<u8>) {
    let mut pay = [0u8; 16];
    pay[0..4].copy_from_slice(&col.to_le_bytes());
    pay[4..6].copy_from_slice(&ixfe.to_le_bytes());
    pay[8..16].copy_from_slice(&v.to_le_bytes());
    write_rec(RID_CELL_REAL, &pay, buf);
}

/// Write a BrtCellBool record.
pub fn write_cell_bool(col: u32, ixfe: u16, v: bool, buf: &mut Vec<u8>) {
    let mut pay = [0u8; 9];
    pay[0..4].copy_from_slice(&col.to_le_bytes());
    pay[4..6].copy_from_slice(&ixfe.to_le_bytes());
    pay[8] = v as u8;
    write_rec(RID_CELL_BOOL, &pay, buf);
}

/// Write a BrtCellIsst record (shared string index).
pub fn write_cell_isst(col: u32, ixfe: u16, isst: u32, buf: &mut Vec<u8>) {
    let mut pay = [0u8; 12];
    pay[0..4].copy_from_slice(&col.to_le_bytes());
    pay[4..6].copy_from_slice(&ixfe.to_le_bytes());
    pay[8..12].copy_from_slice(&isst.to_le_bytes());
    write_rec(RID_CELL_ISST, &pay, buf);
}

/// Write a BrtCellError record (col, ixfe, error byte code).
pub fn write_cell_error(col: u32, ixfe: u16, err_code: u8, buf: &mut Vec<u8>) {
    let mut pay = [0u8; 9];
    pay[0..4].copy_from_slice(&col.to_le_bytes());
    pay[4..6].copy_from_slice(&ixfe.to_le_bytes());
    pay[8] = err_code;
    write_rec(RID_CELL_ERROR, &pay, buf);
}

// ── Row header ────────────────────────────────────────────────────────────────

/// Write a BrtRowHdr record (25-byte payload).
///
/// Layout (verified byte-for-byte against reference xlsb):
/// ```text
/// [0:4]   rw        u32  — row index
/// [4:8]   ixfe      u32  — XF index (ignored when fGhostDirty=0)
/// [8:10]  miyRw     u16  — row height in 1/20pt units
/// [10:12] flags     u16  — bit 5 (0x0020) = fCollapsed, bit 6 (0x0040) = fZeroHeight (hidden)
/// [12]    padding   u8   — 0x00
/// [13:17] ccolspan  u32  — number of BrtColSpan entries (1)
/// [17:21] colFirst  u32  — 0
/// [21:25] colLast   u32  — last col index with a cell record in this row
/// ```
pub fn write_row_hdr(row_idx: u32, last_col: u32, height_twips: u16, hidden: bool, buf: &mut Vec<u8>) {
    let mut pay = [0u8; 25];
    pay[0..4].copy_from_slice(&row_idx.to_le_bytes());
    pay[8..10].copy_from_slice(&height_twips.to_le_bytes());
    let flags: u16 = if hidden { 0x0040 } else { 0x0000 };
    pay[10..12].copy_from_slice(&flags.to_le_bytes());
    pay[13..17].copy_from_slice(&1u32.to_le_bytes());
    pay[21..25].copy_from_slice(&last_col.to_le_bytes());
    write_rec(0x0000, &pay, buf);
}

// ── Row envelope (appears before every row, from row 1 onwards) ───────────────
// Row 0's envelope is provided by the sheet header.

/// The fixed 15-byte row envelope written before each row (rows 1+).
pub const ROW_PRE: &[u8] = &[
    0x25, 0x06, 0x01, 0x00, 0x02, 0x0e, 0x00, 0x80, // BrtBeginList payload
    0x80, 0x08, 0x02, 0x05, 0x00, // 0x0400 record
    0x26, 0x00, // BrtEndList
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_vi() {
        let mut buf = Vec::new();
        write_vi(0, &mut buf);
        assert_eq!(buf, [0x00]);
        buf.clear();
        write_vi(127, &mut buf);
        assert_eq!(buf, [0x7F]);
        buf.clear();
        write_vi(128, &mut buf);
        assert_eq!(buf, [0x80, 0x01]);
        buf.clear();
        write_vi(0x009C, &mut buf);
        assert_eq!(buf, [0x9C, 0x01]); // BrtBundleSh
    }

    #[test]
    fn test_vi_roundtrip() {
        for v in [0u32, 1, 127, 128, 255, 0x3FFF, 0x1FFFFF, u32::MAX >> 4] {
            let mut buf = Vec::new();
            write_vi(v, &mut buf);
            let (decoded, _) = read_vi(&buf, 0);
            assert_eq!(decoded, v, "vi roundtrip failed for {v}");
        }
    }

    #[test]
    fn test_rk_integer() {
        assert_eq!(encode_rk(8129.0), Some(((8129i64 << 2) | 2) as u32));
        assert_eq!(encode_rk(0.0), Some(2));
        assert_eq!(encode_rk(-1.0), Some((((-1i64) << 2) | 2) as u32));
    }

    #[test]
    fn test_rk_x100() {
        // 9.99 → ×100 = 999 → RK integer form with fX100 set
        let rk = encode_rk(9.99).unwrap();
        assert_eq!(rk & 1, 1); // fX100 set
        assert_eq!(rk & 2, 2); // fInt set
        let iv = (rk >> 2) as i64;
        assert_eq!(iv, 999);
    }

    #[test]
    fn test_rk_nan_inf() {
        assert!(encode_rk(f64::NAN).is_none());
        assert!(encode_rk(f64::INFINITY).is_none());
    }

    /// Decode an RK value back to f64, mirroring the 3 encode_rk forms.
    fn decode_rk(rk: u32) -> f64 {
        let f_int = rk & 2 != 0;
        let f_x100 = rk & 1 != 0;
        if f_int {
            let iv = (rk as i32) >> 2;
            if f_x100 { iv as f64 / 100.0 } else { iv as f64 }
        } else {
            f64::from_bits(((rk & !3) as u64) << 32)
        }
    }

    /// Regression: the double-top RK form only stores the high 32 bits of
    /// the f64. A value like 0.1534 has non-zero low bits, so encode_rk
    /// must not silently return a truncated double-top encoding for it
    /// (~0.15339994430541992) — either fall back to None (BrtCellReal), or
    /// use the ×100 form, which is exact to the cent by construction.
    #[test]
    fn test_rk_rejects_lossy_double_top() {
        for v in [0.1534, 0.1, 1.0 / 3.0, 12345.678, -0.02, 9.99] {
            if let Some(rk) = encode_rk(v) {
                let decoded = decode_rk(rk);
                assert!((decoded - v).abs() < 1e-9, "RK encoding of {v} was lossy: decoded as {decoded}");
            }
        }
    }

    #[test]
    fn test_rk_double_top_accepts_exact_values() {
        // 0.5 == 2^-1: low 32 bits of its f64 representation are exactly 0.
        assert!(encode_rk(0.5).is_some());
    }
}
