//! Shared String Table (SST) builder and encoder.
//!
//! The SST maps unique strings → u32 index (insertion order).
//! `indexmap::IndexMap` gives O(1) lookup and preserves insertion order,
//! which is required because the BrtSst record must list strings in index order.
//!
//! Adapted from the `xlsb-writer` crate (MIT License, Copyright (c) 2026 kotucha).

use crate::biff12::{write_r0, write_rec};
use indexmap::IndexMap;

pub struct Sst {
    map: IndexMap<String, u32>,
    /// Total number of `intern()` calls, including repeats — this is
    /// `cstTotal` (references), which is NOT the same as the unique count
    /// (`cstUnique`, `map.len()`). Getting these confused writes a wildly
    /// wrong `cstTotal` for any workbook with repeated strings (e.g. a
    /// "RETAIL" channel value repeated hundreds of thousands of times
    /// still only counts once toward `map.len()`), which is exactly the
    /// kind of internal-consistency mismatch a strict validator (real
    /// Excel, not just lenient readers like calamine/LibreOffice) flags.
    total_refs: u32,
}

impl Sst {
    pub fn new() -> Self {
        Self {
            map: IndexMap::new(),
            total_refs: 0,
        }
    }

    /// Insert a string and return its SST index (idempotent for the index;
    /// each call still counts toward `cstTotal`). Looks up by borrowed
    /// `&str` first and only allocates an owned `String` on a genuine miss
    /// — report data is overwhelmingly categorical (the same value repeated
    /// across hundreds of thousands of rows), so allocating unconditionally
    /// on every call (as `self.map.entry(s.to_owned())` would, since
    /// `entry` needs an owned key up front regardless of whether it's
    /// already present) meant millions of throwaway allocations on exactly
    /// the hot path this crate cares most about.
    pub fn intern(&mut self, s: &str) -> u32 {
        self.total_refs += 1;
        if let Some(&idx) = self.map.get(s) {
            return idx;
        }
        let idx = self.map.len() as u32;
        self.map.insert(s.to_owned(), idx);
        idx
    }

    /// Encode the complete sharedStrings.bin binary.
    ///
    /// Format:
    ///   BrtBeginSst(0x009F)  cstTotal(4) cstUnique(4)
    ///   for each string:
    ///     BrtSstItem(0x0013)  flags(1) cch(4) utf16le(cch*2)
    ///   BrtEndSst(0x00A0)
    pub fn encode(&self) -> Vec<u8> {
        let cst_unique = self.map.len() as u32;
        let mut buf = Vec::with_capacity(self.map.len() * 32);

        let hdr = {
            let mut h = [0u8; 8];
            h[0..4].copy_from_slice(&self.total_refs.to_le_bytes());
            h[4..8].copy_from_slice(&cst_unique.to_le_bytes());
            h
        };
        write_rec(0x009F, &hdr, &mut buf);

        for s in self.map.keys() {
            let utf16: Vec<u16> = s.encode_utf16().collect();
            let cch = utf16.len() as u32;
            let mut pay = Vec::with_capacity(1 + 4 + utf16.len() * 2);
            pay.push(0x00); // fHighByte = 0 (UTF-16LE, not compressed)
            pay.extend_from_slice(&cch.to_le_bytes());
            for ch in &utf16 {
                pay.extend_from_slice(&ch.to_le_bytes());
            }
            write_rec(0x0013, &pay, &mut buf);
        }

        write_r0(0x00A0, &mut buf); // BrtEndSst
        buf
    }
}

impl Default for Sst {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_intern_idempotent() {
        let mut sst = Sst::new();
        assert_eq!(sst.intern("hello"), 0);
        assert_eq!(sst.intern("world"), 1);
        assert_eq!(sst.intern("hello"), 0); // same index
        assert_eq!(sst.map.len(), 2);
    }

    /// Regression: cstTotal (references, including repeats) must NOT be
    /// confused with cstUnique (the SST's own size) — a real Excel
    /// validator flags this internal-consistency mismatch even though
    /// lenient readers ignore the field.
    #[test]
    fn cst_total_counts_repeats_cst_unique_does_not() {
        let mut sst = Sst::new();
        sst.intern("hello");
        sst.intern("hello");
        sst.intern("hello");
        sst.intern("world");
        let bin = sst.encode();
        let recs = crate::biff12::parse_records(&bin);
        let (rid, payload) = &recs[0];
        assert_eq!(*rid, 0x009F);
        let cst_total = u32::from_le_bytes(payload[0..4].try_into().unwrap());
        let cst_unique = u32::from_le_bytes(payload[4..8].try_into().unwrap());
        assert_eq!(cst_total, 4, "4 intern() calls were made, repeats included");
        assert_eq!(cst_unique, 2, "only 2 distinct strings");
    }

    #[test]
    fn test_encode_nonempty() {
        let mut sst = Sst::new();
        sst.intern("A");
        sst.intern("B");
        let bin = sst.encode();
        assert_eq!(bin[0], 0x9F);
        assert_eq!(bin[1], 0x01);
        assert!(!bin.is_empty());
    }
}
