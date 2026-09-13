//! Dev tool: parse the reference styles.bin blob record-by-record so we can
//! see the exact byte layout of BrtFont/BrtFill/BrtBorder/BrtXF records
//! empirically, instead of guessing from spec memory. Run with:
//!   cargo run --example dump_styles

fn main() {
    let bytes = xlsb_write::styles::default_styles_bin();
    let recs = xlsb_write::biff12::parse_records(&bytes);
    for (rid, payload) in &recs {
        let hex: String = payload.iter().map(|b| format!("{b:02x}")).collect::<Vec<_>>().join(" ");
        println!("rid=0x{rid:04x} ({rid:4}) len={:3}  {}", payload.len(), hex);
    }
}
