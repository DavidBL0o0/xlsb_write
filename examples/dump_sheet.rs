//! Dev tool: parse a raw worksheet .bin (extracted from an .xlsb zip) and
//! print its BIFF12 records one per line, so two files can be diffed
//! record-by-record instead of by raw hex offset.
//! Run with: cargo run --example dump_sheet -- <path-to-sheetN.bin>

fn main() {
    let path = std::env::args().nth(1).expect("usage: dump_sheet <path>");
    let bytes = std::fs::read(&path).expect("read file");
    let recs = xlsb_write::biff12::parse_records(&bytes);
    for (rid, payload) in &recs {
        let hex: String = payload.iter().map(|b| format!("{b:02x}")).collect::<Vec<_>>().join(" ");
        println!("rid={rid:4}  len={:6}  {}", payload.len(), hex);
    }
    println!("-- total {} records, {} bytes --", recs.len(), bytes.len());
}
