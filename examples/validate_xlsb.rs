//! Structural validator for a .xlsb file: checks every binary part parses
//! as a well-formed BIFF12 record stream (catching the class of bug that
//! made Excel silently discard/"repair" a worksheet part — a declared
//! record length or dimension that doesn't match the actual content),
//! then opens it with calamine (a reader completely independent of this
//! crate) as a second, value-level check.
//!
//! This does NOT prove Excel will accept the file (only Excel's own
//! parser can prove that) — it proves the byte-level record framing is
//! self-consistent, which is the class of bug that has bitten this
//! project twice so far (a hardcoded BrtWsDim, a short static array).
//!
//! Run with: cargo run --example validate_xlsb -- <path.xlsb>

use calamine::{Reader, Xlsb, open_workbook};
use std::io::Read;
use xlsb_write::biff12::try_parse_records;
use zip::ZipArchive;

fn main() {
    let path = std::env::args().nth(1).expect("usage: validate_xlsb <path.xlsb>");
    let mut ok = true;

    println!("=== Structural check: every .bin part is a well-formed BIFF12 record stream ===");
    let file = std::fs::File::open(&path).expect("open file");
    let mut zip = ZipArchive::new(file).expect("open as zip");
    let mut names: Vec<String> = (0..zip.len())
        .map(|i| zip.by_index(i).unwrap().name().to_string())
        .collect();
    names.sort();

    for name in &names {
        if !name.ends_with(".bin") {
            continue;
        }
        let mut entry = zip.by_name(name).expect("reopen entry");
        let mut bytes = Vec::with_capacity(entry.size() as usize);
        entry.read_to_end(&mut bytes).expect("read entry");
        match try_parse_records(&bytes) {
            Ok(recs) => println!("  OK    {name}  ({} bytes, {} records)", bytes.len(), recs.len()),
            Err(e) => {
                println!("  FAIL  {name}  ({} bytes): {e}", bytes.len());
                ok = false;
            }
        }
    }

    println!("\n=== Content_Types / worksheet part cross-check ===");
    let mut ct = String::new();
    zip.by_name("[Content_Types].xml")
        .expect("Content_Types missing")
        .read_to_string(&mut ct)
        .unwrap();
    for name in &names {
        if name.starts_with("xl/worksheets/sheet") && name.ends_with(".bin") {
            let part_name = format!("/{name}");
            if ct.contains(&part_name) {
                println!("  OK    {name} declared in [Content_Types].xml");
            } else {
                println!("  FAIL  {name} NOT declared in [Content_Types].xml");
                ok = false;
            }
        }
    }

    println!("\n=== calamine (independent reader) value-level check ===");
    match open_workbook::<Xlsb<_>, _>(&path) {
        Ok(mut wb) => {
            for sheet_name in wb.sheet_names().to_owned() {
                match wb.worksheet_range(&sheet_name) {
                    Ok(range) => {
                        let (rows, cols) = range.get_size();
                        let non_empty = range.used_cells().count();
                        println!("  OK    sheet '{sheet_name}': {rows} x {cols}, {non_empty} non-empty cells");
                    }
                    Err(e) => {
                        println!("  FAIL  sheet '{sheet_name}': {e}");
                        ok = false;
                    }
                }
            }
        }
        Err(e) => {
            println!("  FAIL  could not open workbook: {e}");
            ok = false;
        }
    }

    println!("\n=== Result ===");
    if ok {
        println!("PASS — {path} is structurally well-formed.");
    } else {
        println!("FAIL — see errors above.");
        std::process::exit(1);
    }
}
