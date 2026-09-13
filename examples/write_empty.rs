//! Dev tool: write a single empty sheet named "Sheet1" for direct
//! structural comparison against a real-Excel-produced empty .xlsb.
fn main() {
    let mut wb = xlsb_write::Workbook::new();
    wb.add_worksheet("Sheet1");
    wb.save("output_empty.xlsb").unwrap();
    println!("wrote output_empty.xlsb");
}
