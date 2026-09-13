//! Dev tool: print a parquet file's full column list.
//! Run with: cargo run --example dump_schema -- <path.parquet>
use parquet::file::reader::{FileReader, SerializedFileReader};

fn main() {
    let path = std::env::args().nth(1).expect("usage: dump_schema <path.parquet>");
    let file = std::fs::File::open(&path).expect("open file");
    let reader = SerializedFileReader::new(file).expect("open parquet reader");
    let schema = reader.metadata().file_metadata().schema_descr();
    for i in 0..schema.num_columns() {
        let col = schema.column(i);
        println!("{i:3}  {:<10}  {}", format!("{:?}", col.physical_type()), col.name());
    }
}
