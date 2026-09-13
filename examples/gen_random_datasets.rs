//! Generates a fixed matrix of randomly-shaped/typed parquet datasets into
//! `test_fixtures/random/`, used by `tests/robustness.rs` to prove the writer
//! works correctly across a wide range of shapes, not just one hand-picked
//! layout. Not a fuzzer — a deliberately chosen set of tricky (row count,
//! column count, column type mix) combinations, generated with a fixed seed
//! so a failure is always reproducible by re-running this generator, never
//! a flake.
//!
//! Uses a real `polars::DataFrame` + `ParquetWriter` (not the lower-level
//! `parquet` crate already used elsewhere in this repo) so the fixtures are
//! genuinely representative of what a Polars user would hand this crate.
//!
//! Run with: cargo run --example gen_random_datasets

use polars::prelude::*;
use rand::rngs::StdRng;
use rand::{RngExt, SeedableRng};
use std::fs::File;
use std::path::Path;

const SEED: u64 = 1337;
const NULL_RATE: f64 = 0.12;

/// One deliberately-chosen (rows, cols) shape. `col_type_start` shifts which
/// column type each case starts on, so the "first column" varies (numeric,
/// then string, etc.) across the matrix instead of every case coincidentally
/// starting with the same type.
struct Case {
    rows: usize,
    cols: usize,
    col_type_start: usize,
}

const CASES: &[Case] = &[
    Case { rows: 0, cols: 1, col_type_start: 2 }, // empty, single string column
    Case { rows: 1, cols: 1, col_type_start: 0 }, // minimal, single numeric column
    Case { rows: 1, cols: 40, col_type_start: 1 }, // wide, one row
    Case { rows: 7, cols: 3, col_type_start: 0 },
    Case { rows: 7, cols: 40, col_type_start: 3 },
    Case { rows: 500, cols: 3, col_type_start: 2 },
    Case { rows: 500, cols: 12, col_type_start: 0 },
    Case { rows: 500, cols: 40, col_type_start: 4 },
    Case { rows: 20_000, cols: 3, col_type_start: 1 },
    Case { rows: 20_000, cols: 12, col_type_start: 0 },
];

#[derive(Clone, Copy)]
enum ColKind {
    Int,
    Float,
    Categorical,
    UniqueString,
    Bool,
}

impl ColKind {
    fn cycle(i: usize) -> Self {
        match i % 5 {
            0 => ColKind::Int,
            1 => ColKind::Float,
            2 => ColKind::Categorical,
            3 => ColKind::UniqueString,
            _ => ColKind::Bool,
        }
    }
}

const CATEGORIES: &[&str] = &["RETAIL", "ONLINE", "WHOLESALE", "OUTLET"];

fn random_unicode_word(rng: &mut StdRng) -> String {
    // Mixes plain ASCII with accented/CJK/emoji codepoints so string
    // handling is stressed with real non-ASCII text, not just ASCII.
    const POOL: &[char] = &['a', 'b', 'ñ', 'é', 'ü', '中', '文', '🙂', 'ç', 'z'];
    let len = rng.random_range(1..=8);
    (0..len).map(|_| POOL[rng.random_range(0..POOL.len())]).collect()
}

fn build_int_column(rng: &mut StdRng, rows: usize) -> Vec<Option<i64>> {
    (0..rows)
        .map(|i| {
            if rng.random_bool(NULL_RATE) {
                return None;
            }
            match i % 4 {
                0 => Some(rng.random_range(-1000..1000)),
                1 => Some(i64::MAX - rng.random_range(0..1000)),
                2 => Some(i64::MIN + rng.random_range(0..1000)),
                _ => Some(0),
            }
        })
        .collect()
}

fn build_float_column(rng: &mut StdRng, rows: usize) -> Vec<Option<f64>> {
    (0..rows)
        .map(|i| {
            if rng.random_bool(NULL_RATE) {
                return None;
            }
            match i % 5 {
                0 => Some(f64::NAN),
                1 => Some(rng.random_range(-1e9..1e9)),
                2 => Some(rng.random_range(-1.0..1.0) * 1e-9), // tiny magnitude
                3 => Some(-0.0),
                _ => Some(rng.random_range(0.0..1_000_000.0)),
            }
        })
        .collect()
}

fn build_categorical_column(rng: &mut StdRng, rows: usize) -> Vec<Option<String>> {
    (0..rows)
        .map(|_| {
            if rng.random_bool(NULL_RATE) {
                None
            } else {
                Some(CATEGORIES[rng.random_range(0..CATEGORIES.len())].to_owned())
            }
        })
        .collect()
}

fn build_unique_string_column(rng: &mut StdRng, rows: usize) -> Vec<Option<String>> {
    (0..rows)
        .map(|i| {
            if rng.random_bool(NULL_RATE) {
                return None;
            }
            match i % 4 {
                0 => Some(String::new()),                          // empty string
                1 => Some("x".repeat(500)),                        // long string
                2 => Some(random_unicode_word(rng)),                // unicode
                _ => Some(format!("row-{i}-{}", rng.random::<u32>())), // unique
            }
        })
        .collect()
}

fn build_bool_column(rng: &mut StdRng, rows: usize) -> Vec<Option<bool>> {
    (0..rows).map(|_| if rng.random_bool(NULL_RATE) { None } else { Some(rng.random_bool(0.5)) }).collect()
}

fn build_case_df(rng: &mut StdRng, case: &Case) -> DataFrame {
    let mut columns: Vec<Column> = Vec::with_capacity(case.cols);
    for c in 0..case.cols {
        let name = PlSmallStr::from_string(format!("col_{c}"));
        let kind = ColKind::cycle(case.col_type_start + c);
        let series = match kind {
            ColKind::Int => Series::new(name, build_int_column(rng, case.rows)),
            ColKind::Float => Series::new(name, build_float_column(rng, case.rows)),
            ColKind::Categorical => Series::new(name, build_categorical_column(rng, case.rows)),
            ColKind::UniqueString => Series::new(name, build_unique_string_column(rng, case.rows)),
            ColKind::Bool => Series::new(name, build_bool_column(rng, case.rows)),
        };
        columns.push(series.into());
    }
    DataFrame::new(case.rows, columns).expect("build DataFrame")
}

fn main() {
    let out_dir = Path::new("test_fixtures/random");
    std::fs::create_dir_all(out_dir).expect("create test_fixtures/random");

    let mut rng = StdRng::seed_from_u64(SEED);

    for (i, case) in CASES.iter().enumerate() {
        let mut df = build_case_df(&mut rng, case);
        let path = out_dir.join(format!("case_{i:02}_{}x{}.parquet", case.rows, case.cols));
        let file = File::create(&path).unwrap_or_else(|e| panic!("create {}: {e}", path.display()));
        ParquetWriter::new(file).finish(&mut df).unwrap_or_else(|e| panic!("write {}: {e}", path.display()));
        println!("wrote {} ({} rows x {} cols)", path.display(), case.rows, case.cols);
    }

    println!("done: {} random datasets in {}", CASES.len(), out_dir.display());
}
