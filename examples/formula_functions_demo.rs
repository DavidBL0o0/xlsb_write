//! Demo for the 2026-09-13 formula-coverage expansion: string concatenation
//! (`&`/`CONCATENATE`), logical functions (`AND`/`OR`/`NOT`), conditional
//! aggregates (`SUMIF`/`COUNTIF`), text functions (`LEFT`/`RIGHT`/`MID`/
//! `LEN`/`TEXT`), date functions (`TODAY`/`NOW`/`DATE`), and lookup
//! functions (`VLOOKUP`/`INDEX`/`MATCH`).
//!
//! A small synthetic product list sits in columns A-C; every new function is
//! demonstrated once, in its own labeled column, on row 2 (alongside the
//! first product's data) — cells must be written in non-decreasing row
//! order (see GUIDE.md), so every demo formula that references row 2's data
//! lives in row 2 itself, one function per column, rather than spread
//! across separate rows. Every cached value below is hand-computed from the
//! same `PRODUCTS` data the formula reads, matching this crate's own
//! contract (see GUIDE.md's "Writing a formula cell" section).
//!
//! Run with: cargo run --example formula_functions_demo

use xlsb_write::{Format, Formula, Workbook};

struct Product {
    name: &'static str,
    category: &'static str,
    price: f64,
}

const PRODUCTS: &[Product] = &[
    Product {
        name: "Widget",
        category: "Hardware",
        price: 12.5,
    },
    Product {
        name: "Gadget",
        category: "Electronics",
        price: 45.0,
    },
    Product {
        name: "Gizmo",
        category: "Electronics",
        price: 30.0,
    },
    Product {
        name: "Sprocket",
        category: "Hardware",
        price: 8.0,
    },
    Product {
        name: "Widget Pro",
        category: "Hardware",
        price: 22.0,
    },
];

fn main() -> Result<(), xlsb_write::WriteError> {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Products");
    let header = Format::new().set_bold();

    let last_row = PRODUCTS.len() as u32; // rows 1..=last_row hold product data

    // Column layout: A-C are the product table; D onward are one demo
    // column per new function.
    let headers = [
        "Name",
        "Category",
        "Price",
        "",
        "A2&B2",
        "CONCATENATE",
        "AND",
        "OR",
        "NOT",
        "SUMIF(Hardware)",
        "COUNTIF(Hardware)",
        "LEFT",
        "RIGHT",
        "MID",
        "LEN",
        "TEXT",
        "TODAY",
        "NOW",
        "DATE",
        "VLOOKUP(Gizmo)",
        "INDEX",
        "MATCH(Gizmo)",
    ];
    for (col, h) in headers.iter().enumerate() {
        if !h.is_empty() {
            sheet.write_string_with_format(0, col as u32, h, &header);
        }
    }

    // Row 1 (0-based) = first product row + every demo formula, since they
    // all reference row 1's cells or the whole A2:C6 range.
    let p0 = &PRODUCTS[0];
    sheet.write_string(1, 0, p0.name);
    sheet.write_string(1, 1, p0.category);
    sheet.write_number(1, 2, p0.price);

    // `&` operator and `CONCATENATE`.
    let concat_op = Formula::cell(1, 0)
        .concat(Formula::str(" — "))
        .concat(Formula::cell(1, 1));
    sheet.write_formula_str(1, 4, concat_op, &format!("{} — {}", p0.name, p0.category));
    let concatenate_fn = Formula::concatenate(vec![Formula::cell(1, 0), Formula::str("/"), Formula::cell(1, 1)]);
    sheet.write_formula_str(1, 5, concatenate_fn, &format!("{}/{}", p0.name, p0.category));

    // AND / OR / NOT.
    let and_f = Formula::and(vec![
        Formula::cell(1, 1).eq(Formula::str("Hardware")),
        Formula::cell(1, 2).lt(Formula::num(15.0)),
    ]);
    sheet.write_formula_bool(1, 6, and_f, p0.category == "Hardware" && p0.price < 15.0);
    let or_f = Formula::or(vec![
        Formula::cell(1, 1).eq(Formula::str("Electronics")),
        Formula::cell(1, 2).gt(Formula::num(40.0)),
    ]);
    sheet.write_formula_bool(1, 7, or_f, p0.category == "Electronics" || p0.price > 40.0);
    let not_f = Formula::not(Formula::cell(1, 1).eq(Formula::str("Hardware")));
    sheet.write_formula_bool(1, 8, not_f, p0.category != "Hardware");

    // SUMIF / COUNTIF over the whole product range. SUMIF here uses the
    // 3-arg form (criteria_range, criteria, sum_range) to total a
    // *different* column than the one being matched — `Formula::sumif`
    // only covers the 2-arg "sum the matched range itself" form, so the
    // 3-arg form is built directly via `Formula::Func`, per its own doc
    // comment. This exact 3-arg shape (`PtgFuncVar`, cparams=3, same
    // Ftab=345) was separately confirmed byte-for-byte against real Excel.
    let category_range = Formula::range(1, 1, last_row, 1);
    let price_range = Formula::range(1, 2, last_row, 2);
    let hardware_total: f64 = PRODUCTS
        .iter()
        .filter(|p| p.category == "Hardware")
        .map(|p| p.price)
        .sum();
    let sumif_3arg = Formula::Func(
        xlsb_write::FnIndex::SUMIF,
        vec![category_range.clone(), Formula::str("Hardware"), price_range],
    );
    sheet.write_formula_num(1, 9, sumif_3arg, hardware_total);
    let hardware_count = PRODUCTS.iter().filter(|p| p.category == "Hardware").count() as f64;
    sheet.write_formula_num(
        1,
        10,
        Formula::countif(category_range, Formula::str("Hardware")),
        hardware_count,
    );

    // LEFT / RIGHT / MID / LEN / TEXT on the first product's name/price.
    sheet.write_formula_str(
        1,
        11,
        Formula::left(Formula::cell(1, 0), Formula::num(3.0)),
        &p0.name[..3],
    );
    sheet.write_formula_str(
        1,
        12,
        Formula::right(Formula::cell(1, 0), Formula::num(3.0)),
        &p0.name[p0.name.len() - 3..],
    );
    sheet.write_formula_str(
        1,
        13,
        Formula::mid(Formula::cell(1, 0), Formula::num(2.0), Formula::num(3.0)),
        &p0.name.chars().skip(1).take(3).collect::<String>(),
    );
    sheet.write_formula_num(1, 14, Formula::len(Formula::cell(1, 0)), p0.name.len() as f64);
    sheet.write_formula_str(
        1,
        15,
        Formula::text(Formula::cell(1, 2), Formula::str("$0.00")),
        &format!("${:.2}", p0.price),
    );

    // TODAY() / NOW() / DATE(...). TODAY/NOW are volatile — Excel
    // recalculates them on open, so the cached value below is only what's
    // shown before that first recalculation (see GUIDE.md). DATE(...) is
    // not volatile; 46278 is the exact Excel serial for 2026-09-13.
    sheet.write_formula_num(1, 16, Formula::today(), 46278.0);
    sheet.write_formula_num(1, 17, Formula::now(), 46278.5);
    sheet.write_formula_num(
        1,
        18,
        Formula::date(Formula::num(2026.0), Formula::num(9.0), Formula::num(13.0)),
        46278.0,
    );

    // VLOOKUP / INDEX / MATCH looking up "Gizmo" in the product table.
    let table = Formula::range(1, 0, last_row, 2);
    let lookup_name = "Gizmo";
    let lookup_row = PRODUCTS.iter().position(|p| p.name == lookup_name).unwrap();
    sheet.write_formula_num(
        1,
        19,
        Formula::vlookup(Formula::str(lookup_name), table.clone(), Formula::num(3.0), false),
        PRODUCTS[lookup_row].price,
    );
    sheet.write_formula_num(
        1,
        20,
        Formula::index(table, Formula::num((lookup_row + 1) as f64), Formula::num(3.0)),
        PRODUCTS[lookup_row].price,
    );
    sheet.write_formula_num(
        1,
        21,
        Formula::match_(
            Formula::str(lookup_name),
            Formula::range(1, 0, last_row, 0),
            Formula::num(0.0),
        ),
        (lookup_row + 1) as f64,
    );

    // Remaining product rows (no demo formulas needed on these — one
    // demonstration of each function is enough).
    for (i, p) in PRODUCTS.iter().enumerate().skip(1) {
        let row = i as u32 + 1;
        sheet.write_string(row, 0, p.name);
        sheet.write_string(row, 1, p.category);
        sheet.write_number(row, 2, p.price);
    }

    wb.save("formula_functions_demo.xlsb")?;
    println!("Wrote formula_functions_demo.xlsb");
    Ok(())
}
