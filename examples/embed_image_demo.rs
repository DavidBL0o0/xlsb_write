//! Demo: embed a small synthetic image into a cell range with a two-cell
//! anchor (the image resizes if you widen/heighten the columns/rows it
//! spans — try it after opening the output in Excel).
//!
//! Builds an 8x8 solid-color PNG entirely in Rust — no external asset
//! file, no image-encoding crate — using a single "stored" (uncompressed)
//! deflate block inside a minimal zlib stream. That's a real, valid PNG;
//! it's just not compressed, which is fine for a handful of solid-color
//! pixels and means this example has zero extra dependencies.
//!
//! Run with: cargo run --example embed_image_demo
//! Then check the result: cargo run --example validate_xlsb -- embed_image_demo.xlsb

use xlsb_write::{Format, HAlign, ImageFormat, VAlign, Workbook};

/// See `tests/embed_image.rs`'s identical helper for the by-hand PNG
/// construction rationale — duplicated here rather than shared, since
/// examples and integration tests are separate compilation units with no
/// shared internal crate to put a 40-line helper in.
fn tiny_png(width: u32, height: u32, rgb: [u8; 3]) -> Vec<u8> {
    fn crc32(data: &[u8]) -> u32 {
        let mut crc: u32 = 0xFFFF_FFFF;
        for &b in data {
            crc ^= b as u32;
            for _ in 0..8 {
                crc = if crc & 1 != 0 {
                    (crc >> 1) ^ 0xEDB8_8320
                } else {
                    crc >> 1
                };
            }
        }
        !crc
    }
    fn adler32(data: &[u8]) -> u32 {
        let (mut a, mut b) = (1u32, 0u32);
        for &byte in data {
            a = (a + byte as u32) % 65521;
            b = (b + a) % 65521;
        }
        (b << 16) | a
    }
    fn chunk(out: &mut Vec<u8>, kind: &[u8; 4], data: &[u8]) {
        out.extend_from_slice(&(data.len() as u32).to_be_bytes());
        out.extend_from_slice(kind);
        out.extend_from_slice(data);
        let mut crc_input = Vec::with_capacity(4 + data.len());
        crc_input.extend_from_slice(kind);
        crc_input.extend_from_slice(data);
        out.extend_from_slice(&crc32(&crc_input).to_be_bytes());
    }

    let mut png = Vec::new();
    png.extend_from_slice(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]);

    let mut ihdr = Vec::new();
    ihdr.extend_from_slice(&width.to_be_bytes());
    ihdr.extend_from_slice(&height.to_be_bytes());
    ihdr.extend_from_slice(&[8, 2, 0, 0, 0]); // 8-bit depth, truecolor RGB
    chunk(&mut png, b"IHDR", &ihdr);

    let mut raw = Vec::new();
    for _ in 0..height {
        raw.extend_from_slice(&[0u8]); // filter type: None
        for _ in 0..width {
            raw.extend_from_slice(&rgb);
        }
    }

    let mut idat = vec![0x78, 0x01]; // zlib header
    let len = raw.len() as u16;
    idat.push(0x01); // deflate stored-block header: BFINAL=1, BTYPE=00
    idat.extend_from_slice(&len.to_le_bytes());
    idat.extend_from_slice(&(!len).to_le_bytes());
    idat.extend_from_slice(&raw);
    idat.extend_from_slice(&adler32(&raw).to_be_bytes());
    chunk(&mut png, b"IDAT", &idat);

    chunk(&mut png, b"IEND", &[]);
    png
}

fn main() -> Result<(), xlsb_write::WriteError> {
    let logo = tiny_png(8, 8, [200, 30, 30]); // solid red square

    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");

    let label = Format::new().set_bold().set_align(HAlign::Left, VAlign::Center);
    sheet.write_string_with_format(0, 0, "Company logo:", &label);

    // Two-cell anchor spanning B2:D6 (rows 1..=5, cols 1..=3, zero-based) —
    // the image's on-screen size tracks these cells' actual width/height,
    // so widening column B or heightening row 3 in Excel grows the image
    // with them, exactly like inserting a picture "into" a cell range.
    sheet.embed_image(1, 1, 5, 3, &logo, ImageFormat::Png);

    sheet.write_string(7, 0, "(image anchored to B2:D6 above)");

    wb.save("embed_image_demo.xlsb")?;
    println!("Wrote embed_image_demo.xlsb — open it in Excel to see the embedded image.");
    Ok(())
}
