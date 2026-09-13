//! Structural + OPC-wiring tests for `Worksheet::embed_image`.
//!
//! `calamine` (this crate's usual independent-reader check) doesn't surface
//! embedded images at all — it's a cell-data reader, not a drawing reader —
//! so these tests instead verify the actual OPC parts this crate emits:
//! `xl/media/imageN.<ext>`, `xl/drawings/drawingN.xml` + its `_rels`,
//! `xl/worksheets/_rels/sheetN.bin.rels`, the `[Content_Types].xml`
//! additions, and the `BrtDrawing` record inside the sheet's own `.bin`
//! (parsed back with `xlsb_write::biff12`, the same structural check
//! `examples/validate_xlsb.rs` runs). Real-Excel visual/COM verification
//! for this feature is documented separately (see the session's report),
//! since it needs Excel installed and isn't something `cargo test` can
//! assert on in CI.

use std::io::Cursor;
use xlsb_write::{ImageFormat, StreamingWorkbook, Workbook};
use zip::ZipArchive;

/// A tiny (2x2) valid PNG, built by hand: PNG signature + IHDR + one IDAT
/// chunk (a zlib stream wrapping a single *stored* — i.e. uncompressed —
/// deflate block) + IEND. No image-encoding crate needed for a handful of
/// pixels; the "stored" deflate block type exists exactly for this.
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
    ihdr.extend_from_slice(&[8, 2, 0, 0, 0]); // 8-bit depth, truecolor RGB, default compression/filter/interlace
    chunk(&mut png, b"IHDR", &ihdr);

    let mut raw = Vec::new();
    for _ in 0..height {
        raw.extend_from_slice(&[0u8]); // filter type: None
        for _ in 0..width {
            raw.extend_from_slice(&rgb);
        }
    }
    assert!(
        raw.len() <= u16::MAX as usize,
        "tiny_png: image too large for a single stored deflate block"
    );

    let mut idat = vec![0x78, 0x01]; // zlib header (CMF, FLG)
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

/// A tiny (few-byte) JPEG isn't hand-buildable the way a stored-deflate PNG
/// is (JPEG has no "uncompressed" mode simple enough to hand-roll), so JPEG
/// coverage here just needs *some* bytes that start with the real SOI/JFIF
/// magic number `embed_image`'s format-mismatch guard checks — this crate
/// never decodes the image, so a minimal (truncated, non-renderable) JPEG
/// is sufficient to exercise the OPC-wiring/BrtDrawing code paths under test.
fn tiny_jpeg_magic_bytes() -> Vec<u8> {
    vec![0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, b'J', b'F', b'I', b'F', 0x00]
}

struct ExtractedZip {
    names: Vec<String>,
    files: std::collections::HashMap<String, Vec<u8>>,
}

fn extract(bytes: &[u8]) -> ExtractedZip {
    let cursor = Cursor::new(bytes);
    let mut zip = ZipArchive::new(cursor).unwrap();
    let mut names = Vec::new();
    let mut files = std::collections::HashMap::new();
    for i in 0..zip.len() {
        use std::io::Read;
        let mut entry = zip.by_index(i).unwrap();
        let name = entry.name().to_string();
        let mut buf = Vec::new();
        entry.read_to_end(&mut buf).unwrap();
        names.push(name.clone());
        files.insert(name, buf);
    }
    ExtractedZip { names, files }
}

#[test]
fn workbook_embed_image_produces_expected_opc_wiring() {
    let png = tiny_png(2, 2, [200, 30, 30]);

    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.write_string(0, 0, "Logo:");
    sheet.embed_image(1, 1, 3, 3, &png, ImageFormat::Png);

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let z = extract(&buf.into_inner());

    assert!(z.names.contains(&"xl/media/image1.png".to_string()), "{:?}", z.names);
    assert_eq!(
        z.files["xl/media/image1.png"], png,
        "media part must contain the exact original bytes"
    );

    assert!(
        z.names.contains(&"xl/drawings/drawing1.xml".to_string()),
        "{:?}",
        z.names
    );
    let drawing_xml = String::from_utf8(z.files["xl/drawings/drawing1.xml"].clone()).unwrap();
    assert!(drawing_xml.contains("<xdr:twoCellAnchor>"));
    assert!(
        drawing_xml.contains("<xdr:col>1</xdr:col>"),
        "from.col must be first_col"
    );
    assert!(
        drawing_xml.contains("<xdr:row>1</xdr:row>"),
        "from.row must be first_row"
    );
    assert!(
        drawing_xml.contains("<xdr:col>4</xdr:col>"),
        "to.col must be last_col+1"
    );
    assert!(
        drawing_xml.contains("<xdr:row>4</xdr:row>"),
        "to.row must be last_row+1"
    );
    assert!(drawing_xml.contains("r:embed=\"rId1\""));

    assert!(
        z.names.contains(&"xl/drawings/_rels/drawing1.xml.rels".to_string()),
        "{:?}",
        z.names
    );
    let drawing_rels = String::from_utf8(z.files["xl/drawings/_rels/drawing1.xml.rels"].clone()).unwrap();
    assert!(drawing_rels.contains("Target=\"../media/image1.png\""));
    assert!(drawing_rels.contains("relationships/image"));

    assert!(
        z.names.contains(&"xl/worksheets/_rels/sheet1.bin.rels".to_string()),
        "{:?}",
        z.names
    );
    let sheet_rels = String::from_utf8(z.files["xl/worksheets/_rels/sheet1.bin.rels"].clone()).unwrap();
    assert!(sheet_rels.contains("Target=\"../drawings/drawing1.xml\""));
    assert!(sheet_rels.contains("relationships/drawing"));

    let ct = String::from_utf8(z.files["[Content_Types].xml"].clone()).unwrap();
    assert!(ct.contains("Extension=\"png\" ContentType=\"image/png\""));
    assert!(ct.contains("PartName=\"/xl/drawings/drawing1.xml\""));
    assert!(ct.contains("drawing+xml"));

    // The sheet's own .bin must parse as well-formed BIFF12 (structural
    // check, same as examples/validate_xlsb.rs) and contain exactly one
    // BrtDrawing (record 550) with the expected "rId1" payload.
    let sheet_bin = &z.files["xl/worksheets/sheet1.bin"];
    let recs = xlsb_write::biff12::try_parse_records(sheet_bin).expect("sheet1.bin must be well-formed BIFF12");
    let drawing_recs: Vec<&Vec<u8>> = recs.iter().filter(|(rid, _)| *rid == 550).map(|(_, p)| p).collect();
    assert_eq!(drawing_recs.len(), 1, "exactly one BrtDrawing record expected");
    // XLNullableWideString "rId1": cch=4 (u32 LE) + UTF-16LE "rId1".
    let expected_payload: Vec<u8> = {
        let mut p = 4u32.to_le_bytes().to_vec();
        for ch in "rId1".encode_utf16() {
            p.extend_from_slice(&ch.to_le_bytes());
        }
        p
    };
    assert_eq!(*drawing_recs[0], expected_payload);
}

/// A sheet with no images must get none of the drawing-related OPC parts —
/// confirms the feature is fully opt-in and doesn't change existing output.
#[test]
fn sheet_without_images_gets_no_drawing_parts() {
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.write_string(0, 0, "No pictures here");

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let z = extract(&buf.into_inner());

    assert!(!z.names.iter().any(|n| n.starts_with("xl/media/")));
    assert!(!z.names.iter().any(|n| n.starts_with("xl/drawings/")));
    assert!(!z.names.contains(&"xl/worksheets/_rels/sheet1.bin.rels".to_string()));
    let ct = String::from_utf8(z.files["[Content_Types].xml"].clone()).unwrap();
    assert!(!ct.contains("png"));
    assert!(!ct.contains("drawing"));
}

/// Multiple images on the same sheet each get their own `twoCellAnchor` and
/// their own relationship id within one shared `drawingN.xml`/`.rels` pair.
#[test]
fn multiple_images_on_one_sheet_share_one_drawing_part() {
    let png = tiny_png(2, 2, [10, 200, 10]);
    let jpeg = tiny_jpeg_magic_bytes();

    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.embed_image(0, 0, 2, 2, &png, ImageFormat::Png);
    sheet.embed_image(5, 5, 7, 7, &jpeg, ImageFormat::Jpeg);

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let z = extract(&buf.into_inner());

    assert!(z.names.contains(&"xl/media/image1.png".to_string()));
    assert!(z.names.contains(&"xl/media/image2.jpeg".to_string()));
    // Exactly one drawing part for the whole sheet, not one per image.
    assert_eq!(
        z.names
            .iter()
            .filter(|n| n.starts_with("xl/drawings/drawing") && n.ends_with(".xml"))
            .count(),
        1
    );
    let drawing_xml = String::from_utf8(z.files["xl/drawings/drawing1.xml"].clone()).unwrap();
    assert_eq!(drawing_xml.matches("<xdr:twoCellAnchor>").count(), 2);
    assert!(drawing_xml.contains("r:embed=\"rId1\""));
    assert!(drawing_xml.contains("r:embed=\"rId2\""));

    let ct = String::from_utf8(z.files["[Content_Types].xml"].clone()).unwrap();
    assert!(ct.contains("Extension=\"png\""));
    assert!(ct.contains("Extension=\"jpeg\""));
}

/// Media numbering must stay unique across sheets in the same workbook —
/// a second sheet's images continue the counter rather than restarting it.
#[test]
fn media_numbering_is_unique_across_sheets() {
    let png1 = tiny_png(2, 2, [1, 2, 3]);
    let png2 = tiny_png(2, 2, [4, 5, 6]);

    let mut wb = Workbook::new();
    wb.add_worksheet("Sheet1")
        .embed_image(0, 0, 1, 1, &png1, ImageFormat::Png);
    wb.add_worksheet("Sheet2")
        .embed_image(0, 0, 1, 1, &png2, ImageFormat::Png);

    let mut buf = Cursor::new(Vec::new());
    wb.write(&mut buf).unwrap();
    let z = extract(&buf.into_inner());

    assert_eq!(z.files["xl/media/image1.png"], png1);
    assert_eq!(z.files["xl/media/image2.png"], png2);
    assert!(z.names.contains(&"xl/drawings/drawing1.xml".to_string()));
    assert!(z.names.contains(&"xl/drawings/drawing2.xml".to_string()));
}

/// Same feature via `StreamingWorkbook` — must produce equivalent OPC
/// wiring to the `Workbook` path (see GUIDE.md: "Both produce byte-identical
/// output for the same sequence of calls").
#[test]
fn streaming_workbook_embed_image_produces_expected_opc_wiring() {
    let png = tiny_png(2, 2, [200, 30, 30]);

    let mut buf = Cursor::new(Vec::new());
    let mut wb = StreamingWorkbook::create(&mut buf);
    let mut sheet = wb.new_worksheet("Sheet1");
    sheet.write_string(0, 0, "Logo:");
    sheet.embed_image(1, 1, 3, 3, &png, ImageFormat::Png);
    wb.finish_worksheet(sheet).unwrap();
    wb.finish().unwrap();

    let z = extract(&buf.into_inner());
    assert_eq!(z.files["xl/media/image1.png"], png);
    assert!(z.names.contains(&"xl/drawings/drawing1.xml".to_string()));
    assert!(z.names.contains(&"xl/worksheets/_rels/sheet1.bin.rels".to_string()));
    let ct = String::from_utf8(z.files["[Content_Types].xml"].clone()).unwrap();
    assert!(ct.contains("Extension=\"png\""));
}

/// Same feature via `SizedStreamingWorksheet` (the "one row buffered at a
/// time" streaming type) — `embed_image` must work even though the header
/// may already have been sent by the time it's called.
#[test]
fn sized_streaming_worksheet_embed_image_produces_expected_opc_wiring() {
    let png = tiny_png(2, 2, [30, 30, 200]);

    let mut buf = Cursor::new(Vec::new());
    let mut wb = StreamingWorkbook::create(&mut buf);
    let mut sheet = wb.new_worksheet_sized("Sheet1", 10, 10);
    sheet.write_string(0, 0, "row0").unwrap();
    sheet.write_string(1, 0, "row1").unwrap(); // flushes row 0, sends the header
    sheet.embed_image(2, 2, 4, 4, &png, ImageFormat::Png); // called AFTER header already sent
    sheet.finish().unwrap();
    wb.finish().unwrap();

    let z = extract(&buf.into_inner());
    assert_eq!(z.files["xl/media/image1.png"], png);
    assert!(z.names.contains(&"xl/drawings/drawing1.xml".to_string()));
    let drawing_xml = String::from_utf8(z.files["xl/drawings/drawing1.xml"].clone()).unwrap();
    assert!(drawing_xml.contains("<xdr:col>2</xdr:col>"));
    assert!(drawing_xml.contains("<xdr:row>2</xdr:row>"));
}

#[test]
#[should_panic(expected = "magic number")]
fn embed_image_rejects_mismatched_format() {
    let png = tiny_png(2, 2, [1, 1, 1]);
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.embed_image(0, 0, 1, 1, &png, ImageFormat::Jpeg); // wrong format for PNG bytes
}

#[test]
#[should_panic(expected = "inverted")]
fn embed_image_rejects_inverted_range() {
    let png = tiny_png(2, 2, [1, 1, 1]);
    let mut wb = Workbook::new();
    let sheet = wb.add_worksheet("Sheet1");
    sheet.embed_image(5, 5, 1, 1, &png, ImageFormat::Png); // last < first
}
