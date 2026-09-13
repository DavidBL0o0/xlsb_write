//! OPC parts for embedded images: `xl/media/imageN.<ext>`, `xl/drawings/
//! drawingN.xml` and its `_rels`, and the worksheet's own `_rels` file
//! relating it to the drawing. Everything in this module is regular OOXML
//! XML — even inside an otherwise-binary `.xlsb`, [MS-XLSB] itself scopes
//! the *drawing* part to [ISO/IEC29500-1:2016] (DrawingML XML), same as
//! `.xlsx`; only worksheet/workbook/styles/sharedStrings are BIFF12. The
//! one binary-side hook is `BrtDrawing` (record 550), a single relationship
//! reference written into the sheet's own `.bin` — see `sheet::write_sheet_footer`.
//!
//! **How this shape was derived** (per this project's own hardening
//! culture — see `docs/superpowers/plans/2026-09-13-hardening-and-compatibility-backlog.md`
//! item 10a — never trust a byte/XML shape copied from a single guess or
//! even a spec reading alone): the published MS-XLSB HTML pages don't
//! render the actual ABNF grammar file that would say *where* `BrtDrawing`
//! belongs in the worksheet record stream, so instead of guessing, this
//! session used real Excel (COM automation) to build a `.xlsb` from
//! scratch — `Shapes.AddPicture` into a cell range, `SaveAs` format 50 —
//! then inspected the produced zip directly: `xl/drawings/drawing1.xml`'s
//! exact element shape, `xl/drawings/_rels/drawing1.xml.rels` and
//! `xl/worksheets/_rels/sheet1.bin.rels`'s relationship wiring, `[Content_Types].xml`'s
//! `Default`/`Override` entries, and `BrtDrawing`'s exact byte position in
//! `sheet1.bin` (confirmed via `examples/dump_sheet.rs`: it sits right
//! after the `BrtMargins`/print-options records already baked into
//! `sheet::FOOTER_TAIL_A`, and before the per-sheet FRT-wrapped identifier).
//! Two further simplifications (dropping `<a:xfrm>` and `editAs="oneCell"`
//! from what real Excel writes) were verified the same way: stripped from
//! a copy of that real reference file's `drawing1.xml`, repackaged, and
//! reopened in Excel via COM — identical shape/position, no repair
//! prompt, and the image still resized when the anchor range's column
//! width/row height changed (confirmed genuine two-cell-anchor behavior,
//! not a fixed-size picture that merely tracks its top-left corner).

use crate::{WriteError, zip_options};
use std::io::{Seek, Write};
use zip::ZipWriter;

/// PNG or JPEG — the only two raster formats this crate embeds. Excel
/// supports more (BMP, GIF, TIFF, EMF/WMF, ...) but those are out of scope
/// for this crate's first embedded-image pass (see the backlog item this
/// was scoped from).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImageFormat {
    Png,
    Jpeg,
}

impl ImageFormat {
    /// The file extension used for both `xl/media/imageN.<ext>` and the
    /// matching `[Content_Types].xml` `Default Extension="..."` entry.
    pub(crate) fn extension(self) -> &'static str {
        match self {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpeg",
        }
    }

    /// The MIME type registered for this format's `Default Extension` in
    /// `[Content_Types].xml`.
    pub(crate) fn content_type(self) -> &'static str {
        match self {
            ImageFormat::Png => "image/png",
            ImageFormat::Jpeg => "image/jpeg",
        }
    }

    /// Cheap sanity check: does `bytes` actually start with this format's
    /// magic number? `Worksheet::embed_image` asserts on this — passing
    /// the wrong `ImageFormat` for a given file would silently produce a
    /// worksheet Excel can't decode the picture from (declared PNG bytes
    /// that are actually a JPEG, or vice versa), which is exactly the
    /// class of "looks fine until Excel tries to render it" bug this
    /// crate's own hardening work exists to catch before it ships.
    fn matches_magic_bytes(self, bytes: &[u8]) -> bool {
        match self {
            ImageFormat::Png => bytes.starts_with(&[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]),
            ImageFormat::Jpeg => bytes.starts_with(&[0xFF, 0xD8, 0xFF]),
        }
    }

    pub(crate) fn assert_matches(self, bytes: &[u8]) {
        assert!(
            self.matches_magic_bytes(bytes),
            "xlsb_write: embed_image was called with ImageFormat::{self:?} but the image bytes \
             don't start with that format's magic number — double-check `format` matches the \
             actual file content you passed."
        );
    }
}

/// One image staged on a `Worksheet`, anchored to an inclusive cell range.
/// `media_index` is assigned later, at sheet-finish time (see
/// `write_sheet_drawing_parts`) — not at `embed_image`-call time — because
/// the final `xl/media/imageN` numbering has to be unique across the whole
/// workbook, not just this one sheet, and this sheet doesn't know how many
/// images earlier sheets already used until it's actually being written out.
pub(crate) struct EmbeddedImage {
    pub first_row: u32,
    pub first_col: u32,
    pub last_row: u32,
    pub last_col: u32,
    pub format: ImageFormat,
    pub bytes: Vec<u8>,
}

/// Write every OPC part an image-bearing sheet needs: one `xl/media/imageN.<ext>`
/// per image, `xl/drawings/drawingN.xml` (one drawing part per sheet,
/// containing one `twoCellAnchor` per image), its `_rels`, and the sheet's
/// own `xl/worksheets/_rels/sheetN.bin.rels`. No-op if `images` is empty —
/// a sheet with no embedded pictures gets none of these parts, matching
/// real Excel's own behavior (confirmed against the reference file used to
/// derive this module: a sheet with no picture has no `_rels` file at all).
///
/// `next_media_index` is shared across the whole workbook (threaded in by
/// the caller) so `xl/media/image7.png` on sheet 3 can't collide with
/// `xl/media/image7.png` already written for sheet 1's own images.
pub(crate) fn write_sheet_drawing_parts<W: Write + Seek>(
    images: &[EmbeddedImage],
    sheet_number: usize,
    next_media_index: &mut u32,
    zip: &mut ZipWriter<W>,
) -> Result<(), WriteError> {
    if images.is_empty() {
        return Ok(());
    }

    let mut numbered: Vec<(u32, &EmbeddedImage)> = Vec::with_capacity(images.len());
    for img in images {
        let idx = *next_media_index;
        *next_media_index += 1;
        numbered.push((idx, img));
    }

    for (idx, img) in &numbered {
        zip.start_file(format!("xl/media/image{idx}.{}", img.format.extension()), zip_options())?;
        zip.write_all(&img.bytes)?;
    }

    zip.start_file(format!("xl/drawings/drawing{sheet_number}.xml"), zip_options())?;
    zip.write_all(drawing_xml(&numbered).as_bytes())?;

    zip.start_file(
        format!("xl/drawings/_rels/drawing{sheet_number}.xml.rels"),
        zip_options(),
    )?;
    zip.write_all(drawing_rels(&numbered).as_bytes())?;

    zip.start_file(
        format!("xl/worksheets/_rels/sheet{sheet_number}.bin.rels"),
        zip_options(),
    )?;
    zip.write_all(sheet_rels(sheet_number).as_bytes())?;

    Ok(())
}

/// `xl/drawings/drawingN.xml` — one `<xdr:twoCellAnchor>` per image.
///
/// Anchored from the top-left of `(first_row, first_col)` to the top-left
/// of the cell one past `(last_row, last_col)` (zero pixel/EMU offset at
/// both ends) — this exactly covers the inclusive range
/// `[first_row..=last_row] x [first_col..=last_col]`, and is what makes
/// the image's displayed size track that range's actual column widths/row
/// heights (verified empirically — see this module's top-level doc
/// comment) rather than a fixed size.
///
/// Deliberately omits the `<a:xfrm>` (absolute position/size in EMUs) and
/// `editAs="oneCell"` real Excel itself writes here — both confirmed
/// unnecessary by stripping them from a real reference file and
/// reopening it in Excel: identical shape, no repair prompt, and the
/// image still resized with its anchor range. A `twoCellAnchor`'s
/// `from`/`to` are the only data Excel actually needs; `<a:xfrm>` is a
/// cached absolute-geometry hint Excel recomputes from `from`/`to` on
/// load, and `editAs="oneCell"` merely opts back *into* fixed sizing
/// (the opposite of what "resize with cells" needs), which is why real
/// Excel's own AddPicture-inserted default isn't the shape to copy here.
fn drawing_xml(images: &[(u32, &EmbeddedImage)]) -> String {
    let mut s = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
         <xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" \
         xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
    );
    for (i, (_, img)) in images.iter().enumerate() {
        let shape_id = i as u32 + 2; // id 1 is conventionally reserved for the drawing canvas
        let rel_id = i + 1;
        s.push_str(&format!(
            "<xdr:twoCellAnchor>\
             <xdr:from><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>\
             <xdr:to><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>\
             <xdr:pic>\
             <xdr:nvPicPr><xdr:cNvPr id=\"{shape_id}\" name=\"Picture {shape_id}\"/>\
             <xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>\
             <xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" \
             r:embed=\"rId{rel_id}\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>\
             <xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>\
             </xdr:pic>\
             <xdr:clientData/>\
             </xdr:twoCellAnchor>",
            img.first_col,
            img.first_row,
            img.last_col + 1,
            img.last_row + 1,
        ));
    }
    s.push_str("</xdr:wsDr>");
    s
}

/// `xl/drawings/_rels/drawingN.xml.rels` — one relationship per image,
/// `rId1`, `rId2`, ... in the same order `drawing_xml` referenced them.
fn drawing_rels(images: &[(u32, &EmbeddedImage)]) -> String {
    let mut s = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
         <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for (i, (idx, img)) in images.iter().enumerate() {
        let rel_id = i + 1;
        s.push_str(&format!(
            "<Relationship Id=\"rId{rel_id}\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" \
             Target=\"../media/image{idx}.{}\"/>",
            img.format.extension(),
        ));
    }
    s.push_str("</Relationships>");
    s
}

/// `xl/worksheets/_rels/sheetN.bin.rels` — this crate never emits
/// `xl/worksheets/binaryIndexN.bin` (see backlog item 11), so the drawing
/// is always the sheet's only relationship, and always `rId1`. `BrtDrawing`
/// (in `sheet::write_sheet_footer`) hardcodes that same `"rId1"` string —
/// the two must always agree, since `BrtDrawing`'s payload IS the relationship
/// id this file defines.
fn sheet_rels(sheet_number: usize) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n\
         <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
         <Relationship Id=\"rId1\" \
         Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" \
         Target=\"../drawings/drawing{sheet_number}.xml\"/>\
         </Relationships>"
    )
}
