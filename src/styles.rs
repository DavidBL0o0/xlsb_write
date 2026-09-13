//! xl/styles.bin builder.
//!
//! `STYLES_BASE` below is a byte-verified reference `styles.bin` blob
//! (1 font, 2 fills, 1 border, 2 cell XFs — General and Date) originally
//! from the `xlsb-writer` crate (MIT License, Copyright (c) 2026 kotucha).
//! Custom `Font`/`Fill`/`Border`/number-format/`BrtXF` records are spliced
//! in on top of it (splice offsets found dynamically via byte search, not
//! hardcoded, so this stays correct if `STYLES_BASE` ever changes).
//!
//! Byte layouts below (BrtFont/BrtFill/BrtBorder/BrtXF/Color) are taken
//! directly from the published [MS-XLSB] spec — see
//! `Formatting: BrtFont 1` / `Formatting: BrtFill 1` / `Formatting: BrtXF 2`
//! on learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/ —
//! and cross-checked against `STYLES_BASE`'s actual bytes record-by-record.

const STYLES_BASE: &[u8] = &[
    0x96, 0x02, 0x00, 0xe3, 0x04, 0x04, 0x01, 0x00, 0x00, 0x00, 0x2b, 0x27, 0xdc, 0x00, 0x00, 0x00, 0x90, 0x01, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x00, 0x07, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0x02, 0x07, 0x00, 0x00, 0x00, 0x43,
    0x00, 0x61, 0x00, 0x6c, 0x00, 0x69, 0x00, 0x62, 0x00, 0x72, 0x00, 0x69, 0x00, 0x25, 0x06, 0x01, 0x00, 0x02, 0x0e,
    0x00, 0x80, 0x81, 0x08, 0x00, 0x26, 0x00, 0xe4, 0x04, 0x00, 0xdb, 0x04, 0x04, 0x02, 0x00, 0x00, 0x00, 0x2d, 0x44,
    0x00, 0x00, 0x00, 0x00, 0x03, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0x03, 0x41, 0x00, 0x00, 0xff, 0xff, 0xff,
    0xff, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x2d, 0x44, 0x11, 0x00, 0x00, 0x00, 0x03, 0x40,
    0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0x03, 0x41, 0x00, 0x00, 0xff, 0xff, 0xff, 0xff, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0xdc, 0x04, 0x00, 0xe5, 0x04, 0x04, 0x01, 0x00, 0x00, 0x00, 0x2e, 0x33, 0x00, 0x00,
    0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xe6, 0x04, 0x00, 0xf2, 0x04, 0x04, 0x01, 0x00,
    0x00, 0x00, 0x2f, 0x10, 0xff, 0xff, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x00,
    0x00, 0xf3, 0x04, 0x00, 0xe9, 0x04, 0x04, 0x02, 0x00, 0x00, 0x00, 0x2f, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x10, 0x10, 0x00, 0x00, 0x2f, 0x10, 0x00, 0x00, 0x0e, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x08, 0x00, 0x25, 0x00, 0xea, 0x04, 0x00, 0xeb, 0x04, 0x04, 0x01, 0x00, 0x00, 0x00,
    0x25, 0x06, 0x01, 0x00, 0x02, 0x11, 0x00, 0x80, 0x80, 0x18, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x26, 0x00, 0x30, 0x18, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00,
    0x00, 0x06, 0x00, 0x00, 0x00, 0x4e, 0x00, 0x6f, 0x00, 0x72, 0x00, 0x6d, 0x00, 0x61, 0x00, 0x6c, 0x00, 0xec, 0x04,
    0x00, 0xf9, 0x03, 0x04, 0x00, 0x00, 0x00, 0x00, 0xfa, 0x03, 0x00, 0xfc, 0x03, 0x50, 0x00, 0x00, 0x00, 0x00, 0x11,
    0x00, 0x00, 0x00, 0x54, 0x00, 0x61, 0x00, 0x62, 0x00, 0x6c, 0x00, 0x65, 0x00, 0x53, 0x00, 0x74, 0x00, 0x79, 0x00,
    0x6c, 0x00, 0x65, 0x00, 0x4d, 0x00, 0x65, 0x00, 0x64, 0x00, 0x69, 0x00, 0x75, 0x00, 0x6d, 0x00, 0x39, 0x00, 0x11,
    0x00, 0x00, 0x00, 0x50, 0x00, 0x69, 0x00, 0x76, 0x00, 0x6f, 0x00, 0x74, 0x00, 0x53, 0x00, 0x74, 0x00, 0x79, 0x00,
    0x6c, 0x00, 0x65, 0x00, 0x4c, 0x00, 0x69, 0x00, 0x67, 0x00, 0x68, 0x00, 0x74, 0x00, 0x31, 0x00, 0x36, 0x00, 0xfd,
    0x03, 0x00, 0x23, 0x04, 0x02, 0x0e, 0x00, 0x00, 0xeb, 0x08, 0x00, 0xf6, 0x08, 0x2a, 0x00, 0x00, 0x00, 0x00, 0x11,
    0x00, 0x00, 0x00, 0x53, 0x00, 0x6c, 0x00, 0x69, 0x00, 0x63, 0x00, 0x65, 0x00, 0x72, 0x00, 0x53, 0x00, 0x74, 0x00,
    0x79, 0x00, 0x6c, 0x00, 0x65, 0x00, 0x4c, 0x00, 0x69, 0x00, 0x67, 0x00, 0x68, 0x00, 0x74, 0x00, 0x31, 0x00, 0xf7,
    0x08, 0x00, 0xec, 0x08, 0x00, 0x24, 0x00, 0x23, 0x04, 0x03, 0x0f, 0x00, 0x00, 0xb0, 0x10, 0x00, 0xb2, 0x10, 0x32,
    0x00, 0x00, 0x00, 0x00, 0x15, 0x00, 0x00, 0x00, 0x54, 0x00, 0x69, 0x00, 0x6d, 0x00, 0x65, 0x00, 0x53, 0x00, 0x6c,
    0x00, 0x69, 0x00, 0x63, 0x00, 0x65, 0x00, 0x72, 0x00, 0x53, 0x00, 0x74, 0x00, 0x79, 0x00, 0x6c, 0x00, 0x65, 0x00,
    0x4c, 0x00, 0x69, 0x00, 0x67, 0x00, 0x68, 0x00, 0x74, 0x00, 0x31, 0x00, 0xb3, 0x10, 0x00, 0xb1, 0x10, 0x00, 0x24,
    0x00, 0x97, 0x02, 0x00,
];

use crate::biff12::{write_vi, write_wstr};
use indexmap::IndexMap;

// ── Record IDs (verified by parsing STYLES_BASE itself — see examples/dump_styles.rs) ──

const RID_BEGIN_FONTS: u32 = 0x0263;
const RID_END_FONTS: u32 = 0x0264;
const RID_FONT: u32 = 0x002B;
const RID_BEGIN_FILLS: u32 = 0x025B;
const RID_END_FILLS: u32 = 0x025C;
const RID_FILL: u32 = 0x002D;
const RID_BEGIN_BORDERS: u32 = 0x0265;
const RID_END_BORDERS: u32 = 0x0266;
const RID_BORDER: u32 = 0x002E;
const RID_BEGIN_CELL_XFS: u32 = 0x0269;
const RID_END_CELL_XFS: u32 = 0x026A;
const RID_XF: u32 = 0x002F;
const RID_FMT: u32 = 0x002C;

// ── Color (MS-XLSB "Color" structure, 8 bytes) ──────────────────────────────
//
// byte0: bit0 = fValidRGB, bits1-7 = xColorType (0=Automatic,1=Indexed,2=RGB,3=Theme)
// byte1: index (meaning depends on xColorType; unused for RGB)
// bytes2-3: nTintAndShade (SHORT)
// bytes4-7: bRed, bGreen, bBlue, bAlpha

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Color {
    pub const fn rgb(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }
    pub const BLACK: Color = Color::rgb(0, 0, 0);
    pub const WHITE: Color = Color::rgb(0xFF, 0xFF, 0xFF);
    pub const RED: Color = Color::rgb(0xFF, 0, 0);
    pub const GREEN: Color = Color::rgb(0, 0x80, 0);
    pub const BLUE: Color = Color::rgb(0, 0, 0xFF);
}

/// xColorType = RGB (2), explicit color.
fn write_color_rgb(c: Color, buf: &mut Vec<u8>) {
    buf.push(0x01 | (0x02 << 1)); // fValidRGB=1, xColorType=RGB
    buf.push(0x00); // index: unused for RGB
    buf.extend_from_slice(&0i16.to_le_bytes()); // nTintAndShade = 0
    buf.extend_from_slice(&[c.r, c.g, c.b, 0xFF]);
}

/// Default font text color: theme color "Text 1" (matches STYLES_BASE's own
/// default font byte-for-byte: `07 01 00 00 00 00 00 ff`).
fn write_color_theme_text1(buf: &mut Vec<u8>) {
    buf.extend_from_slice(&[0x07, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff]);
}

/// Placeholder color for a border side that has no line (style = None).
/// Matches STYLES_BASE's default (no-border) side bytes byte-for-byte:
/// `01 00 00 00 00 00 00 00`.
fn write_color_border_auto(buf: &mut Vec<u8>) {
    buf.extend_from_slice(&[0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]);
}

// ── Border style ─────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum BorderStyle {
    #[default]
    None,
    Thin,
    Medium,
    Dashed,
    Dotted,
    Thick,
    Double,
    Hair,
    MediumDashed,
    DashDot,
    MediumDashDot,
    DashDotDot,
    MediumDashDotDot,
    SlantDashDot,
}

fn border_line_code(s: BorderStyle) -> u16 {
    match s {
        BorderStyle::None => 0,
        BorderStyle::Thin => 1,
        BorderStyle::Medium => 2,
        BorderStyle::Dashed => 3,
        BorderStyle::Dotted => 4,
        BorderStyle::Thick => 5,
        BorderStyle::Double => 6,
        BorderStyle::Hair => 7,
        BorderStyle::MediumDashed => 8,
        BorderStyle::DashDot => 9,
        BorderStyle::MediumDashDot => 10,
        BorderStyle::DashDotDot => 11,
        BorderStyle::MediumDashDotDot => 12,
        BorderStyle::SlantDashDot => 13,
    }
}

// ── Alignment ────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum HAlign {
    #[default]
    General = 0,
    Left = 1,
    Center = 2,
    Right = 3,
    Fill = 4,
    Justify = 5,
    CenterAcrossSelection = 6,
    Distributed = 7,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum VAlign {
    Top = 0,
    Center = 1,
    // Matches the base styles.bin's own default cell XF (alcv = 2).
    #[default]
    Bottom = 2,
    Justify = 3,
    Distributed = 4,
}

// ── Format (the public, user-facing cell style) ─────────────────────────────

#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct Format {
    pub(crate) bold: bool,
    pub(crate) italic: bool,
    pub(crate) underline: bool,
    pub(crate) strikeout: bool,
    pub(crate) font_color: Option<Color>,
    pub(crate) font_size_half_pt: Option<u16>,
    pub(crate) font_name: Option<String>,
    pub(crate) bg_color: Option<Color>,
    pub(crate) border_style: BorderStyle,
    pub(crate) border_color: Option<Color>,
    pub(crate) halign: HAlign,
    pub(crate) valign: VAlign,
    pub(crate) text_wrap: bool,
    pub(crate) num_format: Option<String>,
}

impl Format {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn set_bold(mut self) -> Self {
        self.bold = true;
        self
    }
    pub fn set_italic(mut self) -> Self {
        self.italic = true;
        self
    }
    pub fn set_underline(mut self) -> Self {
        self.underline = true;
        self
    }
    pub fn set_strikeout(mut self) -> Self {
        self.strikeout = true;
        self
    }
    pub fn set_font_color(mut self, color: Color) -> Self {
        self.font_color = Some(color);
        self
    }
    pub fn set_font_size(mut self, points: f32) -> Self {
        self.font_size_half_pt = Some((points * 2.0).round() as u16);
        self
    }
    pub fn set_font_name(mut self, name: &str) -> Self {
        self.font_name = Some(name.to_owned());
        self
    }
    pub fn set_background_color(mut self, color: Color) -> Self {
        self.bg_color = Some(color);
        self
    }
    pub fn set_border(mut self, style: BorderStyle) -> Self {
        self.border_style = style;
        self
    }
    pub fn set_border_color(mut self, color: Color) -> Self {
        self.border_color = Some(color);
        self
    }
    pub fn set_align(mut self, h: HAlign, v: VAlign) -> Self {
        self.halign = h;
        self.valign = v;
        self
    }
    pub fn set_text_wrap(mut self) -> Self {
        self.text_wrap = true;
        self
    }
    pub fn set_num_format(mut self, fmt: &str) -> Self {
        self.num_format = Some(fmt.to_owned());
        self
    }
}

// ── Number-format shorthand table (unchanged from the original crate) ──────

pub fn resolve_shorthand(s: &str) -> &str {
    match s.to_lowercase().as_str() {
        "general" => "General",
        "int" => "#,##0",
        "int0" => "0",
        "float1" => "#,##0.0",
        "float2" => "#,##0.00",
        "float3" => "#,##0.000",
        "float4" => "#,##0.0000",
        "pct" => "0%",
        "pct1" => "0.0%",
        "pct2" => "0.00%",
        "sci" => "0.00E+00",
        "date" => "YYYY-MM-DD",
        "datetime" => "YYYY-MM-DD HH:MM:SS",
        "time" => "HH:MM:SS",
        "text" => "@",
        "accounting" => "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
        "currency" => "$#,##0.00",
        "euro" => "€#,##0.00",
        _ => s,
    }
}

pub fn builtin_ifmt(fmt: &str) -> Option<u16> {
    match fmt {
        "General" => Some(0),
        "0" => Some(1),
        "0.00" => Some(2),
        "#,##0" => Some(3),
        "#,##0.00" => Some(4),
        "0%" => Some(9),
        "0.0%" => Some(10),
        "0.00%" => Some(10),
        "0.00E+00" => Some(11),
        "m/d/yyyy" => Some(14),
        "@" => Some(49),
        _ => None,
    }
}

// ── Record encoders ──────────────────────────────────────────────────────────

const DEFAULT_FONT_SIZE_HALF_PT: u16 = 22; // 11pt
const DEFAULT_FONT_NAME: &str = "Calibri";

#[allow(clippy::too_many_arguments)]
fn encode_font(
    bold: bool,
    italic: bool,
    underline: bool,
    strikeout: bool,
    color: Option<Color>,
    size_half_pt: u16,
    name: &str,
) -> Vec<u8> {
    let mut p = Vec::with_capacity(25 + name.len() * 2);
    let dy_height = (size_half_pt as u32 * 10) as u16; // half-pt * 10 = twips
    p.extend_from_slice(&dy_height.to_le_bytes());
    let mut grbit_low: u8 = 0;
    if italic {
        grbit_low |= 1 << 1;
    }
    if strikeout {
        grbit_low |= 1 << 3;
    }
    p.push(grbit_low);
    p.push(0x00); // grbit high byte (unused3)
    let bls: u16 = if bold { 700 } else { 400 };
    p.extend_from_slice(&bls.to_le_bytes());
    p.extend_from_slice(&0u16.to_le_bytes()); // sss = none
    p.push(if underline { 0x01 } else { 0x00 }); // uls
    p.push(0x02); // bFamily = swiss
    p.push(0x00); // bCharSet
    p.push(0x00); // unused
    match color {
        Some(c) => write_color_rgb(c, &mut p),
        None => write_color_theme_text1(&mut p),
    }
    p.push(0x02); // bFontScheme = minor
    write_wstr(name, &mut p);
    p
}

fn encode_fill_solid(fg: Color) -> Vec<u8> {
    let mut p = Vec::with_capacity(68);
    p.extend_from_slice(&1u32.to_le_bytes()); // fls = solid
    write_color_rgb(fg, &mut p); // brtColorFore (the visible pattern color)
    write_color_rgb(Color::WHITE, &mut p); // brtColorBack
    p.extend_from_slice(&[0u8; 48]); // gradient fields, unused for solid fills
    p
}

fn encode_border(style: BorderStyle, color: Option<Color>) -> Vec<u8> {
    let mut p = Vec::with_capacity(51);
    p.push(0x00); // dg: no diagonal
    let code = border_line_code(style);
    for _side in 0..4 {
        // left, right, top, bottom — same style/color on all 4 sides
        p.extend_from_slice(&code.to_le_bytes());
        match color {
            Some(c) if style != BorderStyle::None => write_color_rgb(c, &mut p),
            _ => write_color_border_auto(&mut p),
        }
    }
    p.extend_from_slice(&0u16.to_le_bytes()); // diagonal: none
    write_color_border_auto(&mut p);
    p
}

fn encode_xf(ifmt: u16, ifont: u16, ifill: u16, iborder: u16, alc: u8, alcv: u8, wrap: bool) -> [u8; 16] {
    let mut p = [0u8; 16];
    // ixfeParent = 0 (parent is the "Normal" cell-style XF)
    p[2..4].copy_from_slice(&ifmt.to_le_bytes());
    p[4..6].copy_from_slice(&ifont.to_le_bytes());
    p[6..8].copy_from_slice(&ifill.to_le_bytes());
    p[8..10].copy_from_slice(&iborder.to_le_bytes());
    // trot = 0, indent = 0
    let mut b12 = (alc & 0x07) | ((alcv & 0x07) << 3);
    if wrap {
        b12 |= 1 << 6;
    }
    p[12] = b12;
    p[13] = 1 << 4; // fLocked = 1
    // xfGrbitAtr = 0
    p
}

// ── Splicing helpers (dynamic byte search, not hardcoded offsets) ──────────

fn find_bytes(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack.windows(needle.len()).position(|w| w == needle)
}

fn find_bytes_from(haystack: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    haystack[from..]
        .windows(needle.len())
        .position(|w| w == needle)
        .map(|i| i + from)
}

/// Insert `extra_payloads` (each wrapped as a `record_rid` record) just before
/// the `end_rid` marker of a `[BrtBeginX (count:u32) ... BrtEndX]` collection,
/// and patch the count. No-op if `extra_payloads` is empty.
fn splice_collection(
    base: &[u8],
    begin_rid: u32,
    end_rid: u32,
    record_rid: u32,
    extra_payloads: &[Vec<u8>],
) -> Vec<u8> {
    if extra_payloads.is_empty() {
        return base.to_vec();
    }
    let mut begin_tag = Vec::new();
    write_vi(begin_rid, &mut begin_tag);
    begin_tag.push(0x04); // BeginX payload is always a single u32 count (length=4)
    let begin_pos =
        find_bytes(base, &begin_tag).unwrap_or_else(|| panic!("begin rid 0x{begin_rid:04x} not found in styles.bin"));
    let count_pos = begin_pos + begin_tag.len();
    let old_count = u32::from_le_bytes(base[count_pos..count_pos + 4].try_into().unwrap());

    let mut end_tag = Vec::new();
    write_vi(end_rid, &mut end_tag);
    let end_pos = find_bytes_from(base, &end_tag, count_pos)
        .unwrap_or_else(|| panic!("end rid 0x{end_rid:04x} not found in styles.bin"));

    let mut extra_bytes = Vec::new();
    for payload in extra_payloads {
        crate::biff12::write_rec(record_rid, payload, &mut extra_bytes);
    }

    let mut out = Vec::with_capacity(base.len() + extra_bytes.len());
    out.extend_from_slice(&base[..count_pos]);
    out.extend_from_slice(&(old_count + extra_payloads.len() as u32).to_le_bytes());
    out.extend_from_slice(&base[count_pos + 4..end_pos]);
    out.extend_from_slice(&extra_bytes);
    out.extend_from_slice(&base[end_pos..]);
    out
}

/// Custom number formats (`BrtFmt` records) have no Begin/End wrapper in
/// practice — this mirrors the splice point used by the original crate's
/// production pipeline (verified against real Excel output there).
fn splice_fmts(base: &[u8], custom_fmts: &IndexMap<String, u16>) -> Vec<u8> {
    if custom_fmts.is_empty() {
        return base.to_vec();
    }
    let mut fmt_bytes = Vec::new();
    for (s, &ifmt) in custom_fmts {
        let mut pay = Vec::new();
        pay.extend_from_slice(&ifmt.to_le_bytes());
        write_wstr(s, &mut pay);
        crate::biff12::write_rec(RID_FMT, &pay, &mut fmt_bytes);
    }
    let mut out = Vec::with_capacity(base.len() + fmt_bytes.len());
    out.extend_from_slice(&base[..3]); // BrtBeginStyleSheet (zero-length record)
    out.extend_from_slice(&fmt_bytes);
    out.extend_from_slice(&base[3..]);
    out
}

// ── StylesBuilder ────────────────────────────────────────────────────────────

type FontKey = (bool, bool, bool, bool, Option<Color>, u16, Option<String>);

#[derive(Default)]
pub struct StylesBuilder {
    custom_fmts: IndexMap<String, u16>,
    font_cache: IndexMap<FontKey, u16>,
    font_records: Vec<Vec<u8>>,
    fill_cache: IndexMap<Color, u16>,
    fill_records: Vec<Vec<u8>>,
    border_cache: IndexMap<(BorderStyle, Option<Color>), u16>,
    border_records: Vec<Vec<u8>>,
    xf_cache: IndexMap<Format, u16>,
    xf_rows: Vec<[u8; 16]>,
}

impl StylesBuilder {
    pub fn new() -> Self {
        Self::default()
    }

    fn resolve_fmt(&mut self, fmt: &str) -> u16 {
        let fmt = resolve_shorthand(fmt);
        if let Some(id) = builtin_ifmt(fmt) {
            return id;
        }
        let next_id = 164 + self.custom_fmts.len() as u16;
        *self.custom_fmts.entry(fmt.to_owned()).or_insert(next_id)
    }

    fn resolve_font(&mut self, fmt: &Format) -> u16 {
        let has_custom = fmt.bold
            || fmt.italic
            || fmt.underline
            || fmt.strikeout
            || fmt.font_color.is_some()
            || fmt.font_size_half_pt.is_some()
            || fmt.font_name.is_some();
        if !has_custom {
            return 0; // the default font already present in STYLES_BASE
        }
        let key: FontKey = (
            fmt.bold,
            fmt.italic,
            fmt.underline,
            fmt.strikeout,
            fmt.font_color,
            fmt.font_size_half_pt.unwrap_or(DEFAULT_FONT_SIZE_HALF_PT),
            fmt.font_name.clone(),
        );
        if let Some(&idx) = self.font_cache.get(&key) {
            return idx;
        }
        let idx = 1 + self.font_records.len() as u16; // font 0 = default
        let name = fmt.font_name.clone().unwrap_or_else(|| DEFAULT_FONT_NAME.to_owned());
        let rec = encode_font(
            fmt.bold,
            fmt.italic,
            fmt.underline,
            fmt.strikeout,
            fmt.font_color,
            fmt.font_size_half_pt.unwrap_or(DEFAULT_FONT_SIZE_HALF_PT),
            &name,
        );
        self.font_records.push(rec);
        self.font_cache.insert(key, idx);
        idx
    }

    fn resolve_fill(&mut self, fmt: &Format) -> u16 {
        let Some(color) = fmt.bg_color else {
            return 0; // "no fill" already present in STYLES_BASE
        };
        if let Some(&idx) = self.fill_cache.get(&color) {
            return idx;
        }
        let idx = 2 + self.fill_records.len() as u16; // fills 0,1 = base (none, gray125)
        self.fill_records.push(encode_fill_solid(color));
        self.fill_cache.insert(color, idx);
        idx
    }

    fn resolve_border(&mut self, fmt: &Format) -> u16 {
        if fmt.border_style == BorderStyle::None {
            return 0; // "no border" already present in STYLES_BASE
        }
        let key = (fmt.border_style, fmt.border_color);
        if let Some(&idx) = self.border_cache.get(&key) {
            return idx;
        }
        let idx = 1 + self.border_records.len() as u16; // border 0 = base (none)
        self.border_records
            .push(encode_border(fmt.border_style, fmt.border_color));
        self.border_cache.insert(key, idx);
        idx
    }

    /// Register a `Format`, returning its cell XF index (deduplicated — two
    /// equal `Format`s always return the same index). `Format::default()`
    /// always maps to `0`, the workbook's built-in default style.
    pub fn register(&mut self, fmt: &Format) -> u16 {
        if *fmt == Format::default() {
            return 0;
        }
        if let Some(&idx) = self.xf_cache.get(fmt) {
            return idx;
        }
        let ifmt = fmt.num_format.as_deref().map(|s| self.resolve_fmt(s)).unwrap_or(0);
        let ifont = self.resolve_font(fmt);
        let ifill = self.resolve_fill(fmt);
        let iborder = self.resolve_border(fmt);
        let row = encode_xf(
            ifmt,
            ifont,
            ifill,
            iborder,
            fmt.halign as u8,
            fmt.valign as u8,
            fmt.text_wrap,
        );

        // Excel's real per-workbook cell-style limit is 65,490 — past that,
        // real Excel itself refuses to add more styles. `idx` is also a
        // `u16`, so without this check a caller who blew past the limit
        // (e.g. constructing a fresh `Format` inside a hot per-cell loop
        // instead of hoisting it out and reusing `&format`) would silently
        // wrap into a garbage, already-used XF index instead of getting any
        // warning — exactly the class of "would otherwise silently corrupt
        // the file" bug this crate's other contracts (`stage_cell`'s row
        // order, Round 6's `check_row_not_flushed`) already guard against.
        const MAX_DISTINCT_STYLES: usize = 65_490;
        assert!(
            self.xf_rows.len() < MAX_DISTINCT_STYLES,
            "xlsb_write: exceeded Excel's real limit of {MAX_DISTINCT_STYLES} distinct cell \
             styles in one workbook — build each `Format` once and pass it by reference to \
             every cell that shares it, rather than constructing a new one per cell."
        );
        let idx = 2 + self.xf_rows.len() as u16; // xf 0,1 = base (General, Date)
        self.xf_rows.push(row);
        self.xf_cache.insert(fmt.clone(), idx);
        idx
    }

    pub fn build(&self) -> Vec<u8> {
        let mut out = splice_fmts(STYLES_BASE, &self.custom_fmts);
        out = splice_collection(&out, RID_BEGIN_FONTS, RID_END_FONTS, RID_FONT, &self.font_records);
        out = splice_collection(&out, RID_BEGIN_FILLS, RID_END_FILLS, RID_FILL, &self.fill_records);
        out = splice_collection(
            &out,
            RID_BEGIN_BORDERS,
            RID_END_BORDERS,
            RID_BORDER,
            &self.border_records,
        );
        let xf_payloads: Vec<Vec<u8>> = self.xf_rows.iter().map(|r| r.to_vec()).collect();
        out = splice_collection(&out, RID_BEGIN_CELL_XFS, RID_END_CELL_XFS, RID_XF, &xf_payloads);
        out
    }
}

pub fn default_styles_bin() -> Vec<u8> {
    StylesBuilder::new().build()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_styles_is_650_bytes() {
        assert_eq!(default_styles_bin().len(), 650);
    }

    #[test]
    fn default_format_registers_as_xf0() {
        let mut sb = StylesBuilder::new();
        assert_eq!(sb.register(&Format::default()), 0);
    }

    #[test]
    fn bold_format_grows_styles_and_dedupes() {
        let mut sb = StylesBuilder::new();
        let bold = Format::new().set_bold();
        let idx1 = sb.register(&bold);
        let idx2 = sb.register(&bold);
        assert_eq!(idx1, idx2, "identical Format must dedupe to the same XF");
        assert!(idx1 >= 2);
        assert!(sb.build().len() > STYLES_BASE.len());
    }

    #[test]
    fn distinct_formats_get_distinct_xfs() {
        let mut sb = StylesBuilder::new();
        let a = sb.register(&Format::new().set_bold());
        let b = sb.register(&Format::new().set_italic());
        let c = sb.register(&Format::new().set_background_color(Color::RED));
        assert_ne!(a, b);
        assert_ne!(b, c);
        assert_ne!(a, c);
    }

    #[test]
    fn built_styles_bin_is_parseable_biff12() {
        let mut sb = StylesBuilder::new();
        sb.register(
            &Format::new()
                .set_bold()
                .set_italic()
                .set_font_color(Color::rgb(10, 20, 30))
                .set_background_color(Color::rgb(200, 200, 0))
                .set_border(BorderStyle::Thin)
                .set_border_color(Color::BLACK)
                .set_align(HAlign::Center, VAlign::Center)
                .set_text_wrap()
                .set_num_format("float2"),
        );
        let bin = sb.build();
        // Must still parse as a valid record stream (no panics, no leftover bytes).
        let recs = crate::biff12::parse_records(&bin);
        assert!(recs.iter().any(|(rid, _)| *rid == 0x0117)); // BrtEndStyleSheet still present
    }
}
