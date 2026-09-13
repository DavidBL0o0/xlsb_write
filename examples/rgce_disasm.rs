//! Dev tool: disassemble a raw Rgce (formula token stream) hex string,
//! printing each token's byte offset, opcode, and size — to cross-check
//! PtgAttrIf/PtgAttrGoto/PtgAttrIfError offset math against the spec
//! without doing the arithmetic by hand.
//!
//! Run with: cargo run --example rgce_disasm -- "44 04 00 00 00 24 c0 1e ..."

fn main() {
    let hex: String = std::env::args().skip(1).collect::<Vec<_>>().join(" ");
    let bytes: Vec<u8> = hex
        .split_whitespace()
        .map(|b| u8::from_str_radix(b, 16).expect("bad hex byte"))
        .collect();

    let mut i = 0usize;
    while i < bytes.len() {
        let start = i;
        let op = bytes[i];
        match op {
            0x44 | 0x24 | 0x64 => {
                // PtgRef family: row(4) + col(2)
                let row = u32::from_le_bytes(bytes[i + 1..i + 5].try_into().unwrap());
                let colraw = u16::from_le_bytes(bytes[i + 5..i + 7].try_into().unwrap());
                let col = colraw & 0x3FFF;
                let col_rel = (colraw >> 14) & 1;
                let row_rel = (colraw >> 15) & 1;
                println!(
                    "{start:3}: PtgRef  op={op:#04x} row={row} col={col} colRel={col_rel} rowRel={row_rel}  size=7"
                );
                i += 7;
            }
            0x45 | 0x25 | 0x65 => {
                let r0 = u32::from_le_bytes(bytes[i + 1..i + 5].try_into().unwrap());
                let r1 = u32::from_le_bytes(bytes[i + 5..i + 9].try_into().unwrap());
                let c0 = u16::from_le_bytes(bytes[i + 9..i + 11].try_into().unwrap()) & 0x3FFF;
                let c1 = u16::from_le_bytes(bytes[i + 11..i + 13].try_into().unwrap()) & 0x3FFF;
                println!("{start:3}: PtgArea op={op:#04x} r0={r0} r1={r1} c0={c0} c1={c1}  size=13");
                i += 13;
            }
            0x1f => {
                let v = f64::from_le_bytes(bytes[i + 1..i + 9].try_into().unwrap());
                println!("{start:3}: PtgNum  op={op:#04x} v={v}  size=9");
                i += 9;
            }
            0x1e => {
                let v = u16::from_le_bytes(bytes[i + 1..i + 3].try_into().unwrap());
                println!("{start:3}: PtgInt  op={op:#04x} v={v}  size=3");
                i += 3;
            }
            0x17 => {
                let cch = bytes[i + 1] as usize;
                let flag = bytes[i + 2];
                let size = 3 + 2 * cch;
                println!("{start:3}: PtgStr  op={op:#04x} cch={cch} flag={flag:#04x}  size={size}");
                i += size;
            }
            0x03..=0x0e => {
                let name = match op {
                    0x03 => "PtgAdd",
                    0x04 => "PtgSub",
                    0x05 => "PtgMul",
                    0x06 => "PtgDiv",
                    0x07 => "PtgPower",
                    0x08 => "PtgConcat",
                    0x09 => "PtgLt",
                    0x0a => "PtgLe",
                    0x0b => "PtgEq",
                    0x0c => "PtgGe",
                    0x0d => "PtgGt",
                    0x0e => "PtgNe",
                    _ => "PtgOp?",
                };
                println!("{start:3}: {name} op={op:#04x}  size=1");
                i += 1;
            }
            0x19 => {
                let flags = bytes[i + 1];
                let offset = u16::from_le_bytes(bytes[i + 2..i + 4].try_into().unwrap());
                let kind = match flags {
                    0x01 => "bitSemi",
                    0x02 => "bitIf",
                    0x04 => "bitChoose",
                    0x08 => "bitGoto",
                    0x10 => "bitSum",
                    0x20 => "bitBaxcel",
                    0x40 => "bitSpace",
                    0x80 => "bitIfError",
                    other => {
                        println!("{start:3}: PtgAttr op=0x19 UNKNOWN flags={other:#04x}");
                        i += 4;
                        continue;
                    }
                };
                println!("{start:3}: PtgAttr op=0x19 kind={kind} flags={flags:#04x} offset={offset}  size=4");
                i += 4;
            }
            0x42 | 0x22 | 0x62 => {
                let cparams = bytes[i + 1];
                let tab = u16::from_le_bytes(bytes[i + 2..i + 4].try_into().unwrap());
                println!("{start:3}: PtgFuncVar op={op:#04x} cparams={cparams} tab={tab:#06x}  size=4");
                i += 4;
            }
            0x41 | 0x21 | 0x61 => {
                let iftab = u16::from_le_bytes(bytes[i + 1..i + 3].try_into().unwrap());
                println!("{start:3}: PtgFunc op={op:#04x} iftab={iftab:#06x}  size=3");
                i += 3;
            }
            other => {
                println!("{start:3}: UNKNOWN op={other:#04x} — stopping disassembly");
                break;
            }
        }
    }
    println!("-- {} bytes total --", bytes.len());
}
