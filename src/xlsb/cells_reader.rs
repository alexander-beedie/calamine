// SPDX-License-Identifier: MIT
//
// Copyright 2016-2025, Johann Tuffe.

use std::io::{Read, Seek};

use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat},
    utils::{read_f64, read_i32, read_u32, read_usize},
    Cell, CellErrorType, Dimensions, XlsbError,
};

use super::{cell_format, find_ptgexp, parse_formula, wide_str, RecordIter};

/// Stored shared formula: (rwFirst, rwLast, colFirst, colLast, token_data).
type ShrFmla = (u32, u32, u32, u32, Vec<u8>);

/// A cells reader for xlsb files
pub struct XlsbCellsReader<'a, RS>
where
    RS: Read + Seek,
{
    iter: RecordIter<'a, RS>,
    formats: &'a [CellFormat],
    strings: &'a [String],
    extern_sheets: &'a [String],
    metadata_names: &'a [(String, String)],
    typ: u16,
    row: u32,
    is_1904: bool,
    dimensions: Dimensions,
    buf: Vec<u8>,
    /// Shared formula definitions collected during formula iteration.
    shared_formulas: Vec<ShrFmla>,
    /// PtgExp cells pending resolution: (cell_pos, host_row).
    ptgexp_cells: Vec<((u32, u32), u32)>,
    /// Resolved PtgExp formulas being drained (in reverse order, via pop()).
    resolved: Vec<Cell<String>>,
    /// True once BrtEndSheetData has been reached.
    end_of_sheet: bool,
}

impl<'a, RS> XlsbCellsReader<'a, RS>
where
    RS: Read + Seek,
{
    pub(crate) fn new(
        mut iter: RecordIter<'a, RS>,
        formats: &'a [CellFormat],
        strings: &'a [String],
        extern_sheets: &'a [String],
        metadata_names: &'a [(String, String)],
        is_1904: bool,
    ) -> Result<Self, XlsbError> {
        let mut buf = Vec::with_capacity(1024);
        // BrtWsDim
        let _ = iter.next_skip_blocks(
            0x0094,
            &[
                (0x0081, None), // BrtBeginSheet
                (0x0093, None), // BrtWsProp
            ],
            &mut buf,
        )?;
        let dimensions = parse_dimensions(&buf[..16]);

        // BrtBeginSheetData
        let _ = iter.next_skip_blocks(
            0x0091,
            &[
                (0x0085, Some(0x0086)), // Views
                (0x0025, Some(0x0026)), // AC blocks
                (0x01E5, None),         // BrtWsFmtInfo
                (0x0186, Some(0x0187)), // Col Infos
            ],
            &mut buf,
        )?;

        Ok(XlsbCellsReader {
            iter,
            formats,
            is_1904,
            strings,
            extern_sheets,
            metadata_names,
            dimensions,
            typ: 0,
            row: 0,
            buf,
            shared_formulas: Vec::new(),
            ptgexp_cells: Vec::new(),
            resolved: Vec::new(),
            end_of_sheet: false,
        })
    }

    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    pub fn next_cell(&mut self) -> Result<Option<Cell<DataRef<'a>>>, XlsbError> {
        // loop until end of sheet
        let value = loop {
            self.buf.clear();
            self.typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;
            let value = match self.typ {
                // 0x0001 => continue, // Data::Empty, // BrtCellBlank
                0x0002 => {
                    // BrtCellRk MS-XLSB 2.5.122
                    let d100 = (self.buf[8] & 1) != 0;
                    let is_int = (self.buf[8] & 2) != 0;
                    self.buf[8] &= 0xFC;

                    if is_int {
                        let v = (read_i32(&self.buf[8..12]) >> 2) as i64;
                        if d100 {
                            let v = (v as f64) / 100.0;
                            format_excel_f64_ref(
                                v,
                                cell_format(self.formats, &self.buf),
                                self.is_1904,
                            )
                        } else {
                            DataRef::Int(v)
                        }
                    } else {
                        let mut v = [0u8; 8];
                        v[4..].copy_from_slice(&self.buf[8..12]);
                        let v = read_f64(&v);
                        let v = if d100 { v / 100.0 } else { v };
                        format_excel_f64_ref(v, cell_format(self.formats, &self.buf), self.is_1904)
                    }
                }
                0x0003 => {
                    let error = match self.buf[8] {
                        0x00 => CellErrorType::Null,
                        0x07 => CellErrorType::Div0,
                        0x0F => CellErrorType::Value,
                        0x17 => CellErrorType::Ref,
                        0x1D => CellErrorType::Name,
                        0x24 => CellErrorType::Num,
                        0x2A => CellErrorType::NA,
                        0x2B => CellErrorType::GettingData,
                        c => return Err(XlsbError::CellError(c)),
                    };
                    // BrtCellError
                    DataRef::Error(error)
                }
                0x0004 | 0x000A => DataRef::Bool(self.buf[8] != 0), // BrtCellBool or BrtFmlaBool
                0x0005 | 0x0009 => {
                    let v = read_f64(&self.buf[8..16]);
                    format_excel_f64_ref(v, cell_format(self.formats, &self.buf), self.is_1904)
                } // BrtCellReal or BrtFmlaNum
                0x0006 | 0x0008 => DataRef::String(wide_str(&self.buf[8..], &mut 0)?.into_owned()), // BrtCellSt or BrtFmlaString
                0x0007 => {
                    // BrtCellIsst
                    let isst = read_usize(&self.buf[8..12]);
                    DataRef::SharedString(&self.strings[isst])
                }
                0x0000 => {
                    // BrtRowHdr
                    self.row = read_u32(&self.buf);
                    if self.row > 0x0010_0000 {
                        return Ok(None); // invalid row
                    }
                    continue;
                }
                0x0092 => return Ok(None), // BrtEndSheetData
                _ => continue, // anything else, ignore and try next, without changing idx
            };
            break value;
        };
        let col = read_u32(&self.buf);
        Ok(Some(Cell::new((self.row, col), value)))
    }

    /// Extract the rgce token slice from a formula record buffer.
    /// Returns the rgce slice based on the record type.
    fn extract_rgce(&self) -> Option<&[u8]> {
        let cce_offset = match self.typ {
            0x0009 => 18, // BrtFmlaNum: col(4)+style(4)+value(8)+flags(2)
            0x0008 => {
                // BrtFmlaString: col(4)+style(4)+cch(4)+str(cch*2)+flags(2)
                let cch = read_u32(&self.buf[8..]) as usize;
                14 + cch * 2
            }
            0x000A | 0x000B => 11, // BrtFmlaBool/Error: col(4)+style(4)+val(1)+flags(2)
            _ => return None,
        };
        if cce_offset + 4 > self.buf.len() {
            return None;
        }
        let cce = read_u32(&self.buf[cce_offset..]) as usize;
        let start = cce_offset + 4;
        if start + cce > self.buf.len() {
            return None;
        }
        Some(&self.buf[start..start + cce])
    }

    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsbError> {
        // Drain any resolved PtgExp formulas first
        if let Some(cell) = self.resolved.pop() {
            return Ok(Some(cell));
        }
        if self.end_of_sheet {
            return Ok(None);
        }

        let value = loop {
            self.typ = self.iter.read_type()?;
            let _ = self.iter.fill_buffer(&mut self.buf)?;

            let value = match self.typ {
                0x0008..=0x000B => {
                    let col = read_u32(&self.buf);
                    let cell_pos = (self.row, col);

                    if let Some(rgce) = self.extract_rgce() {
                        if let Some(host_row) = find_ptgexp(rgce) {
                            // PtgExp: defer resolution until we've seen all BrtShrFmla records
                            self.ptgexp_cells.push((cell_pos, host_row));
                            continue;
                        }
                        parse_formula(rgce, self.extern_sheets, self.metadata_names, None)?
                    } else {
                        String::new()
                    }
                }
                // BrtShrFmla: shared formula definition
                // Layout: rwFirst(4) + rwLast(4) + colFirst(4) + colLast(4) + cce(4) + rgce(cce)
                0x01AB => {
                    if self.buf.len() >= 20 {
                        let rw_first = read_u32(&self.buf[0..4]);
                        let rw_last = read_u32(&self.buf[4..8]);
                        let col_first = read_u32(&self.buf[8..12]);
                        let col_last = read_u32(&self.buf[12..16]);
                        let cce = read_u32(&self.buf[16..20]) as usize;
                        if 20 + cce <= self.buf.len() {
                            let tokens = self.buf[20..20 + cce].to_vec();
                            self.shared_formulas
                                .push((rw_first, rw_last, col_first, col_last, tokens));
                        }
                    }
                    continue;
                }
                0x0000 => {
                    // BrtRowHdr
                    self.row = read_u32(&self.buf);
                    if self.row > 0x0010_0000 {
                        return self.resolve_pending();
                    }
                    continue;
                }
                0x0092 => {
                    // BrtEndSheetData — resolve all pending PtgExp cells
                    return self.resolve_pending();
                }
                _ => continue,
            };
            break value;
        };
        let col = read_u32(&self.buf);
        Ok(Some(Cell::new((self.row, col), value)))
    }

    /// Resolve all pending PtgExp cells and start returning them.
    fn resolve_pending(&mut self) -> Result<Option<Cell<String>>, XlsbError> {
        self.end_of_sheet = true;
        let pending = std::mem::take(&mut self.ptgexp_cells);
        for (cell_pos, host_row) in pending {
            // Find the shared formula whose range contains this cell
            let shr =
                self.shared_formulas
                    .iter()
                    .find(|(rw_first, rw_last, col_first, col_last, _)| {
                        *rw_first == host_row
                            && cell_pos.0 <= *rw_last
                            && cell_pos.1 >= *col_first
                            && cell_pos.1 <= *col_last
                    });

            if let Some((_, _, _, _, tokens)) = shr {
                let fmla = parse_formula(
                    tokens,
                    self.extern_sheets,
                    self.metadata_names,
                    Some(cell_pos),
                )
                .unwrap_or_else(|e| {
                    format!(
                        "Unrecognised shared formula for cell ({}, {}): {e:?}",
                        cell_pos.0, cell_pos.1
                    )
                });
                self.resolved.push(Cell::new(cell_pos, fmla));
            }
        }
        // Reverse so pop() yields cells in forward order
        self.resolved.reverse();
        Ok(self.resolved.pop())
    }
}

fn parse_dimensions(buf: &[u8]) -> Dimensions {
    Dimensions {
        start: (read_u32(&buf[0..4]), read_u32(&buf[8..12])),
        end: (read_u32(&buf[4..8]), read_u32(&buf[12..16])),
    }
}
