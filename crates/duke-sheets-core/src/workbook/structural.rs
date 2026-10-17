use super::*;
use crate::structural::{rewrite_references, AxisEdit, ReferenceEdit};

impl Workbook {
    /// Insert rows before a zero-based index, updating references in every sheet.
    pub fn insert_rows(&mut self, sheet: usize, index: u32, count: u32) -> Result<()> {
        self.edit_axis(
            sheet,
            AxisEdit {
                rows: true,
                index,
                count,
                delete: false,
            },
        )
    }

    /// Delete rows, shrinking ranges and replacing deleted references with #REF!.
    pub fn delete_rows(&mut self, sheet: usize, index: u32, count: u32) -> Result<()> {
        self.edit_axis(
            sheet,
            AxisEdit {
                rows: true,
                index,
                count,
                delete: true,
            },
        )
    }

    /// Insert columns before a zero-based index.
    pub fn insert_columns(&mut self, sheet: usize, index: u16, count: u16) -> Result<()> {
        self.edit_axis(
            sheet,
            AxisEdit {
                rows: false,
                index: index.into(),
                count: count.into(),
                delete: false,
            },
        )
    }

    /// Delete columns, updating references in every sheet.
    pub fn delete_columns(&mut self, sheet: usize, index: u16, count: u16) -> Result<()> {
        self.edit_axis(
            sheet,
            AxisEdit {
                rows: false,
                index: index.into(),
                count: count.into(),
                delete: true,
            },
        )
    }

    fn edit_axis(&mut self, sheet: usize, edit: AxisEdit) -> Result<()> {
        edit.validate()?;
        let ws = self
            .worksheet(sheet)
            .ok_or(Error::SheetOutOfBounds(sheet, self.sheet_count()))?;
        if !edit.delete {
            let dimensions_overflow = if edit.rows {
                ws.custom_row_heights()
                    .keys()
                    .chain(ws.hidden_rows().keys())
                    .chain(ws.row_outline_levels().keys())
                    .chain(ws.collapsed_rows().keys())
                    .any(|index| edit.point(*index).is_none())
            } else {
                ws.custom_column_widths()
                    .keys()
                    .chain(ws.hidden_columns().keys())
                    .chain(ws.column_outline_levels().keys())
                    .chain(ws.collapsed_columns().keys())
                    .any(|index| edit.point((*index).into()).is_none())
            };
            if dimensions_overflow
                || ws
                    .hyperlinks()
                    .keys()
                    .any(|cell| edit.cell(cell.row, cell.col).is_none())
                || ws
                    .comments()
                    .any(|((row, col), _)| edit.cell(row, col).is_none())
            {
                return Err(Error::other(
                    "Insertion would move worksheet metadata off the grid",
                ));
            }
            if let Some(range) = ws.used_range() {
                let last = if edit.rows {
                    range.end.row
                } else {
                    range.end.col.into()
                };
                if edit.point(last).is_none() {
                    return Err(Error::other("Insertion would move cells off the worksheet"));
                }
            }
        }
        let name = ws.name().to_owned();
        self.worksheets[sheet].edit_axis(edit);
        self.rewrite_references(&ReferenceEdit::Axis { sheet: &name, edit });
        self.structural_generation += 1;
        self.calc_cache = None;
        Ok(())
    }

    pub(super) fn rewrite_references(&mut self, change: &ReferenceEdit<'_>) {
        for sheet in &mut self.worksheets {
            sheet.rewrite_references(change);
        }
        for sheet in &mut self.chartsheets {
            let context = sheet.name.clone();
            crate::worksheet::structural::rewrite_chart(&mut sheet.chart, &|text| {
                *text = rewrite_references(text, &context, change)
            });
        }
        let names: Vec<_> = self.named_ranges.iter().cloned().collect();
        for mut name in names {
            let context = match name.scope {
                NameScope::Sheet(index) => self.worksheet(index).map(|s| s.name()).unwrap_or(""),
                NameScope::Workbook => "",
            };
            name.refers_to = rewrite_references(&name.refers_to, context, change);
            self.named_ranges.define_or_update(name);
        }
        self.calc_cache = None;
    }

    pub(super) fn remap_name_scopes(&mut self, remap: impl Fn(usize) -> Option<usize>) {
        let names: Vec<_> = self.named_ranges.iter().cloned().collect();
        self.named_ranges = NamedRangeCollection::new();
        for mut name in names {
            if let NameScope::Sheet(index) = name.scope {
                let Some(index) = remap(index) else {
                    continue;
                };
                name.scope = NameScope::Sheet(index);
            }
            self.named_ranges.define_or_update(name);
        }
    }
}
