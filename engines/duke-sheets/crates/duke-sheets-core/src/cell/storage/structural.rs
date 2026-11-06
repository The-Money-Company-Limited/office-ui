use super::*;
use crate::structural::{rewrite_references, AxisEdit, ReferenceEdit};

impl CellStorage {
    pub(crate) fn edit_axis(&mut self, edit: AxisEdit) {
        self.cells = std::mem::take(&mut self.cells)
            .into_iter()
            .filter_map(|((row, col), mut cell)| {
                if matches!(cell.value, CellValue::SpillTarget { .. }) {
                    cell.value = CellValue::Empty;
                }
                edit.cell(row, col).map(|key| (key, cell))
            })
            .collect();
        // A Ctrl+Shift+Enter block moves with its top-left cell and grows or shrinks with the
        // rows or columns inserted or deleted inside it.
        self.formulas = std::mem::take(&mut self.formulas)
            .into_iter()
            .filter_map(|((row, col), mut formula)| {
                let key = edit.cell(row, col)?;
                if let crate::cell::FormulaKind::Cse { rows, cols } = formula.kind {
                    let block = edit.range(crate::cell::CellRange::from_indices(row, col, row + rows - 1, col + cols - 1))?;
                    formula.kind = crate::cell::FormulaKind::Cse {
                        rows: block.end.row - block.start.row + 1,
                        cols: block.end.col - block.start.col + 1,
                    };
                }
                Some((key, formula))
            })
            .collect();
        macro_rules! shift {
            ($map:ident, $ty:ty) => {
                self.$map = std::mem::take(&mut self.$map)
                    .into_iter()
                    .filter_map(|(index, value)| {
                        edit.point(index.into()).map(|index| (index as $ty, value))
                    })
                    .collect();
            };
        }
        if edit.rows {
            shift!(row_heights, u32);
            shift!(hidden_rows, u32);
            shift!(row_outline_levels, u32);
            shift!(row_collapsed, u32);
        } else {
            shift!(column_widths, u16);
            shift!(hidden_columns, u16);
            shift!(column_outline_levels, u16);
            shift!(column_collapsed, u16);
        }
        self.merged_regions = self
            .merged_regions
            .iter()
            .filter_map(|range| edit.range(*range))
            .collect();
        self.spill_sources.clear();
        self.spill_saved_arrays();
        self.invalidate_bounds();
    }

    /// Spills each formula's saved array result again at the formula's cell, where its cells are
    /// free, so an array the engine doesn't calculate again (one reading another workbook) keeps
    /// its values after rows or columns move. Calculating replaces them.
    fn spill_saved_arrays(&mut self) {
        let arrays: Vec<((u32, u16), (u32, u16))> = self
            .formulas
            .iter()
            .filter_map(|(&cell, formula)| {
                let array = formula.array_result.as_ref()?;
                let (rows, cols) = (array.len() as u32, array.first().map_or(0, Vec::len) as u16);
                (rows * cols as u32 > 1).then_some((cell, (rows, cols)))
            })
            .collect();
        for ((row, col), (rows, cols)) in arrays {
            if row.checked_add(rows - 1).is_none_or(|end| end >= crate::MAX_ROWS)
                || col.checked_add(cols - 1).is_none_or(|end| end >= crate::MAX_COLS)
                || !self.can_spill_to(row, col, rows, cols)
            {
                continue;
            }
            self.register_spill_source(row, col, SpillInfo::new(rows, cols));
            for r in 0..rows {
                for c in 0..cols {
                    if r == 0 && c == 0 {
                        continue;
                    }
                    let target = CellValue::SpillTarget { source_row: row, source_col: col, offset_row: r, offset_col: c };
                    let style = self.get(row + r, col + c).map_or(0, |cell| cell.style_index);
                    self.set(row + r, col + c, CellData::with_style(target, style));
                }
            }
        }
    }

    /// Rewrites the references in each formula's text. Arrays keep their values and spills until
    /// they are calculated again.
    pub(crate) fn rewrite_references(&mut self, context: &str, change: &ReferenceEdit<'_>) {
        for formula in self.formulas.values_mut() {
            formula.text = rewrite_references(&formula.text, context, change);
        }
    }
}
