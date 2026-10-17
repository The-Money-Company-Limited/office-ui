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
        self.formulas = std::mem::take(&mut self.formulas)
            .into_iter()
            .filter_map(|((row, col), formula)| edit.cell(row, col).map(|key| (key, formula)))
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
        self.invalidate_bounds();
    }

    pub(crate) fn rewrite_references(&mut self, context: &str, change: &ReferenceEdit<'_>) {
        for formula in self.formulas.values_mut() {
            formula.text = rewrite_references(&formula.text, context, change);
            formula.array_result = None;
        }
        self.cells.retain(|_, cell| {
            if matches!(cell.value, CellValue::SpillTarget { .. }) {
                cell.value = CellValue::Empty;
                return cell.style_index != 0;
            }
            true
        });
        self.spill_sources.clear();
        self.invalidate_bounds();
    }
}
