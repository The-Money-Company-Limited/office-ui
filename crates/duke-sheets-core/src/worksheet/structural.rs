use super::*;
use crate::conditional_format::{CfRuleType, CfValueType};
use crate::structural::{rewrite_references, AxisEdit, ReferenceEdit};
use crate::validation::ValidationType;
use duke_sheets_chart::{DataReference, EditAs};

impl Worksheet {
    pub(crate) fn edit_axis(&mut self, edit: AxisEdit) {
        self.cells.edit_axis(edit);
        self.hyperlinks = std::mem::take(&mut self.hyperlinks)
            .into_iter()
            .filter_map(|(mut addr, link)| {
                let (row, col) = edit.cell(addr.row, addr.col)?;
                addr.row = row;
                addr.col = col;
                Some((addr, link))
            })
            .collect();
        self.data_validations.retain_mut(|rule| {
            rule.ranges = rule
                .ranges
                .iter()
                .filter_map(|range| edit.range(*range))
                .collect();
            !rule.ranges.is_empty()
        });
        self.conditional_formats.retain_mut(|rule| {
            rule.ranges = rule
                .ranges
                .iter()
                .filter_map(|range| edit.range(*range))
                .collect();
            !rule.ranges.is_empty()
        });
        self.protected_ranges.retain_mut(|rule| {
            rule.ranges = rule
                .ranges
                .iter()
                .filter_map(|range| edit.range(*range))
                .collect();
            !rule.ranges.is_empty()
        });
        self.tables.retain_mut(|table| {
            let Some(range) = edit.range(table.reference) else {
                return false;
            };
            if !edit.rows {
                let old_start = table.reference.start.col;
                let mut columns = Vec::new();
                let mut next_id = table.columns.iter().map(|c| c.id).max().unwrap_or(0) + 1;
                for col in range.start.col..=range.end.col {
                    let old = if edit.delete {
                        Some(if u32::from(col) >= edit.index {
                            u32::from(col) + edit.count
                        } else {
                            col.into()
                        })
                    } else if u32::from(col) < edit.index {
                        Some(col.into())
                    } else if u32::from(col) >= edit.index + edit.count {
                        Some(u32::from(col) - edit.count)
                    } else {
                        None
                    };
                    let existing = old
                        .and_then(|n| n.checked_sub(old_start.into()))
                        .and_then(|n| table.columns.get(n as usize))
                        .cloned();
                    columns.push(existing.unwrap_or_else(|| {
                        let id = next_id;
                        next_id += 1;
                        let mut name = format!("Column{id}");
                        while table.columns.iter().any(|c| c.name == name) {
                            name.push('_');
                        }
                        crate::table::TableColumn::new(id, name)
                    }));
                }
                table.columns = columns;
            }
            table.reference = range;
            true
        });
        if let Some(filter) = &mut self.auto_filter {
            if let Some(range) = edit.range(filter.range) {
                if !edit.rows {
                    filter.filter_columns.retain_mut(|column| {
                        let Some(col) =
                            edit.point(u32::from(filter.range.start.col) + column.col_id)
                        else {
                            return false;
                        };
                        column.col_id = col - u32::from(range.start.col);
                        true
                    });
                }
                filter.range = range;
            } else {
                self.auto_filter = None;
            }
        }
        self.drawings.retain_mut(|drawing| {
            if let DrawingKind::Comment { row, col, .. } = &mut drawing.kind {
                let Some((r, c)) = edit.cell(*row, *col) else {
                    return false;
                };
                *row = r;
                *col = c;
            }
            match &mut drawing.anchor {
                DrawingAnchor::TwoCell { from, to, edit_as }
                    if *edit_as != Some(EditAs::Absolute) =>
                {
                    edit.anchor(&mut from.row, &mut from.col);
                    edit.anchor(&mut to.row, &mut to.col);
                }
                DrawingAnchor::OneCell { from, .. } => edit.anchor(&mut from.row, &mut from.col),
                _ => {}
            }
            true
        });
        if let Some(panes) = &mut self.freeze_panes {
            edit.anchor(&mut panes.row, &mut panes.col);
        }
        if let Some(panes) = &mut self.split_panes {
            if let Some((row, col)) = &mut panes.top_left {
                edit.anchor(row, col);
            }
        }
        let breaks = if edit.rows {
            &mut self.row_breaks
        } else {
            &mut self.col_breaks
        };
        breaks.retain_mut(|br| {
            if let Some(id) = edit.point(br.id) {
                br.id = id;
                true
            } else {
                false
            }
        });
        self.clear_image_metadata();
    }

    pub(crate) fn rewrite_references(&mut self, change: &ReferenceEdit<'_>) {
        let context = self.name.clone();
        let rewrite = |text: &mut String| *text = rewrite_references(text, &context, change);
        self.cells.rewrite_references(&context, change);
        for link in self.hyperlinks.values_mut() {
            if let Some(location) = &mut link.location {
                rewrite(location);
            }
            if let Some(target) = link.target.strip_prefix('#') {
                link.target = format!("#{}", rewrite_references(target, &context, change));
            }
        }
        for rule in &mut self.data_validations {
            match &mut rule.validation_type {
                ValidationType::Whole { value1, value2, .. }
                | ValidationType::Decimal { value1, value2, .. }
                | ValidationType::Date { value1, value2, .. }
                | ValidationType::Time { value1, value2, .. }
                | ValidationType::TextLength { value1, value2, .. } => {
                    rewrite(value1);
                    if let Some(value) = value2 {
                        rewrite(value);
                    }
                }
                ValidationType::List { source } => rewrite(source),
                ValidationType::Custom { formula } => rewrite(formula),
                ValidationType::None => {}
            }
        }
        for rule in &mut self.conditional_formats {
            match &mut rule.rule_type {
                CfRuleType::CellIs {
                    formula1, formula2, ..
                } => {
                    rewrite(formula1);
                    if let Some(formula) = formula2 {
                        rewrite(formula);
                    }
                }
                CfRuleType::Expression { formula } => rewrite(formula),
                CfRuleType::ColorScale { colors } => {
                    for color in colors {
                        if color.value_type == CfValueType::Formula {
                            if let Some(value) = &mut color.value {
                                rewrite(value);
                            }
                        }
                    }
                }
                CfRuleType::DataBar {
                    min_value,
                    max_value,
                    ..
                } => {
                    for value in [min_value, max_value] {
                        if value.value_type == CfValueType::Formula {
                            if let Some(text) = &mut value.value {
                                rewrite(text);
                            }
                        }
                    }
                }
                CfRuleType::IconSet { values, .. } => {
                    for value in values {
                        if value.value_type == CfValueType::Formula {
                            if let Some(text) = &mut value.value {
                                rewrite(text);
                            }
                        }
                    }
                }
                _ => {}
            }
        }
        for table in &mut self.tables {
            for col in &mut table.columns {
                if let Some(formula) = &mut col.calculated_column_formula {
                    rewrite(formula);
                }
                if let Some(formula) = &mut col.totals_row_formula {
                    rewrite(formula);
                }
            }
        }
        for selection in &mut self.selections {
            if let Some(cell) = &mut selection.active_cell {
                rewrite(cell);
            }
            if let Some(range) = &mut selection.sqref {
                rewrite(range);
            }
        }
        for drawing in &mut self.drawings {
            rewrite_drawing(&mut drawing.kind, &rewrite);
        }
        self.mutation_count += 1;
        self.topology_generation += 1;
        self.dirty_value_ranges.clear();
        self.clear_image_metadata();
    }
}

pub(crate) fn rewrite_chart(chart: &mut Chart, rewrite: &impl Fn(&mut String)) {
    for series in &mut chart.series {
        if let DataReference::Formula(formula) = &mut series.values {
            rewrite(formula);
        }
        if let Some(DataReference::Formula(formula)) = &mut series.categories {
            rewrite(formula);
        }
    }
    for group in &mut chart.type_groups {
        for series in &mut group.series {
            if let DataReference::Formula(formula) = &mut series.values {
                rewrite(formula);
            }
            if let Some(DataReference::Formula(formula)) = &mut series.categories {
                rewrite(formula);
            }
        }
    }
}

fn rewrite_drawing(kind: &mut DrawingKind, rewrite: &impl Fn(&mut String)) {
    match kind {
        DrawingKind::Chart(chart) => rewrite_chart(chart, rewrite),
        DrawingKind::ChartEx(chart) => {
            use duke_sheets_chart::chart_ex::ChartExDimension;
            for data in &mut chart.data {
                for dimension in &mut data.dimensions {
                    let (formula, nf_formula) = match dimension {
                        ChartExDimension::String {
                            formula,
                            nf_formula,
                            ..
                        }
                        | ChartExDimension::Numeric {
                            formula,
                            nf_formula,
                            ..
                        } => (formula, nf_formula),
                    };
                    if let Some(formula) = formula {
                        rewrite(formula);
                    }
                    if let Some(formula) = nf_formula {
                        rewrite(formula);
                    }
                }
            }
        }
        DrawingKind::FormControl(control) => match &mut control.kind {
            FormControlKind::ListBox {
                input_range,
                cell_link,
                ..
            }
            | FormControlKind::Dropdown {
                input_range,
                cell_link,
                ..
            } => {
                if let Some(formula) = input_range {
                    rewrite(formula);
                }
                if let Some(formula) = cell_link {
                    rewrite(formula);
                }
            }
            FormControlKind::Checkbox { cell_link, .. }
            | FormControlKind::OptionButton { cell_link, .. }
            | FormControlKind::Scrollbar { cell_link, .. }
            | FormControlKind::Spinner { cell_link, .. } => {
                if let Some(formula) = cell_link {
                    rewrite(formula);
                }
            }
            _ => {}
        },
        DrawingKind::Group(group) => {
            for child in &mut group.children {
                rewrite_drawing(&mut child.kind, rewrite);
            }
        }
        _ => {}
    }
}
