use duke_sheets_chart::{CellMarker, Chart, ChartType, DrawingAnchor};
use duke_sheets_core::named_range::NameScope;
use duke_sheets_core::*;

fn workbook() -> Workbook {
    let mut wb = Workbook::new();
    wb.rename_worksheet(0, "Old Name").unwrap();
    wb.add_worksheet_with_name("Other").unwrap();
    wb
}

#[test]
fn insert_rows_moves_cells_and_all_attached_metadata() {
    let mut wb = workbook();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("B3", 42).unwrap();
    let mut style = Style::default();
    style.font.bold = true;
    ws.set_cell_style("B3", &style).unwrap();
    ws.set_cell_formula("C3", "=$B$3+B3").unwrap();
    ws.set_row_height(2, 33.0);
    ws.set_comment("B3", CellComment::new("Ann", "note"))
        .unwrap();
    ws.set_hyperlink(
        "B3",
        Hyperlink {
            target: "#'Old Name'!B3".into(),
            display: None,
            tooltip: None,
            location: None,
        },
    )
    .unwrap();
    ws.merge_cells(&CellRange::parse("D2:E4").unwrap()).unwrap();
    ws.add_data_validation(
        DataValidation::custom("B3>0").with_range(CellRange::parse("B3:B5").unwrap()),
    );
    ws.add_conditional_format(
        ConditionalFormatRule::expression("B3>1").with_range(CellRange::parse("B3:B5").unwrap()),
    );
    ws.add_chart(
        Chart::new(ChartType::ColumnClustered),
        DrawingAnchor::OneCell {
            from: CellMarker {
                row: 2,
                col: 1,
                ..Default::default()
            },
            width_emu: 100,
            height_emu: 100,
        },
    )
    .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("A1", "='Old Name'!$B$3+SUM('Old Name'!B2:B4)+B3+\"B3\"")
        .unwrap();
    wb.define_name("Total", "'Old Name'!$B$3").unwrap();
    wb.worksheet_mut(0).unwrap().add_data_validation(
        DataValidation::custom("A1>0").with_range(CellRange::parse("A1:A1048576").unwrap()),
    );
    wb.insert_rows(0, 2, 2).unwrap();
    let ws = wb.worksheet(0).unwrap();
    assert_eq!(
        ws.cell("B5").unwrap().unwrap().value,
        CellValue::Number(42.0)
    );
    assert!(ws.cell_style("B5").unwrap().unwrap().font.bold);
    assert_eq!(ws.get_formula_at(4, 2), Some("=$B$5+B5"));
    assert_eq!(ws.row_height(4), 33.0);
    assert_eq!(
        ws.data_validations()[1].ranges[0],
        CellRange::parse("A1:A1048576").unwrap()
    );
    assert_eq!(ws.comment("B5").unwrap().unwrap().plain_text(), "note");
    assert_eq!(ws.hyperlink("B5").unwrap().target, "#'Old Name'!B5");
    assert_eq!(ws.merged_regions()[0], CellRange::parse("D2:E6").unwrap());
    assert_eq!(
        ws.data_validations()[0].ranges[0],
        CellRange::parse("B5:B7").unwrap()
    );
    assert!(
        matches!(&ws.data_validations()[0].validation_type, ValidationType::Custom { formula } if formula == "B5>0")
    );
    assert_eq!(
        ws.conditional_formats()[0].ranges[0],
        CellRange::parse("B5:B7").unwrap()
    );
    assert!(
        matches!(&ws.drawings().last().unwrap().anchor, DrawingAnchor::OneCell { from, .. } if from.row == 4)
    );
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("='Old Name'!$B$5+SUM('Old Name'!B2:B6)+B3+\"B3\"")
    );
    assert_eq!(
        wb.get_named_range("Total", 1).unwrap().refers_to,
        "'Old Name'!$B$5"
    );
}

#[test]
fn delete_rows_shrinks_ranges_and_invalidates_deleted_references() {
    let mut wb = workbook();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("A5", 5).unwrap();
    ws.set_comment("B3", CellComment::text_only("removed"))
        .unwrap();
    ws.merge_cells(&CellRange::parse("D2:E5").unwrap()).unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula(
            "A1",
            "='Old Name'!B3+SUM('Old Name'!$B$2:B5)+SUM('Old Name'!2:5)+SUM('Old Name'!3:4)",
        )
        .unwrap();
    wb.delete_rows(0, 2, 2).unwrap();
    assert_eq!(
        wb.worksheet(0).unwrap().cell("A3").unwrap().unwrap().value,
        CellValue::Number(5.0)
    );
    assert_eq!(wb.worksheet(0).unwrap().comment_count(), 0);
    assert_eq!(
        wb.worksheet(0).unwrap().merged_regions()[0],
        CellRange::parse("D2:E3").unwrap()
    );
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("=#REF!+SUM('Old Name'!$B$2:B3)+SUM('Old Name'!2:3)+SUM(#REF!)")
    );
}

#[test]
fn insert_and_delete_columns_preserve_width_and_rewrite_whole_columns() {
    let mut wb = workbook();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_cell_value("C2", 7).unwrap();
    ws.set_column_width(2, 27.0);
    ws.set_comment("C2", CellComment::text_only("moves"))
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula(
            "A1",
            "=SUM('Old Name'!$B:$D)+'Old Name'!C2+LOG10(100)+Table1[C2]+'[Book.xlsx]Old Name'!C2",
        )
        .unwrap();
    wb.insert_columns(0, 2, 2).unwrap();
    assert_eq!(wb.worksheet(0).unwrap().column_width(4), 27.0);
    assert!(wb.worksheet(0).unwrap().comment("E2").unwrap().is_some());
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("=SUM('Old Name'!$B:$F)+'Old Name'!E2+LOG10(100)+Table1[C2]+'[Book.xlsx]Old Name'!C2")
    );
    wb.delete_columns(0, 4, 1).unwrap();
    assert_eq!(wb.worksheet(0).unwrap().comment_count(), 0);
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("=SUM('Old Name'!$B:$E)+#REF!+LOG10(100)+Table1[C2]+'[Book.xlsx]Old Name'!C2")
    );
}

#[test]
fn rename_rewrites_quoted_references_and_names_but_not_strings_or_external_links() {
    let mut wb = workbook();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("A1", "='old name'!A1+\"'Old Name'!A1\"+'[book]Old Name'!A1")
        .unwrap();
    wb.define_name("Total", "'Old Name'!$A$1").unwrap();
    wb.rename_worksheet(0, "O'Brien").unwrap();
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("='O''Brien'!A1+\"'Old Name'!A1\"+'[book]Old Name'!A1")
    );
    assert_eq!(
        wb.get_named_range("Total", 1).unwrap().refers_to,
        "'O''Brien'!$A$1"
    );
    assert!(wb.rename_worksheet(0, "Other").is_err());
}

#[test]
fn reorder_and_remove_keep_local_name_scope_and_active_sheet() {
    let mut wb = workbook();
    wb.define_name_for_sheet("Local", "'Old Name'!A1", 0)
        .unwrap();
    wb.move_worksheet(0, 1).unwrap();
    assert_eq!(wb.active_sheet(), 1);
    assert_eq!(wb.worksheet(1).unwrap().name(), "Old Name");
    assert!(wb.get_named_range("Local", 0).is_none());
    assert_eq!(
        wb.get_named_range("Local", 1).unwrap().scope,
        NameScope::Sheet(1)
    );
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_formula("A1", "='Old Name'!A1")
        .unwrap();
    wb.remove_worksheet(1).unwrap();
    assert!(wb.get_named_range("Local", 0).is_none());
    assert_eq!(
        wb.worksheet(0).unwrap().get_formula_at(0, 0),
        Some("=#REF!")
    );
}

#[test]
fn invalid_edits_are_atomic() {
    let mut wb = workbook();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("A1048576", 1)
        .unwrap();
    assert!(wb.insert_rows(0, 0, 1).is_err());
    assert!(wb.delete_rows(0, 0, 0).is_err());
    assert!(wb.delete_columns(0, 16383, 2).is_err());
    assert!(wb.insert_rows(50, 0, 1).is_err());
    assert_eq!(
        wb.worksheet(0)
            .unwrap()
            .cell("A1048576")
            .unwrap()
            .unwrap()
            .value,
        CellValue::Number(1.0)
    );
    let mut dimensions = Workbook::new();
    dimensions
        .worksheet_mut(0)
        .unwrap()
        .set_row_height(MAX_ROWS - 1, 30.0);
    assert!(dimensions.insert_rows(0, 0, 1).is_err());
    assert_eq!(
        dimensions.worksheet(0).unwrap().row_height(MAX_ROWS - 1),
        30.0
    );
}

#[test]
fn workbook_name_retains_its_definition_sheet_after_structural_edits() {
    let mut wb = workbook();
    wb.set_active_sheet(1).unwrap();
    wb.define_name("Total", "B3").unwrap();
    wb.define_name("Calculation", "=SUM(B3:B4)+42+Other!A1+\"B3\"")
        .unwrap();
    wb.set_active_sheet(0).unwrap();
    wb.insert_rows(1, 2, 1).unwrap();
    assert_eq!(
        wb.get_named_range("Total", 0).unwrap().refers_to,
        "'Other'!B4"
    );
    assert_eq!(
        wb.get_named_range("Calculation", 0).unwrap().refers_to,
        "=SUM('Other'!B4:B5)+42+Other!A1+\"B3\""
    );
    wb.insert_rows(0, 2, 1).unwrap();
    assert_eq!(
        wb.get_named_range("Total", 0).unwrap().refers_to,
        "'Other'!B4"
    );
}

#[test]
fn ranges_with_two_sheet_qualified_endpoints_move_together() {
    let mut wb = workbook();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("A1", "=SUM('Old Name'!$A$1:'Old Name'!A3)")
        .unwrap();
    wb.insert_rows(0, 0, 1).unwrap();
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("=SUM('Old Name'!$A$2:'Old Name'!A4)")
    );
    wb.rename_worksheet(0, "New Name").unwrap();
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("=SUM('New Name'!$A$2:'New Name'!A4)")
    );
    wb.delete_rows(0, 1, 1).unwrap();
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("=SUM('New Name'!$A$2:'New Name'!A3)")
    );
}

#[test]
fn whole_grid_formula_ranges_keep_their_terminal_edge() {
    let mut wb = workbook();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("A1", "=SUM('Old Name'!A:XFD)+SUM('Old Name'!1:1048576)")
        .unwrap();
    wb.insert_columns(0, 1, 1).unwrap();
    wb.insert_rows(0, 1, 1).unwrap();
    assert_eq!(
        wb.worksheet(1).unwrap().get_formula_at(0, 0),
        Some("=SUM('Old Name'!A:XFD)+SUM('Old Name'!1:1048576)")
    );
}

#[test]
fn print_area_and_repeating_headings_follow_structural_edits() {
    let mut wb = workbook();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_print_area(CellRange::parse("A1:D10").unwrap());
    ws.set_repeat_rows(0, 2);
    ws.set_repeat_cols(0, 2);
    wb.insert_rows(0, 1, 1).unwrap();
    wb.insert_columns(0, 1, 1).unwrap();
    let ws = wb.worksheet(0).unwrap();
    assert_eq!(ws.print_area(), Some(&CellRange::parse("A1:E11").unwrap()));
    assert_eq!(ws.repeat_rows(), Some((0, 3)));
    assert_eq!(ws.repeat_cols(), Some((0, 3)));
    wb.delete_rows(0, 0, 4).unwrap();
    wb.delete_columns(0, 0, 4).unwrap();
    let ws = wb.worksheet(0).unwrap();
    assert_eq!(ws.print_area(), Some(&CellRange::parse("A1:A7").unwrap()));
    assert_eq!(ws.repeat_rows(), None);
    assert_eq!(ws.repeat_cols(), None);
}

#[test]
fn chart_series_name_references_follow_sheet_rename_and_insertions() {
    use duke_sheets_chart::{DataReference, DataSeries};
    let mut wb = workbook();
    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(
        DataSeries::new(DataReference::formula("'Old Name'!B2:B4")).with_name("'Old Name'!$B$1"),
    );
    chart.add_series(DataSeries::new(DataReference::numbers(vec![1.0])).with_name("B1"));
    wb.worksheet_mut(1)
        .unwrap()
        .add_chart(chart, DrawingAnchor::default())
        .unwrap();
    wb.rename_worksheet(0, "New Name").unwrap();
    wb.insert_rows(0, 0, 1).unwrap();
    let chart = wb.worksheet(1).unwrap().charts().next().unwrap();
    assert_eq!(
        chart.payload.series[0].name.as_deref(),
        Some("'New Name'!$B$2")
    );
    assert_eq!(chart.payload.series[1].name.as_deref(), Some("B1"));
}

fn drawing_bounds(ws: &Worksheet, anchor: &DrawingAnchor) -> (i128, i128, i128, i128) {
    let DrawingAnchor::TwoCell { from, to, .. } = anchor.to_two_cell_with_metrics(ws) else {
        unreachable!()
    };
    let (x, y) = duke_sheets_chart::marker_position_emu(&from, ws);
    let (right, bottom) = duke_sheets_chart::marker_position_emu(&to, ws);
    (x, y, right - x, bottom - y)
}

#[test]
fn drawing_placement_modes_preserve_size_and_absolute_position() {
    use duke_sheets_chart::EditAs;
    let mut wb = workbook();
    let ws = wb.worksheet_mut(0).unwrap();
    ws.set_row_height(0, 40.0);
    ws.set_row_height(3, 35.0);
    ws.set_column_width(0, 25.0);
    for edit_as in [EditAs::OneCell, EditAs::Absolute] {
        ws.add_chart(
            Chart::new(ChartType::ColumnClustered),
            DrawingAnchor::TwoCell {
                from: CellMarker {
                    row: 2,
                    col: 2,
                    row_offset_emu: 1234,
                    col_offset_emu: 2345,
                },
                to: CellMarker {
                    row: 5,
                    col: 5,
                    row_offset_emu: 4321,
                    col_offset_emu: 5432,
                },
                edit_as: Some(edit_as),
            },
        )
        .unwrap();
    }
    let before: Vec<_> = ws
        .drawings()
        .iter()
        .map(|d| drawing_bounds(ws, &d.anchor))
        .collect();
    wb.insert_rows(0, 3, 2).unwrap();
    wb.insert_columns(0, 3, 2).unwrap();
    let ws = wb.worksheet(0).unwrap();
    assert_eq!(drawing_bounds(ws, &ws.drawings()[0].anchor), before[0]);
    assert_eq!(drawing_bounds(ws, &ws.drawings()[1].anchor), before[1]);
    wb.insert_rows(0, 0, 1).unwrap();
    wb.insert_columns(0, 0, 1).unwrap();
    let ws = wb.worksheet(0).unwrap();
    let moved = drawing_bounds(ws, &ws.drawings()[0].anchor);
    assert!(moved.0 > before[0].0 && moved.1 > before[0].1);
    assert_eq!((moved.2, moved.3), (before[0].2, before[0].3));
    assert_eq!(drawing_bounds(ws, &ws.drawings()[1].anchor), before[1]);
}

#[test]
fn insertion_does_not_freeze_a_previously_unfrozen_axis() {
    let mut wb = workbook();
    wb.worksheet_mut(0).unwrap().set_freeze_panes(0, 1);
    wb.insert_rows(0, 0, 1).unwrap();
    assert_eq!(wb.worksheet(0).unwrap().freeze_panes().unwrap().row, 0);
    wb.worksheet_mut(0).unwrap().set_freeze_panes(1, 0);
    wb.insert_columns(0, 0, 1).unwrap();
    assert_eq!(wb.worksheet(0).unwrap().freeze_panes().unwrap().col, 0);
}
