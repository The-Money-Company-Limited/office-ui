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
