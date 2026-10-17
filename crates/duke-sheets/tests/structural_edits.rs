use duke_sheets::prelude::*;
use duke_sheets_chart::{
    CellMarker, Chart, ChartType, DataReference, DataSeries, DrawingAnchor, EmbeddedImage,
    ImageFormat,
};
use duke_sheets_core::{CellRange, ConditionalFormatRule, DataValidation, Hyperlink, Style};
use std::io::Cursor;

#[test]
fn structural_edits_survive_xlsx_roundtrip() {
    let mut wb = Workbook::new();
    wb.add_worksheet_with_name("Other").unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_cell_value("B3", 42)
        .unwrap();
    wb.worksheet_mut(0)
        .unwrap()
        .set_comment("B3", duke_sheets_core::CellComment::text_only("note"))
        .unwrap();
    wb.worksheet_mut(1)
        .unwrap()
        .set_cell_formula("A1", "=Sheet1!B3")
        .unwrap();
    wb.define_name("Total", "Sheet1!$B$3").unwrap();
    let sheet = wb.worksheet_mut(0).unwrap();
    sheet.set_print_area(CellRange::parse("A1:D10").unwrap());
    sheet.set_repeat_rows(0, 2);
    sheet.set_repeat_cols(0, 2);
    let mut style = Style::default();
    style.font.bold = true;
    sheet.set_cell_style("B3", &style).unwrap();
    sheet.set_row_height(2, 30.0);
    sheet.set_column_width(1, 22.0);
    sheet
        .merge_cells(&CellRange::parse("D2:E4").unwrap())
        .unwrap();
    sheet
        .set_hyperlink(
            "B3",
            Hyperlink {
                target: String::new(),
                location: Some("Sheet1!B3".into()),
                tooltip: None,
                display: None,
            },
        )
        .unwrap();
    sheet.add_data_validation(
        DataValidation::custom("B3>0").with_range(CellRange::parse("B3:B5").unwrap()),
    );
    sheet.add_conditional_format(
        ConditionalFormatRule::expression("B3>0").with_range(CellRange::parse("B3:B5").unwrap()),
    );
    let anchor = DrawingAnchor::TwoCell {
        from: CellMarker {
            row: 2,
            col: 1,
            ..Default::default()
        },
        to: CellMarker {
            row: 4,
            col: 3,
            ..Default::default()
        },
        edit_as: None,
    };
    sheet
        .add_image(
            EmbeddedImage {
                format: ImageFormat::Png,
                media_path: String::new(),
                svg_media_path: None,
                width_emu: 1_000_000,
                height_emu: 2_000_000,
                rotation: None,
                flip_h: false,
                flip_v: false,
                data: TEST_PNG_1X1.to_vec(),
                svg_data: None,
            },
            anchor.clone(),
        )
        .unwrap();
    let mut chart = Chart::new(ChartType::ColumnClustered);
    chart.add_series(DataSeries::new(DataReference::formula("Sheet1!$B$3:$B$5")));
    sheet.add_chart(chart, anchor).unwrap();
    let mut initial_bytes = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut initial_bytes)).unwrap();
    let mut wb = XlsxReader::read(Cursor::new(initial_bytes)).unwrap();
    wb.insert_rows(0, 1, 2).unwrap();
    wb.insert_columns(0, 1, 2).unwrap();
    wb.delete_rows(0, 0, 1).unwrap();
    wb.delete_columns(0, 0, 1).unwrap();
    wb.rename_worksheet(0, "Renamed Sheet").unwrap();
    wb.move_worksheet(0, 1).unwrap();
    let mut bytes = Vec::new();
    XlsxWriter::write(&wb, Cursor::new(&mut bytes)).unwrap();
    let loaded = XlsxReader::read(Cursor::new(&bytes)).unwrap();
    assert_eq!(loaded.worksheet(1).unwrap().name(), "Renamed Sheet");
    assert_eq!(
        loaded
            .worksheet(1)
            .unwrap()
            .cell("C4")
            .unwrap()
            .unwrap()
            .value,
        CellValue::Number(42.0)
    );
    assert_eq!(
        loaded
            .worksheet(0)
            .unwrap()
            .get_formula_at(0, 0)
            .unwrap()
            .trim_start_matches('='),
        "'Renamed Sheet'!C4"
    );
    assert_eq!(
        loaded.get_named_range("Total", 0).unwrap().refers_to,
        "'Renamed Sheet'!$C$4"
    );
    assert_eq!(
        loaded
            .worksheet(1)
            .unwrap()
            .comment("C4")
            .unwrap()
            .unwrap()
            .plain_text(),
        "note"
    );
    let sheet = loaded.worksheet(1).unwrap();
    assert_eq!(
        sheet.print_area(),
        Some(&CellRange::parse("A1:E11").unwrap())
    );
    assert_eq!(sheet.repeat_rows(), Some((0, 3)));
    assert_eq!(sheet.repeat_cols(), Some((0, 3)));
    assert!(sheet.cell_style("C4").unwrap().unwrap().font.bold);
    assert_eq!(sheet.row_height(3), 30.0);
    assert_eq!(sheet.column_width(2), 22.0);
    assert_eq!(
        sheet.merged_regions()[0],
        CellRange::parse("E3:F5").unwrap()
    );
    assert_eq!(
        sheet.hyperlink("C4").unwrap().location.as_deref(),
        Some("'Renamed Sheet'!C4")
    );
    assert_eq!(
        sheet.data_validations()[0].ranges[0],
        CellRange::parse("C4:C6").unwrap()
    );
    assert_eq!(
        sheet.conditional_formats()[0].ranges[0],
        CellRange::parse("C4:C6").unwrap()
    );
    let image = sheet.images().next().unwrap();
    assert_eq!(image.payload.data(), TEST_PNG_1X1);
    assert!(
        matches!(&image.object.unwrap().anchor, DrawingAnchor::TwoCell { from, to, .. } if from.row == 3 && from.col == 2 && to.row == 5 && to.col == 4)
    );
    let chart = sheet.charts().next().unwrap();
    assert!(
        matches!(&chart.payload.series[0].values, DataReference::Formula(formula) if formula == "'Renamed Sheet'!$C$4:$C$6")
    );
}

#[test]
fn custom_error_bar_references_follow_structural_edits_after_loading() {
    use duke_sheets_chart::{ErrorBarDirection, ErrorBarType, ErrorBars, ErrorValueType};

    let mut workbook = Workbook::new();
    let mut chart = Chart::new(ChartType::ColumnClustered);
    let mut series = DataSeries::new(DataReference::formula("Sheet1!$B$3:$B$4"));
    series.error_bars = Some(ErrorBars {
        direction: ErrorBarDirection::Y,
        bar_type: ErrorBarType::Both,
        value_type: ErrorValueType::Custom,
        value: None,
        no_end_cap: None,
        plus: Some(DataReference::formula("Sheet1!$C$3:$C$4")),
        minus: Some(DataReference::formula("Sheet1!$D$3:$D$4")),
    });
    chart.add_series(series);
    workbook.worksheet_mut(0).unwrap().add_chart(chart, DrawingAnchor::TwoCell {
        from: CellMarker::default(),
        to: CellMarker { row: 10, col: 5, ..Default::default() },
        edit_as: None,
    }).unwrap();
    let mut bytes = Vec::new();
    XlsxWriter::write(&workbook, Cursor::new(&mut bytes)).unwrap();
    let mut workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
    workbook.insert_rows(0, 1, 2).unwrap();
    workbook.rename_worksheet(0, "Renamed Sheet").unwrap();
    let mut bytes = Vec::new();
    XlsxWriter::write(&workbook, Cursor::new(&mut bytes)).unwrap();
    let workbook = XlsxReader::read(Cursor::new(bytes)).unwrap();
    let chart = workbook.worksheet(0).unwrap().charts().next().unwrap();
    let errors = chart.payload.series[0].error_bars.as_ref().unwrap();
    assert_eq!(errors.plus, Some(DataReference::formula("'Renamed Sheet'!$C$5:$C$6")));
    assert_eq!(errors.minus, Some(DataReference::formula("'Renamed Sheet'!$D$5:$D$6")));
}

const TEST_PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR length + name
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // width=1 height=1
    0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, // 8-bit RGBA + IHDR CRC
    0x89, 0x00, 0x00, 0x00, 0x0B, 0x49, 0x44, 0x41, // IDAT length + name
    0x54, 0x78, 0x9C, 0x63, 0x60, 0x00, 0x02, 0x00, // zlib-deflated 1 pixel
    0x00, 0x05, 0x00, 0x01, 0x7A, 0x5E, 0xAB, 0x3F, // IDAT CRC
    0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND length + name
    0xAE, 0x42, 0x60, 0x82, // IEND CRC
];
