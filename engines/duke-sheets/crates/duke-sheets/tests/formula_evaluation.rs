//! Tests for formula evaluation with cell references

use duke_sheets::prelude::*;
use duke_sheets::{evaluate, parse_formula, EvaluationContext, FormulaValue};
use duke_sheets_formula::EvalCache;
use std::path::PathBuf;

/// Test basic formula evaluation without cell references
#[test]
fn test_evaluate_simple_formulas() {
    // Arithmetic
    let ast = parse_formula("=1+2*3").unwrap();
    let ctx = EvaluationContext::simple();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(7.0));

    // String concatenation
    let ast = parse_formula("=\"Hello \"&\"World\"").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::String("Hello World".into()));

    // Comparison
    let ast = parse_formula("=5>3").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));
}

/// Test SUM function
#[test]
fn test_evaluate_sum() {
    let ast = parse_formula("=SUM(1,2,3,4,5)").unwrap();
    let ctx = EvaluationContext::simple();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(15.0));
}

/// Test IF function
#[test]
fn test_evaluate_if() {
    let ctx = EvaluationContext::simple();

    let ast = parse_formula("=IF(1>0,\"Yes\",\"No\")").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::String("Yes".into()));

    let ast = parse_formula("=IF(1<0,\"Yes\",\"No\")").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::String("No".into()));
}

/// Test formula evaluation with cell references
#[test]
fn test_evaluate_with_cell_references() {
    // Create a workbook with some data
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 10.0).unwrap();
    sheet.set_cell_value("A2", 20.0).unwrap();
    sheet.set_cell_value("A3", 30.0).unwrap();
    sheet.set_cell_value("B1", 5.0).unwrap();

    // Create evaluation context with workbook reference
    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // Test simple cell reference
    let ast = parse_formula("=A1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(10.0));

    // Test cell reference in arithmetic
    let ast = parse_formula("=A1+B1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(15.0));

    // Test cell reference in comparison
    let ast = parse_formula("=A1>B1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));
}

/// Test formula evaluation with range references
#[test]
fn test_evaluate_with_range_references() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 10.0).unwrap();
    sheet.set_cell_value("A2", 20.0).unwrap();
    sheet.set_cell_value("A3", 30.0).unwrap();

    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // Test SUM with range
    let ast = parse_formula("=SUM(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(60.0));

    // Test AVERAGE with range
    let ast = parse_formula("=AVERAGE(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(20.0));

    // Test MIN/MAX with range
    let ast = parse_formula("=MIN(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(10.0));

    let ast = parse_formula("=MAX(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(30.0));
}

/// Test complex nested formulas
#[test]
fn test_evaluate_complex_formulas() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 100.0).unwrap();
    sheet.set_cell_value("A2", 50.0).unwrap();
    sheet.set_cell_value("B1", 0.1).unwrap(); // 10%

    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // Calculate: IF A1 > A2, calculate 10% of A1, else calculate 10% of A2
    let ast = parse_formula("=IF(A1>A2,A1*B1,A2*B1)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(10.0));

    // Nested SUM and multiplication
    let ast = parse_formula("=SUM(A1,A2)*B1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(15.0));
}

/// Test error propagation in formulas
#[test]
fn test_error_propagation() {
    let ctx = EvaluationContext::simple();

    // Division by zero
    let ast = parse_formula("=1/0").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert!(matches!(result, FormulaValue::Error(_)));

    // Error in arithmetic propagates
    let ast = parse_formula("=1/0+5").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert!(matches!(result, FormulaValue::Error(_)));
}

/// Test empty cell handling
#[test]
fn test_empty_cell_handling() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", 10.0).unwrap();
    // A2 is empty
    sheet.set_cell_value("A3", 30.0).unwrap();

    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // Empty cells are treated as 0 in arithmetic
    let ast = parse_formula("=A1+A2").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(10.0)); // 10 + 0

    // SUM ignores empty cells
    let ast = parse_formula("=SUM(A1:A3)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Number(40.0)); // 10 + 0 + 30
}

/// Test string operations
#[test]
fn test_string_operations() {
    let mut wb = Workbook::new();
    let sheet = wb.worksheet_mut(0).unwrap();

    sheet.set_cell_value("A1", "Hello").unwrap();
    sheet.set_cell_value("B1", "World").unwrap();

    let ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);

    // String concatenation with cell references
    let ast = parse_formula("=A1&\" \"&B1").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::String("Hello World".into()));
}

/// Test boolean functions
#[test]
fn test_boolean_functions() {
    let ctx = EvaluationContext::simple();

    // AND function
    let ast = parse_formula("=AND(TRUE,TRUE,TRUE)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));

    let ast = parse_formula("=AND(TRUE,FALSE,TRUE)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(false));

    // OR function
    let ast = parse_formula("=OR(FALSE,FALSE,TRUE)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));

    // NOT function
    let ast = parse_formula("=NOT(FALSE)").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));

    // Combined
    let ast = parse_formula("=AND(NOT(FALSE),OR(TRUE,FALSE))").unwrap();
    let result = evaluate(&ast, &ctx).unwrap();
    assert_eq!(result, FormulaValue::Boolean(true));
}

#[test]
fn test_workbook_calculate_power_zero_zero() {
    let mut wb = Workbook::new();
    {
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_formula("A1", "=POWER(0,0)").unwrap();
    }

    let cache = EvalCache::new([(String::from("Sheet1"), 0usize)].into_iter().collect());
    let mut ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);
    ctx.eval_cache = Some(&cache);
    let ast = parse_formula(wb.worksheet(0).unwrap().get_formula_at(0, 0).unwrap()).unwrap();
    // Excel returns #NUM! for POWER(0,0)
    assert_eq!(evaluate(&ast, &ctx).unwrap(), FormulaValue::Error(CellError::Num));

    wb.calculate().unwrap();

    let sheet = wb.worksheet(0).unwrap();
    assert_eq!(sheet.get_value_at(0, 0), CellValue::Error(CellError::Num));
}

#[test]
fn test_workbook_calculate_index_zero_zero() {
    let mut wb = Workbook::new();
    {
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet
            .set_cell_formula("A1", "=INDEX({1,2,3;4,5,6},0,0)")
            .unwrap();
    }

    let cache = EvalCache::new([(String::from("Sheet1"), 0usize)].into_iter().collect());
    let mut ctx = EvaluationContext::new(Some(&wb), 0, 0, 0);
    ctx.eval_cache = Some(&cache);
    let ast = parse_formula(wb.worksheet(0).unwrap().get_formula_at(0, 0).unwrap()).unwrap();
    // INDEX(area, 0, 0) is the whole area, which spills from A1
    assert!(matches!(evaluate(&ast, &ctx).unwrap(), FormulaValue::Array { .. }));

    wb.calculate().unwrap();

    let sheet = wb.worksheet(0).unwrap();
    assert_eq!(sheet.get_value_at(0, 0), CellValue::Number(1.0));
    assert_eq!(sheet.get_value_at(1, 2), CellValue::Number(6.0));
}

#[test]
fn test_loaded_parity_cells_recalculate() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join("data/formula-parity.xlsx");
    if !fixture_path.exists() {
        return;
    }

    let mut wb = XlsxReader::read_file(&fixture_path).unwrap();
    let tests_sheet_idx = wb.sheet_index("Tests").unwrap();
    let tests_sheet = wb.worksheet(tests_sheet_idx).unwrap();

    assert_eq!(
        tests_sheet.get_formula_at(441, 2),
        Some("=INDEX({1,2,3;4,5,6},0,0)")
    );
    assert_eq!(tests_sheet.get_formula_at(456, 2), Some("=POWER(0,0)"));

    let index_ast = parse_formula(tests_sheet.get_formula_at(441, 2).unwrap()).unwrap();
    let power_ast = parse_formula(tests_sheet.get_formula_at(456, 2).unwrap()).unwrap();

    let cache = EvalCache::new(
        [(String::from("Tests"), tests_sheet_idx)]
            .into_iter()
            .collect(),
    );
    let mut index_ctx = EvaluationContext::new(Some(&wb), tests_sheet_idx, 441, 2);
    let mut power_ctx = EvaluationContext::new(Some(&wb), tests_sheet_idx, 456, 2);
    index_ctx.eval_cache = Some(&cache);
    power_ctx.eval_cache = Some(&cache);

    assert_eq!(
        evaluate(&index_ast, &index_ctx).unwrap(),
        FormulaValue::Number(1.0) // Excel returns first element for INDEX(2D,0,0)
    );
    assert_eq!(
        evaluate(&power_ast, &power_ctx).unwrap(),
        FormulaValue::Error(CellError::Num) // Excel returns #NUM! for POWER(0,0)
    );

    wb.calculate().unwrap();

    let tests_sheet = wb.worksheet(tests_sheet_idx).unwrap();
    assert_eq!(
        tests_sheet.get_value_at(441, 2),
        CellValue::Number(1.0)
    );
    assert_eq!(tests_sheet.get_value_at(456, 2), CellValue::Error(CellError::Num));
}

/// A range with a fixed start and a moving end grows as it is copied down (COD-1770).
#[test]
fn test_workbook_calculate_growing_ranges() {
    let mut wb = Workbook::new();
    {
        let sheet = wb.worksheet_mut(0).unwrap();
        sheet.set_cell_value("A1", "header").unwrap();
        for row in 2..=6 {
            sheet.set_cell_value(&format!("A{row}"), (row - 1) as f64).unwrap();
            sheet.set_cell_formula(&format!("B{row}"), &format!("=SUM($A$2:A{row})")).unwrap();
            sheet.set_cell_formula(&format!("C{row}"), &format!("=COUNT($A$2:A{row})")).unwrap();
            sheet.set_cell_formula(&format!("D{row}"), &format!("=ROWS($A$2:A{row})")).unwrap();
            sheet.set_cell_formula(&format!("E{row}"), &format!("=COUNTIF(A$2:A{row},\">1\")")).unwrap();
        }
    }
    wb.calculate().unwrap();

    let sheet = wb.worksheet(0).unwrap();
    for row in 2..=6u32 {
        let n = (row - 1) as f64;
        let r = row - 1;
        assert_eq!(sheet.get_value_at(r, 1), CellValue::Number(n * (n + 1.0) / 2.0), "SUM in row {row}");
        assert_eq!(sheet.get_value_at(r, 2), CellValue::Number(n), "COUNT in row {row}");
        assert_eq!(sheet.get_value_at(r, 3), CellValue::Number(n), "ROWS in row {row}");
        assert_eq!(sheet.get_value_at(r, 4), CellValue::Number(n - 1.0), "COUNTIF in row {row}");
    }
}

/// Whole columns and rows read the sheet's used cells, on this sheet and others (COD-1772).
#[test]
fn test_workbook_calculate_whole_columns_and_rows() {
    let mut wb = Workbook::new();
    {
        let sheet = wb.worksheet_mut(0).unwrap();
        for (row, (key, value)) in [("x", 1.0), ("y", 2.0), ("x", 3.0)].into_iter().enumerate() {
            sheet.set_cell_value(&format!("A{}", row + 1), key).unwrap();
            sheet.set_cell_value(&format!("B{}", row + 1), value).unwrap();
        }
        sheet.set_cell_value("F10", 5.0).unwrap();
        sheet.set_cell_value("G10", 7.0).unwrap();
        let formulas = [
            ("D1", "=SUM(B:B)"),
            ("D2", "=COUNTA(A:A)"),
            ("D3", "=SUMIF(A:A,\"x\",B:B)"),
            ("D4", "=COUNTIFS($A:$A,$A1,$B:$B,\">1\")"),
            ("D5", "=VLOOKUP(\"y\",A:B,2,FALSE)"),
            ("D6", "=MATCH(7,10:10,0)"),
            ("D7", "=SUM($10:$10)"),
            ("D8", "=SUM(Data!A:A)"),
            ("D9", "=ROWS(A:A)"),
            ("D11", "=SUM(A:B)"),
            ("E1", "=SUM($B$1:B1)*COUNTA(A:A)"),
            ("E2", "=SUM($B$1:B2)*COUNTA(A:A)"),
        ];
        for (cell, formula) in formulas {
            sheet.set_cell_formula(cell, formula).unwrap();
        }
    }
    let data = wb.add_worksheet_with_name("Data").unwrap();
    {
        let sheet = wb.worksheet_mut(data).unwrap();
        sheet.set_cell_value("A1", 4.0).unwrap();
        sheet.set_cell_value("A2", 6.0).unwrap();
    }
    wb.calculate().unwrap();

    let sheet = wb.worksheet(0).unwrap();
    let number = |cell: &str| sheet.get_value(cell).unwrap();
    assert_eq!(number("D1"), CellValue::Number(6.0));
    assert_eq!(number("D2"), CellValue::Number(3.0));
    assert_eq!(number("D3"), CellValue::Number(4.0));
    assert_eq!(number("D4"), CellValue::Number(1.0));
    assert_eq!(number("D5"), CellValue::Number(2.0));
    assert_eq!(number("D6"), CellValue::Number(7.0));
    assert_eq!(number("D7"), CellValue::Number(12.0));
    assert_eq!(number("D8"), CellValue::Number(10.0));
    assert_eq!(number("D9"), CellValue::Number(1_048_576.0));
    assert_eq!(number("D11"), CellValue::Number(6.0));
    assert_eq!(number("E1"), CellValue::Number(3.0));
    assert_eq!(number("E2"), CellValue::Number(9.0));
}
