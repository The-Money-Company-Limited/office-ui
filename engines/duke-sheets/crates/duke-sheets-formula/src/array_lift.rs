//! How a function's arguments meet arrays: Excel applies a function element by element where an
//! argument takes one value (`ScalarArgs`), broadcasting a single row or column across the other
//! arrays, and reads values typed into the formula differently from values in cells (`TypedArgs`).

use crate::ast::FormulaExpr;
use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use crate::functions::{FunctionDef, TypedArgs};
use duke_sheets_core::CellError;

/// The rows and columns of an array.
pub(crate) fn dims(arr: &[Vec<FormulaValue>]) -> (usize, usize) {
    (arr.len(), arr.first().map_or(0, Vec::len))
}

/// The element of a `rows` x `cols` array at (row, col) of a broadcast: a single row repeats down
/// and a single column repeats across; past the end of a longer side there is nothing (#N/A).
pub(crate) fn broadcast_at(
    arr: &[Vec<FormulaValue>],
    (rows, cols): (usize, usize),
    row: usize,
    col: usize,
) -> Option<&FormulaValue> {
    let row = if rows == 1 { 0 } else { row };
    let col = if cols == 1 { 0 } else { col };
    arr.get(row).and_then(|line| line.get(col))
}

/// Whether an argument's value comes from cells: a reference, a defined name, or a function that
/// hands one through (INDEX, OFFSET, or IF(c,A1,B1)). Anything else is typed into the formula,
/// including IF(1,"5"), other functions' results and LET names.
fn from_cells(expr: &FormulaExpr, ctx: &EvaluationContext) -> bool {
    match expr {
        FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_) | FormulaExpr::StructuredRef(_) | FormulaExpr::ExternalRef(_) => true,
        FormulaExpr::NameRef(name) => ctx.local(name).is_none(),
        // These hand back one of their arguments, so the value is from cells when an argument is.
        FormulaExpr::Function { name, args } => match crate::evaluator::function_name(name) {
            "INDEX" | "OFFSET" | "INDIRECT" => true,
            "IF" | "IFS" | "IFERROR" | "IFNA" | "CHOOSE" | "SWITCH" | "XLOOKUP" | "LET" => {
                args.iter().any(|arg| from_cells(arg, ctx))
            }
            _ => false,
        },
        _ => false,
    }
}

/// Reads typed values per the function's `TypedArgs`: `SUM("5",TRUE)` is 6 and `SUM("x")` is
/// #VALUE!, while text and TRUE from cells are left out.
pub(crate) fn coerce_typed_arguments(
    rule: TypedArgs,
    exprs: &[FormulaExpr],
    values: &mut [FormulaValue],
    ctx: &EvaluationContext,
) {
    if rule == TypedArgs::AsGiven {
        return;
    }
    let numeric = rule == TypedArgs::Numbers;
    for (expr, value) in exprs.iter().zip(values.iter_mut()) {
        let typed = !from_cells(expr, ctx);
        match value {
            FormulaValue::String(text) if typed => {
                *value = if numeric {
                    match crate::coerce::text_to_number(text) {
                        Some(n) => FormulaValue::Number(n),
                        None => FormulaValue::Error(CellError::Value),
                    }
                } else {
                    match text.to_ascii_uppercase().as_str() {
                        "TRUE" => FormulaValue::Boolean(true),
                        "FALSE" => FormulaValue::Boolean(false),
                        _ => FormulaValue::Error(CellError::Value),
                    }
                };
            }
            FormulaValue::Boolean(b) if typed && numeric => {
                *value = FormulaValue::Number(if *b { 1.0 } else { 0.0 });
            }
            FormulaValue::String(_) => *value = FormulaValue::Empty,
            FormulaValue::Boolean(_) if numeric => *value = FormulaValue::Empty,
            _ => {}
        }
    }
}

/// Applies the function to each element of the arrays at its scalar arguments and returns the
/// array of results, or `None` when no such argument holds an array.
pub(crate) fn lift_over_arrays(
    func: &FunctionDef,
    args: &[FormulaValue],
    ctx: &EvaluationContext,
) -> Option<FormulaResult<FormulaValue>> {
    let arrays: Vec<(usize, (usize, usize))> = args
        .iter()
        .enumerate()
        .filter(|&(position, _)| func.scalar_args.includes(position))
        .filter_map(|(position, value)| match value {
            FormulaValue::Array { data, .. } => {
                let d = dims(data);
                (d.0 > 0 && d.1 > 0 && d != (1, 1)).then_some((position, d))
            }
            _ => None,
        })
        .collect();
    let rows = arrays.iter().map(|(_, (r, _))| *r).max()?;
    let cols = arrays.iter().map(|(_, (_, c))| *c).max()?;
    let mut out = Vec::with_capacity(rows);
    for i in 0..rows {
        let mut row = Vec::with_capacity(cols);
        for j in 0..cols {
            let mut element_args = args.to_vec();
            for &(position, d) in &arrays {
                let FormulaValue::Array { data, .. } = &args[position] else { continue };
                element_args[position] =
                    broadcast_at(data, d, i, j).cloned().unwrap_or(FormulaValue::Error(CellError::Na));
            }
            let value = match (func.implementation)(&element_args, ctx) {
                Ok(FormulaValue::Array { data, .. }) => data
                    .into_iter()
                    .next()
                    .and_then(|line| line.into_iter().next())
                    .unwrap_or(FormulaValue::Error(CellError::Value)),
                Ok(value) => value,
                Err(error) => return Some(Err(error)),
            };
            row.push(value);
        }
        out.push(row);
    }
    Some(Ok(FormulaValue::Array { data: out, source: None }))
}
