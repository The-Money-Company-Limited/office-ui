//! How a function's arguments meet arrays: Excel applies a function element by element where an
//! argument takes one value (`ArgPositions`), broadcasting a single row or column across the other
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

/// A result fitted to a `rows` x `cols` block, as a Ctrl+Shift+Enter array formula shows it: a
/// single value or a single row or column repeats, and cells past the result's end are #N/A.
pub fn fit_array(result: &FormulaValue, rows: usize, cols: usize) -> Vec<Vec<FormulaValue>> {
    let data = crate::evaluator::to_array(result);
    let d = dims(&data);
    (0..rows)
        .map(|r| {
            (0..cols)
                .map(|c| {
                    broadcast_at(&data, d, r, c)
                        .cloned()
                        .unwrap_or(FormulaValue::Error(CellError::Na))
                })
                .collect()
        })
        .collect()
}

/// Whether an argument's value comes from cells: a reference, a defined name, a LET name bound to
/// a reference, or a function that hands one through (INDEX, OFFSET, XLOOKUP, IF(c,A1,B1),
/// IFERROR(A1,0), LET's result). Anything else is typed into the formula, including IF(A1,"5")
/// and other functions' results.
fn from_cells(expr: &FormulaExpr, ctx: &EvaluationContext) -> bool {
    match expr {
        FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_) | FormulaExpr::StructuredRef(_) | FormulaExpr::ExternalRef(_) => true,
        FormulaExpr::NameRef(name) => {
            matches!(ctx.local(name), None | Some(crate::lambda::Binding::Reference(_)))
        }
        // `@A1:A9` reads one cell of the range.
        FormulaExpr::UnaryOp {
            op: crate::ast::UnaryOperator::ImplicitIntersection,
            operand,
        } => from_cells(operand, ctx),
        // These hand back one of their arguments, so the value is from cells when an argument is.
        FormulaExpr::Function { name, args } => {
            let name = crate::evaluator::function_name(name);
            match crate::functions::registry().get(name) {
                Some(def) => match def.returns {
                    crate::functions::Returns::Reference => true,
                    crate::functions::Returns::Last => {
                        args.last().is_some_and(|arg| from_cells(arg, ctx))
                    }
                    // IFERROR and IFNA take values, and hand back one of them as it is.
                    _ if matches!(name, "IFERROR" | "IFNA") => {
                        args.iter().any(|arg| from_cells(arg, ctx))
                    }
                    _ => args.iter().enumerate().any(|(i, arg)| {
                        def.result_args.includes_in(i, args.len()) && from_cells(arg, ctx)
                    }),
                },
                None => false,
            }
        }
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

/// A one-element array where the function takes one value is that value:
/// `LEFT(TEXTSPLIT("a, b",", "),1)` with one piece of text.
pub(crate) fn unwrap_single_elements(func: &FunctionDef, args: &mut [FormulaValue]) {
    for (position, value) in args.iter_mut().enumerate() {
        if let FormulaValue::Array { data, .. } = value {
            if func.scalar_args.includes(position) && dims(data) == (1, 1) {
                *value = data[0][0].clone();
            }
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
