//! The cells an expression refers to. A reference is written as a cell or range (`B2`, `Data!A:C`),
//! a defined name, a span of sheets (`Jan:Dec!B2`), another workbook's cells (`[1]Sheet!A1`), or
//! computed by INDEX, OFFSET or INDIRECT. Resolving it is kept apart from reading it: ROWS(A:A)
//! is 1,048,576 while reading A:A takes only the sheet's used cells.

use crate::ast::FormulaExpr;
use crate::error::{FormulaError, FormulaResult};
use crate::evaluator::{evaluate, EvaluationContext, FormulaValue};
use duke_sheets_core::{CellError, CellRange, MAX_COLS, MAX_ROWS};

/// A resolved reference: its sheet (None for the formula's own sheet) and its cells.
pub(crate) type Target = (Option<String>, CellRange);

/// What a sheet name in a reference stands for. The parser keeps a span (`Jan:Dec`) and another
/// workbook's sheet (`[1]May 2021`) in the sheet name as Excel writes them; this reads them once.
pub(crate) enum Sheet<'s> {
    One(Option<&'s str>),
    Span(&'s str, &'s str),
    External(&'s str),
}

impl<'s> Sheet<'s> {
    pub(crate) fn of(name: Option<&'s str>) -> Self {
        match name {
            Some(name) if name.starts_with('[') => Sheet::External(name),
            Some(name) => match name.split_once(':') {
                Some((first, last)) => Sheet::Span(first, last),
                None => Sheet::One(Some(name)),
            },
            None => Sheet::One(None),
        }
    }
}

/// Another workbook isn't open, so a formula that reads it keeps the value saved in the file, as
/// Excel shows it until the link is refreshed.
pub(crate) fn external(book: &str) -> FormulaError {
    FormulaError::ExternalReference(book.to_string())
}

/// The cells `expr` refers to: `None` when it isn't a reference (a number, an array, …), or the
/// Excel error a computed reference gives (`INDEX(A1:A3,9)` is #REF!).
pub(crate) fn reference_of(
    expr: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<Option<Result<Target, CellError>>> {
    Ok(Some(match expr {
        FormulaExpr::ExternalRef(r) => return Err(external(&r.book)),
        FormulaExpr::CellRef(c) => match Sheet::of(c.sheet.as_deref()) {
            Sheet::External(book) => return Err(external(book)),
            _ => Ok((c.sheet.clone(), CellRange::new(c.address, c.address))),
        },
        FormulaExpr::RangeRef(r) => match Sheet::of(r.sheet.as_deref()) {
            Sheet::External(book) => return Err(external(book)),
            _ => Ok((r.sheet.clone(), r.range)),
        },
        FormulaExpr::NameRef(name) if ctx.local(name).is_none() => {
            match name_formula(name, ctx) {
                Some(ast @ (FormulaExpr::CellRef(_) | FormulaExpr::RangeRef(_))) => return reference_of(&ast, ctx),
                _ => return Ok(None),
            }
        }
        FormulaExpr::Function { name, args } => match crate::evaluator::function_name(name) {
            "INDEX" => index_target(args, ctx)?,
            "OFFSET" => offset_target(args, ctx)?,
            "INDIRECT" => indirect_target(args, ctx)?,
            _ => return Ok(None),
        },
        _ => return Ok(None),
    }))
}

/// The formula a defined name holds, parsed. Excel stores it without the '=': a reference
/// (`Sheet1!$A$1:$D$10`), a calculation (`SUM(Sheet1!$A$1:$A$3)`) or a constant.
pub(crate) fn name_formula(name: &str, ctx: &EvaluationContext) -> Option<FormulaExpr> {
    let named = ctx.workbook?.get_named_range(name, ctx.current_sheet)?;
    crate::parser::parse_formula(&format!("={}", named.refers_to.trim_start_matches('='))).ok()
}

/// A reference's values: one cell's value, or the range's values read as far as the used cells
/// go; a span of sheets stacks the same cells from each sheet, which SUM, COUNT, AVERAGE, MIN and
/// MAX read like one range.
pub(crate) fn read(ctx: &EvaluationContext, (sheet, range): &Target) -> FormulaValue {
    let (r0, c0, r1, c1) = (range.start.row, range.start.col, range.end.row, range.end.col);
    match Sheet::of(sheet.as_deref()) {
        Sheet::Span(first, last) => {
            let Some(workbook) = ctx.workbook else {
                return FormulaValue::Error(CellError::Ref);
            };
            let (Ok(a), Ok(b)) =
                (ctx.resolve_sheet_index(Some(first), workbook), ctx.resolve_sheet_index(Some(last), workbook))
            else {
                return FormulaValue::Error(CellError::Ref);
            };
            let mut data = Vec::new();
            for index in a.min(b)..=a.max(b) {
                let Some(name) = workbook.worksheet(index).map(|ws| ws.name().to_string()) else { continue };
                match ctx.get_range_values(Some(&name), r0, c0, r1, c1) {
                    FormulaValue::Array { data: rows, .. } => data.extend(rows),
                    FormulaValue::Error(e) => return FormulaValue::Error(e),
                    value => data.push(vec![value]),
                }
            }
            FormulaValue::Array { data, source: None }
        }
        _ if range.start == range.end => ctx.get_cell_value(sheet.as_deref(), r0, c0),
        _ => ctx.get_range_values(sheet.as_deref(), r0, c0, r1, c1),
    }
}

/// `A1:INDEX(B1:B9,3)`, `INDEX(…):INDEX(…)`, `Name:A9`: a range between two references, at least
/// one of them computed. It is the smallest rectangle holding both, as in Excel.
pub(crate) fn evaluate_range_between(
    left: &FormulaExpr,
    right: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let (left, right) = match (reference_of(left, ctx)?, reference_of(right, ctx)?) {
        (Some(Ok(l)), Some(Ok(r))) => (l, r),
        (Some(Err(e)), _) | (_, Some(Err(e))) => return Ok(FormulaValue::Error(e)),
        _ => return Ok(FormulaValue::Error(CellError::Value)),
    };
    let sheet = match (left.0, right.0) {
        (Some(l), Some(r)) if !l.eq_ignore_ascii_case(&r) => return Ok(FormulaValue::Error(CellError::Ref)),
        (l, r) => l.or(r),
    };
    let (a, b) = (left.1, right.1);
    let range = CellRange::from_indices(
        a.start.row.min(b.start.row),
        a.start.col.min(b.start.col),
        a.end.row.max(b.end.row),
        a.end.col.max(b.end.col),
    );
    Ok(read(ctx, &(sheet, range)))
}

/// A whole number from an argument, or the Excel error it gives.
fn whole(expr: &FormulaExpr, ctx: &EvaluationContext) -> FormulaResult<Result<i64, CellError>> {
    Ok(match evaluate(expr, ctx)? {
        FormulaValue::Error(e) => Err(e),
        value => value.as_number().map(|n| n.trunc() as i64).ok_or(CellError::Value),
    })
}

/// The base reference of INDEX or OFFSET; a value that isn't a reference is #REF! (an error
/// passes through).
fn base(expr: &FormulaExpr, ctx: &EvaluationContext) -> FormulaResult<Result<Target, CellError>> {
    Ok(match reference_of(expr, ctx)? {
        Some(target) => target,
        None => match evaluate(expr, ctx)? {
            FormulaValue::Error(e) => Err(e),
            _ => Err(CellError::Ref),
        },
    })
}

/// INDEX(area, row, [col]) as a reference: one cell, or a whole row or column of the area for 0.
fn index_target(args: &[FormulaExpr], ctx: &EvaluationContext) -> FormulaResult<Result<Target, CellError>> {
    let [area, rest @ ..] = args else { return Ok(Err(CellError::Value)) };
    let (sheet, range) = match base(area, ctx)? {
        Ok(target) => target,
        Err(e) => return Ok(Err(e)),
    };
    let rows = (range.end.row - range.start.row + 1) as i64;
    let cols = (range.end.col - range.start.col + 1) as i64;
    // With one position, it counts along a single row or column, else down the rows.
    let (row, col) = match rest {
        [r] if cols == 1 => (whole(r, ctx)?, Ok(1)),
        [r] if rows == 1 => (Ok(1), whole(r, ctx)?),
        [r] => (whole(r, ctx)?, Ok(0)),
        [r, c, ..] => (whole(r, ctx)?, whole(c, ctx)?),
        [] => return Ok(Err(CellError::Value)),
    };
    let (row, col) = match (row, col) {
        (Ok(row), Ok(col)) => (row, col),
        (Err(e), _) | (_, Err(e)) => return Ok(Err(e)),
    };
    if row < 0 || col < 0 {
        return Ok(Err(CellError::Value));
    }
    if row > rows || col > cols {
        return Ok(Err(CellError::Ref));
    }
    let (r0, r1) = match row {
        0 => (range.start.row, range.end.row),
        _ => (range.start.row + row as u32 - 1, range.start.row + row as u32 - 1),
    };
    let (c0, c1) = match col {
        0 => (range.start.col, range.end.col),
        _ => (range.start.col + col as u16 - 1, range.start.col + col as u16 - 1),
    };
    Ok(Ok((sheet, CellRange::from_indices(r0, c0, r1, c1))))
}

/// OFFSET(reference, rows, cols, [height], [width]) as a reference.
pub(crate) fn offset_target(
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<Result<Target, CellError>> {
    let [reference, rows, cols, size @ ..] = args else { return Ok(Err(CellError::Value)) };
    let (dr, dc) = match (whole(rows, ctx)?, whole(cols, ctx)?) {
        (Ok(dr), Ok(dc)) => (dr, dc),
        (Err(e), _) | (_, Err(e)) => return Ok(Err(e)),
    };
    let (sheet, range) = match base(reference, ctx)? {
        Ok(target) => target,
        Err(e) => return Ok(Err(e)),
    };
    // A left-out height or width keeps the reference's own; one below 1 is #REF!
    let extent = |arg: Option<&FormulaExpr>, own: i64| -> FormulaResult<Result<i64, CellError>> {
        Ok(match arg {
            None | Some(FormulaExpr::Empty) => Ok(own),
            Some(expr) => match whole(expr, ctx)? {
                Ok(n) if n >= 1 => Ok(n),
                Ok(_) => Err(CellError::Ref),
                Err(e) => Err(e),
            },
        })
    };
    let height = extent(size.first(), (range.end.row - range.start.row + 1) as i64)?;
    let width = extent(size.get(1), (range.end.col - range.start.col + 1) as i64)?;
    let (height, width) = match (height, width) {
        (Ok(h), Ok(w)) => (h, w),
        (Err(e), _) | (_, Err(e)) => return Ok(Err(e)),
    };
    let r0 = range.start.row as i64 + dr;
    let c0 = range.start.col as i64 + dc;
    if r0 < 0 || c0 < 0 || r0 + height > MAX_ROWS as i64 || c0 + width > MAX_COLS as i64 {
        return Ok(Err(CellError::Ref));
    }
    Ok(Ok((
        sheet,
        CellRange::from_indices(r0 as u32, c0 as u16, (r0 + height - 1) as u32, (c0 + width - 1) as u16),
    )))
}

/// INDIRECT(text, [a1]) as a reference. Its text is evaluated once; an error in it is the result.
pub(crate) fn indirect_target(
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<Result<Target, CellError>> {
    let [text, style @ ..] = args else { return Ok(Err(CellError::Value)) };
    let text = match evaluate(text, ctx)? {
        FormulaValue::Error(e) => return Ok(Err(e)),
        FormulaValue::Array { .. } => return Ok(Err(CellError::Value)),
        value => value.as_string(),
    };
    let a1 = match style.first() {
        Some(expr) => match evaluate(expr, ctx)? {
            FormulaValue::Error(e) => return Ok(Err(e)),
            value => value.as_bool().unwrap_or(true),
        },
        None => true,
    };
    if ctx.workbook.is_none() {
        return Ok(Err(CellError::Ref));
    }
    Ok(indirect_reference(text.trim(), a1, ctx).ok_or(CellError::Ref))
}

/// The reference INDIRECT's text names: A1 style (`B2`, `Data!A1:C9`, `A:A`, `1:5`), or R1C1 style
/// (`R2C2`, `R[-1]C`, `R1C1:R3C2`) relative to the formula's cell.
fn indirect_reference(text: &str, a1: bool, ctx: &EvaluationContext) -> Option<Target> {
    if a1 {
        return match crate::parser::parse_formula(&format!("={}", text)) {
            Ok(FormulaExpr::CellRef(c)) => Some((c.sheet, CellRange::new(c.address, c.address))),
            Ok(FormulaExpr::RangeRef(r)) => Some((r.sheet, r.range)),
            _ => None,
        };
    }
    let (sheet, cells) = match text.rsplit_once('!') {
        Some((sheet, cells)) => (Some(sheet.trim_matches('\'').to_string()), cells),
        None => (None, text),
    };
    let point = |part: &str| -> Option<(u32, u16)> {
        let upper = part.trim().to_ascii_uppercase();
        let (row_part, col_part) = upper.strip_prefix('R')?.split_once('C')?;
        let axis = |part: &str, here: i64| -> Option<i64> {
            if part.is_empty() {
                Some(here)
            } else if let Some(offset) = part.strip_prefix('[').and_then(|p| p.strip_suffix(']')) {
                Some(here + offset.parse::<i64>().ok()?)
            } else {
                Some(part.parse::<i64>().ok()? - 1)
            }
        };
        let row = axis(row_part, ctx.current_row as i64)?;
        let col = axis(col_part, ctx.current_col as i64)?;
        ((0..MAX_ROWS as i64).contains(&row) && (0..MAX_COLS as i64).contains(&col)).then_some((row as u32, col as u16))
    };
    let (first, last) = match cells.split_once(':') {
        Some((a, b)) => (point(a)?, point(b)?),
        None => {
            let p = point(cells)?;
            (p, p)
        }
    };
    Some((sheet, CellRange::from_indices(first.0, first.1, last.0, last.1)))
}
