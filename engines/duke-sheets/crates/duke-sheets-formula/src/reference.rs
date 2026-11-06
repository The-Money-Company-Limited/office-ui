//! The cells an expression refers to. A reference is written as a cell or range (`B2`, `Data!A:C`),
//! a defined name, a span of sheets (`Jan:Dec!B2`), another workbook's cells (`[1]Sheet!A1`), or
//! computed by INDEX, OFFSET or INDIRECT. Resolving it is kept apart from reading it: ROWS(A:A)
//! is 1,048,576 while reading A:A takes only the sheet's used cells.

use crate::ast::{BinaryOperator, FormulaExpr, RangeReference, UnaryOperator};
use crate::error::{FormulaError, FormulaResult};
use crate::evaluator::{evaluate, EvaluationContext, FormulaValue};
use crate::functions::ArgPositions;
use crate::lambda::Binding;
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

/// Whether `expr` is a reference into another workbook: `[1]May!A1`, `'[1]May 2021'!A2:A9`.
pub(crate) fn is_external(expr: &FormulaExpr) -> bool {
    match expr {
        FormulaExpr::ExternalRef(_) => true,
        FormulaExpr::CellRef(r) => matches!(Sheet::of(r.sheet.as_deref()), Sheet::External(_)),
        FormulaExpr::RangeRef(r) => matches!(Sheet::of(r.sheet.as_deref()), Sheet::External(_)),
        _ => false,
    }
}

/// Another workbook isn't open, so a formula that reads it keeps the value saved in the file, as
/// Excel shows it until the link is refreshed.
pub(crate) fn external(book: &str) -> FormulaError {
    FormulaError::ExternalReference(book.to_string())
}

/// What an expression gives, read once: the cells it refers to, or its value.
pub(crate) enum Resolved {
    Reference(Target),
    Value(FormulaValue),
}

/// The cells `expr` refers to: `None` when it isn't a reference (a number, an array, …), or the
/// Excel error a computed reference gives (`INDEX(A1:A3,9)` is #REF!). Besides cells, ranges and
/// tables, a reference can be a name for one, a LET name bound to one, `(A1:B2)`, `A1#`, the
/// cells two ranges share (`A1:B3 B2:C4`), `A1:INDEX(…)`, what INDEX, OFFSET and INDIRECT work
/// out, or what IF, CHOOSE, IFS, SWITCH and XLOOKUP hand back. What isn't a reference is left
/// unevaluated, or evaluated only as far as deciding needs.
pub(crate) fn reference_of(
    expr: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<Option<Result<Target, CellError>>> {
    Ok(match resolve(expr, ctx, false)? {
        Some(Resolved::Reference(target)) => Some(Ok(target)),
        Some(Resolved::Value(FormulaValue::Error(e))) => Some(Err(e)),
        _ => None,
    })
}

/// `expr` read once: the cells it refers to, else its value. `@IF(RAND()>0.5,A1:A3*2,0)` decides
/// once.
pub(crate) fn resolve_value(
    expr: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<Resolved> {
    match resolve(expr, ctx, true)? {
        Some(resolved) => Ok(resolved),
        None => Ok(Resolved::Value(evaluate(expr, ctx)?)),
    }
}

/// The cells, or the error, of a reference worked out.
fn worked_out(target: Result<Target, CellError>) -> Option<Resolved> {
    Some(match target {
        Ok(target) => Resolved::Reference(target),
        Err(e) => Resolved::Value(FormulaValue::Error(e)),
    })
}

/// `reference_of` (`value` false) or `resolve_value` (`value` true): what isn't a reference is
/// `None`, or its value.
fn resolve(
    expr: &FormulaExpr,
    ctx: &EvaluationContext,
    value: bool,
) -> FormulaResult<Option<Resolved>> {
    let otherwise = || -> FormulaResult<Option<Resolved>> {
        Ok(if value {
            Some(Resolved::Value(evaluate(expr, ctx)?))
        } else {
            None
        })
    };
    Ok(match expr {
        FormulaExpr::ExternalRef(r) => return Err(external(&r.book)),
        FormulaExpr::CellRef(c) => match Sheet::of(c.sheet.as_deref()) {
            Sheet::External(book) => return Err(external(book)),
            _ => Some(Resolved::Reference((
                c.sheet.clone(),
                CellRange::new(c.address, c.address),
            ))),
        },
        FormulaExpr::RangeRef(r) => match Sheet::of(r.sheet.as_deref()) {
            Sheet::External(book) => return Err(external(book)),
            _ => Some(Resolved::Reference((r.sheet.clone(), r.range))),
        },
        FormulaExpr::StructuredRef(sr) => worked_out(ctx.structured_target(sr)?),
        FormulaExpr::UnaryOp {
            op: UnaryOperator::Paren,
            operand,
        } => return resolve(operand, ctx, value),
        FormulaExpr::UnaryOp {
            op: UnaryOperator::SpillRange,
            operand,
        } => match operand.as_ref() {
            FormulaExpr::CellRef(c) => worked_out(spill_target(
                c.sheet.as_deref(),
                c.address.row,
                c.address.col,
                ctx,
            )),
            _ => return otherwise(),
        },
        FormulaExpr::BinaryOp {
            op: BinaryOperator::Range,
            left,
            right,
        } => worked_out(range_between(left, right, ctx)?),
        FormulaExpr::BinaryOp {
            op: BinaryOperator::Intersect,
            left,
            right,
        } => worked_out(shared_cells(left, right, ctx)?),
        FormulaExpr::NameRef(name) => match ctx.local(name) {
            Some(Binding::Reference(target)) => Some(Resolved::Reference(target.clone())),
            Some(Binding::Value(v)) if value => Some(Resolved::Value(v.clone())),
            Some(_) => return otherwise(),
            None => match name_formula(name, ctx) {
                Some(formula) => match nested(|| resolve(&formula, ctx, value)) {
                    Some(resolved) => resolved?,
                    None => worked_out(Err(CellError::Ref)),
                },
                None => return otherwise(),
            },
        },
        FormulaExpr::Function { name, args } => {
            let lookup = crate::evaluator::function_name(name);
            match lookup {
                "INDEX" => return index(args, ctx, value),
                "OFFSET" => worked_out(offset_target(args, ctx)?),
                "INDIRECT" => worked_out(indirect_target(args, ctx)?),
                "XLOOKUP" => match xlookup_target(args, ctx)? {
                    Some(target) => worked_out(target),
                    None => return otherwise(),
                },
                _ => match crate::functions::registry().get(lookup) {
                    Some(def) if def.result_args != ArgPositions::None => {
                        return chosen(lookup, def, args, ctx, value)
                    }
                    _ => return otherwise(),
                },
            }
        }
        _ => return otherwise(),
    })
}

/// Runs `f` one level deeper into names that refer to names, or gives `None` past 64 levels.
pub(crate) fn nested<T>(f: impl FnOnce() -> T) -> Option<T> {
    thread_local! {
        static DEPTH: std::cell::Cell<u32> = const { std::cell::Cell::new(0) };
    }
    let depth = DEPTH.with(|d| d.get());
    if depth >= 64 {
        return None;
    }
    DEPTH.with(|d| d.set(depth + 1));
    let result = f();
    DEPTH.with(|d| d.set(depth));
    Some(result)
}

/// What IF, CHOOSE, IFS or SWITCH hands back: the result argument its condition, index or value
/// picks, read as a reference when it is one. What decides is evaluated once; when it doesn't
/// pick one argument (an array, an error), the function runs on the values already read.
fn chosen(
    name: &str,
    def: &crate::functions::FunctionDef,
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
    value: bool,
) -> FormulaResult<Option<Resolved>> {
    let mut given: Vec<Option<FormulaValue>> = vec![None; args.len()];
    let read =
        |i: usize, given: &mut Vec<Option<FormulaValue>>| -> FormulaResult<Option<FormulaValue>> {
            let Some(arg) = args.get(i) else {
                return Ok(None);
            };
            let v = evaluate(arg, ctx)?;
            given[i] = Some(v.clone());
            Ok(match v {
                FormulaValue::Array { .. } | FormulaValue::Error(_) => None,
                v => Some(v),
            })
        };
    let picked = match name {
        "IF" => read(0, &mut given)?
            .and_then(|v| v.as_bool())
            .map(|yes| if yes { 1 } else { 2 }),
        "CHOOSE" => read(0, &mut given)?
            .and_then(|v| v.as_number())
            .map(f64::trunc)
            .filter(|n| *n >= 1.0)
            .map(|n| n as usize),
        "IFS" => {
            let mut picked = None;
            for i in (0..args.len()).step_by(2) {
                match read(i, &mut given)? {
                    Some(v) if v.as_bool() == Some(true) => {
                        picked = Some(i + 1);
                        break;
                    }
                    Some(_) => {}
                    None => break,
                }
            }
            picked
        }
        "SWITCH" => match read(0, &mut given)? {
            Some(target) => {
                let mut picked = None;
                let mut i = 1;
                while i + 1 < args.len() {
                    match read(i, &mut given)? {
                        Some(v)
                            if crate::evaluator::compare_values(&target, &v) == 0
                                && std::mem::discriminant(&target)
                                    == std::mem::discriminant(&v) =>
                        {
                            picked = Some(i + 1);
                            break;
                        }
                        Some(_) => {}
                        None => {
                            picked = None;
                            i = args.len();
                            break;
                        }
                    }
                    i += 2;
                }
                // No match: the default, a last argument left over.
                picked.or_else(|| (i + 1 == args.len()).then_some(i))
            }
            None => None,
        },
        _ => None,
    };
    let picked = picked.filter(|&i| {
        def.result_args.includes_in(i, args.len())
            && args.get(i).is_some_and(|a| *a != FormulaExpr::Empty)
    });
    match picked {
        Some(i) => resolve(&args[i], ctx, value),
        None if value => Ok(Some(Resolved::Value(crate::evaluator::call_with(
            def, args, given, ctx,
        )?))),
        None => Ok(None),
    }
}

/// The cell or row XLOOKUP hands back from a return range, or its if_not_found reference.
fn xlookup_target(
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<Option<Result<Target, CellError>>> {
    let [lookup, lookup_array, returned, rest @ ..] = args else {
        return Ok(None);
    };
    let Some(Ok((sheet, range))) = reference_of(returned, ctx)? else {
        return Ok(None);
    };
    // The lookup array's shape, read once: a reference's, or its value's.
    let (looked_in, (rows, cols)) = match reference_of(lookup_array, ctx)? {
        Some(Ok((_, r))) => (None, (r.row_count() as usize, r.col_count() as usize)),
        Some(Err(e)) => return Ok(Some(Err(e))),
        None => {
            let value = evaluate(lookup_array, ctx)?;
            let shape = match &value {
                FormulaValue::Array { data, .. } => crate::array_lift::dims(data),
                _ => (1, 1),
            };
            (Some(value), shape)
        }
    };
    let vertical = rows > 1 || cols == 1;
    // As XLOOKUP: the lookup array is one row or column, and the return range runs as far.
    let fits = if vertical { range.row_count() as usize == rows } else { range.col_count() as usize == cols };
    if (rows > 1 && cols > 1) || !fits {
        return Ok(Some(Err(CellError::Value)));
    }
    let mut xmatch = vec![lookup.clone(), lookup_array.clone()];
    xmatch.extend(rest.iter().skip(1).cloned());
    let position = match (looked_in, crate::functions::registry().get("XMATCH")) {
        (Some(value), Some(def)) => {
            let mut given = vec![None; xmatch.len()];
            given[1] = Some(value);
            crate::evaluator::call_with(def, &xmatch, given, ctx)?
        }
        _ => crate::evaluator::evaluate_function("XMATCH", &xmatch, ctx)?,
    };
    Ok(match position {
        FormulaValue::Number(n) if n >= 1.0 => {
            let n = n as u32 - 1;
            Some(Ok(if vertical {
                (
                    sheet,
                    CellRange::from_indices(
                        range.start.row + n,
                        range.start.col,
                        range.start.row + n,
                        range.end.col,
                    ),
                )
            } else {
                let n = n as u16;
                (
                    sheet,
                    CellRange::from_indices(
                        range.start.row,
                        range.start.col + n,
                        range.end.row,
                        range.start.col + n,
                    ),
                )
            }))
        }
        FormulaValue::Error(CellError::Na) => match rest.first() {
            Some(FormulaExpr::Empty) | None => Some(Err(CellError::Na)),
            Some(not_found) => return reference_of(not_found, ctx),
        },
        FormulaValue::Error(e) => Some(Err(e)),
        _ => None,
    })
}

/// `A1#`: the cells a formula in A1 spills over, or A1 alone.
fn spill_target(
    sheet: Option<&str>,
    row: u32,
    col: u16,
    ctx: &EvaluationContext,
) -> Result<Target, CellError> {
    let workbook = ctx.workbook.ok_or(CellError::Ref)?;
    let index = ctx
        .resolve_sheet_index(sheet, workbook)
        .map_err(|_| CellError::Ref)?;
    let (rows, cols) = workbook
        .worksheet(index)
        .and_then(|ws| ws.get_spill_info(row, col))
        .map_or((1, 1), |info| (info.rows, info.cols));
    Ok((
        sheet.map(str::to_string),
        CellRange::from_indices(row, col, row + rows - 1, col + cols - 1),
    ))
}

/// The cells two references share (`A1:B3 B2:C4` is B2:B3), or #NULL! when they share none.
fn shared_cells(
    left: &FormulaExpr,
    right: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<Result<Target, CellError>> {
    let ((sheet, a), (other, b)) = match (base(left, ctx)?, base(right, ctx)?) {
        (Ok(l), Ok(r)) => (l, r),
        (Err(e), _) | (_, Err(e)) => return Ok(Err(e)),
    };
    // A reference without a sheet is on the formula's own sheet.
    let own = ctx
        .workbook
        .and_then(|wb| wb.worksheet(ctx.current_sheet))
        .map(|ws| ws.name());
    let named = |s: &Option<String>| s.as_deref().or(own).map(str::to_ascii_lowercase);
    let same_sheet = named(&sheet) == named(&other);
    let sheet = sheet.or(other);
    let (r0, r1) = (a.start.row.max(b.start.row), a.end.row.min(b.end.row));
    let (c0, c1) = (a.start.col.max(b.start.col), a.end.col.min(b.end.col));
    if !same_sheet || r0 > r1 || c0 > c1 {
        return Ok(Err(CellError::Null));
    }
    Ok(Ok((sheet, CellRange::from_indices(r0, c0, r1, c1))))
}

/// The cell where a reference meets the formula's own row and column, as `@` reads it: a single
/// row or column meets the formula's column or row, and a block both. #VALUE! when none does.
pub(crate) fn intersect_cell((sheet, range): &Target, ctx: &EvaluationContext) -> FormulaValue {
    let along = |first: u32, last: u32, here: u32| {
        if first == last {
            Some(first)
        } else {
            (first..=last).contains(&here).then_some(here)
        }
    };
    let row = along(range.start.row, range.end.row, ctx.current_row);
    let col = along(
        range.start.col.into(),
        range.end.col.into(),
        ctx.current_col.into(),
    );
    match (row, col) {
        (Some(row), Some(col)) => read(
            ctx,
            &(
                sheet.clone(),
                CellRange::from_indices(row, col as u16, row, col as u16),
            ),
        ),
        _ => FormulaValue::Error(CellError::Value),
    }
}

/// The formula a defined name holds, parsed. Excel stores it without the '=': a reference
/// (`Sheet1!$A$1:$D$10`), a calculation (`SUM(Sheet1!$A$1:$A$3)`) or a constant.
pub(crate) fn name_formula(name: &str, ctx: &EvaluationContext) -> Option<FormulaExpr> {
    let workbook = ctx.workbook?;
    match workbook.get_named_range(name, ctx.current_sheet) {
        Some(named) => {
            crate::parser::parse_formula(&format!("={}", named.refers_to.trim_start_matches('=')))
                .ok()
        }
        // A table's name alone is its data rows, as `Table1[#Data]`.
        None => workbook
            .worksheets()
            .any(|ws| ws.table_by_name(name).is_some())
            .then(|| {
                FormulaExpr::StructuredRef(crate::ast::StructuredReference {
                    table: Some(name.to_string()),
                    column: None,
                    last_column: None,
                    specifiers: Vec::new(),
                })
            }),
    }
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
    Ok(match range_between(left, right, ctx)? {
        Ok(target) => read(ctx, &target),
        Err(e) => FormulaValue::Error(e),
    })
}

fn range_between(
    left: &FormulaExpr,
    right: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<Result<Target, CellError>> {
    let (left, right) = match (reference_of(left, ctx)?, reference_of(right, ctx)?) {
        (Some(Ok(l)), Some(Ok(r))) => (l, r),
        (Some(Err(e)), _) | (_, Some(Err(e))) => return Ok(Err(e)),
        _ => return Ok(Err(CellError::Value)),
    };
    let sheet = match (left.0, right.0) {
        (Some(l), Some(r)) if !l.eq_ignore_ascii_case(&r) => return Ok(Err(CellError::Ref)),
        (l, r) => l.or(r),
    };
    let (a, b) = (left.1, right.1);
    let range = CellRange::from_indices(
        a.start.row.min(b.start.row),
        a.start.col.min(b.start.col),
        a.end.row.max(b.end.row),
        a.end.col.max(b.end.col),
    );
    Ok(Ok((sheet, range)))
}

/// A whole number from a value, or the Excel error it gives.
fn whole_of(value: &FormulaValue) -> Result<i64, CellError> {
    match value {
        FormulaValue::Error(e) => Err(*e),
        value => value
            .as_number()
            .map(|n| n.trunc() as i64)
            .ok_or(CellError::Value),
    }
}

/// A whole number from an argument, or the Excel error it gives.
fn whole(expr: &FormulaExpr, ctx: &EvaluationContext) -> FormulaResult<Result<i64, CellError>> {
    Ok(whole_of(&evaluate(expr, ctx)?))
}

/// The base reference of OFFSET and of a shared cell range; a value that isn't a reference is
/// #REF! (an error passes through).
fn base(expr: &FormulaExpr, ctx: &EvaluationContext) -> FormulaResult<Result<Target, CellError>> {
    Ok(match resolve_value(expr, ctx)? {
        Resolved::Reference(target) => Ok(target),
        Resolved::Value(FormulaValue::Error(e)) => Err(e),
        Resolved::Value(_) => Err(CellError::Ref),
    })
}

/// The areas of INDEX's reference: one range, or each range of a union, `(A1:B2,C1:D2)`. `None`
/// when a part isn't a reference: INDEX then picks from its value.
fn areas(
    expr: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<Option<Result<Vec<Target>, CellError>>> {
    fn union_parts<'e>(expr: &'e FormulaExpr, parts: &mut Vec<&'e FormulaExpr>) {
        match expr {
            FormulaExpr::UnaryOp {
                op: UnaryOperator::Paren,
                operand,
            } => union_parts(operand, parts),
            FormulaExpr::BinaryOp {
                op: BinaryOperator::Union,
                left,
                right,
            } => {
                union_parts(left, parts);
                union_parts(right, parts);
            }
            other => parts.push(other),
        }
    }
    let mut parts = Vec::new();
    union_parts(expr, &mut parts);
    let mut areas = Vec::with_capacity(parts.len());
    for part in parts {
        match reference_of(part, ctx)? {
            Some(Ok(target)) => areas.push(target),
            Some(Err(e)) => return Ok(Some(Err(e))),
            None => return Ok(None),
        }
    }
    Ok(Some(Ok(areas)))
}

/// INDEX(reference, row, [col], [area]) as a reference: one cell, or a whole row or column of the
/// area for 0. A reference written as one range has only area 1. Given arrays of rows or
/// columns, it is the array of the cells they pick (`INDEX(Table1[Amount],{1;2;3})`). Given an
/// array rather than a reference, INDEX picks from its value.
fn index(
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
    value: bool,
) -> FormulaResult<Option<Resolved>> {
    use crate::array_lift::{broadcast_at, dims};
    let general = || -> FormulaResult<Option<Resolved>> {
        Ok(if value {
            let def = crate::functions::registry()
                .get("INDEX")
                .expect("INDEX is registered");
            Some(Resolved::Value(crate::evaluator::call_with(
                def,
                args,
                vec![None; args.len()],
                ctx,
            )?))
        } else {
            None
        })
    };
    let [reference, rest @ ..] = args else {
        return Ok(worked_out(Err(CellError::Value)));
    };
    let areas = match areas(reference, ctx)? {
        Some(Ok(areas)) => areas,
        Some(Err(e)) => return Ok(worked_out(Err(e))),
        None => return general(),
    };
    let mut numbers = Vec::with_capacity(rest.len());
    for expr in rest {
        numbers.push(match expr {
            FormulaExpr::Empty => None,
            expr => Some(evaluate(expr, ctx)?),
        });
    }
    let shape = numbers
        .iter()
        .flatten()
        .fold((1, 1), |(rows, cols), number| match number {
            FormulaValue::Array { data, .. } => (rows.max(dims(data).0), cols.max(dims(data).1)),
            _ => (rows, cols),
        });
    let pick = |i: usize, j: usize| -> Result<Target, CellError> {
        let at = |k: usize| -> Result<Option<i64>, CellError> {
            match numbers.get(k) {
                None | Some(None) => Ok(None),
                Some(Some(FormulaValue::Array { data, .. })) => {
                    match broadcast_at(data, dims(data), i, j) {
                        Some(element) => whole_of(element).map(Some),
                        None => Err(CellError::Na),
                    }
                }
                Some(Some(number)) => whole_of(number).map(Some),
            }
        };
        index_range(&areas, rest.len(), at(0)?, at(1)?, at(2)?)
    };
    if shape == (1, 1) {
        return Ok(worked_out(pick(0, 0)));
    }
    if !value {
        return Ok(None);
    }
    let data = (0..shape.0)
        .map(|i| {
            (0..shape.1)
                .map(|j| match pick(i, j) {
                    Ok(target) => match read(ctx, &target) {
                        FormulaValue::Array { data, .. } => data
                            .into_iter()
                            .next()
                            .and_then(|r| r.into_iter().next())
                            .unwrap_or(FormulaValue::Empty),
                        value => value,
                    },
                    Err(e) => FormulaValue::Error(e),
                })
                .collect()
        })
        .collect();
    Ok(Some(Resolved::Value(FormulaValue::Array {
        data,
        source: None,
    })))
}

/// The cells INDEX picks from `areas` given its numbers (`count` of them after the reference):
/// with one, it counts along a single row or column, else down the rows; 0 or a left-out column
/// is the whole row or column.
fn index_range(
    areas: &[Target],
    count: usize,
    row: Option<i64>,
    col: Option<i64>,
    area: Option<i64>,
) -> Result<Target, CellError> {
    let area = match area {
        None => 1,
        Some(n) if n >= 1 => n as usize,
        Some(_) => return Err(CellError::Value),
    };
    let (sheet, range) = areas.get(area - 1).cloned().ok_or(CellError::Ref)?;
    let rows = (range.end.row - range.start.row + 1) as i64;
    let cols = (range.end.col - range.start.col + 1) as i64;
    let first = row.unwrap_or(0);
    let (row, col) = match count {
        0 => return Err(CellError::Value),
        1 if cols == 1 => (first, 1),
        1 if rows == 1 => (1, first),
        1 => (first, 0),
        _ => (first, col.unwrap_or(0)),
    };
    if row < 0 || col < 0 {
        return Err(CellError::Value);
    }
    if row > rows || col > cols {
        return Err(CellError::Ref);
    }
    let (r0, r1) = match row {
        0 => (range.start.row, range.end.row),
        _ => (range.start.row + row as u32 - 1, range.start.row + row as u32 - 1),
    };
    let (c0, c1) = match col {
        0 => (range.start.col, range.end.col),
        _ => (range.start.col + col as u16 - 1, range.start.col + col as u16 - 1),
    };
    Ok((sheet, CellRange::from_indices(r0, c0, r1, c1)))
}

/// OFFSET's numbers (rows, cols, height, width): each a whole number, `None` when left out.
type OffsetNumbers = [Result<Option<i64>, CellError>];

/// OFFSET(reference, rows, cols, [height], [width]) as a reference.
pub(crate) fn offset_target(
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<Result<Target, CellError>> {
    // rows and cols are required; `OFFSET(A1)` is an error, as Excel won't take it.
    let [reference, numbers @ ..] = args else {
        return Ok(Err(CellError::Value));
    };
    if numbers.len() < 2 {
        return Ok(Err(CellError::Value));
    }
    let mut values = Vec::with_capacity(numbers.len());
    for expr in numbers {
        values.push(match expr {
            FormulaExpr::Empty => Ok(None),
            expr => whole(expr, ctx)?.map(Some),
        });
    }
    if let Some(Err(e)) = values.iter().take(2).find(|v| v.is_err()) {
        return Ok(Err(*e));
    }
    Ok(match base(reference, ctx)? {
        Ok(target) => offset_range(target, &values),
        Err(e) => Err(e),
    })
}

/// The range OFFSET gives: the base moved by rows and cols, with a left-out height or width
/// keeping the base's own; one below 1, or a range off the sheet, is #REF!.
fn offset_range((sheet, range): Target, numbers: &OffsetNumbers) -> Result<Target, CellError> {
    let number = |i: usize| numbers.get(i).cloned().unwrap_or(Ok(None));
    let (dr, dc) = (number(0)?.unwrap_or(0), number(1)?.unwrap_or(0));
    let extent = |i: usize, own: i64| match number(i)? {
        None => Ok(own),
        Some(n) if n >= 1 => Ok(n),
        Some(_) => Err(CellError::Ref),
    };
    let height = extent(2, (range.end.row - range.start.row + 1) as i64)?;
    let width = extent(3, (range.end.col - range.start.col + 1) as i64)?;
    let r0 = range.start.row as i64 + dr;
    let c0 = range.start.col as i64 + dc;
    if r0 < 0 || c0 < 0 || r0 + height > MAX_ROWS as i64 || c0 + width > MAX_COLS as i64 {
        return Err(CellError::Ref);
    }
    Ok((
        sheet,
        CellRange::from_indices(
            r0 as u32,
            c0 as u16,
            (r0 + height - 1) as u32,
            (c0 + width - 1) as u16,
        ),
    ))
}

/// SUBTOTAL, SUMIF, COUNTIF and the like read their ranges as references (`per_range` in the
/// registry): given OFFSET with an array of offsets or sizes, one result per range OFFSET gives,
/// and SUMIF's sum range takes the shape of its range. `None` when neither applies.
pub(crate) fn evaluate_per_range(
    name: &str,
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> Option<FormulaResult<FormulaValue>> {
    match over_offsets(name, args, ctx).transpose() {
        Some(result) => Some(result),
        None if matches!(name, "SUMIF" | "AVERAGEIF") && args.len() == 3 => {
            resized_sum_range(name, args, ctx)
        }
        None => None,
    }
}

/// `SUMPRODUCT(SUBTOTAL(9,OFFSET(B2:B8,,{0,1,2})))` adds three columns' totals. The OFFSET among
/// the arguments has its arguments evaluated once; the function then runs once per range it
/// gives, or once with the one range. An array at an argument that takes one value (COUNTIF's
/// criteria) broadcasts with the ranges: each call gets its element. `None` when no argument is
/// an OFFSET.
fn over_offsets(
    name: &str,
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<Option<FormulaValue>> {
    use crate::array_lift::{broadcast_at, dims};
    let found = args
        .iter()
        .enumerate()
        .find_map(|(position, arg)| match arg {
            FormulaExpr::Function { name, args }
                if crate::evaluator::function_name(name) == "OFFSET" && args.len() >= 3 =>
            {
                Some((position, args))
            }
            _ => None,
        });
    let Some((position, offset_args)) = found else {
        return Ok(None);
    };
    let mut values = Vec::with_capacity(offset_args.len() - 1);
    for expr in &offset_args[1..] {
        values.push(match expr {
            FormulaExpr::Empty => None,
            expr => Some(evaluate(expr, ctx)?),
        });
    }
    let base = base(&offset_args[0], ctx)?;
    let def = crate::functions::registry().get(name);
    let mut criteria: Vec<(usize, Vec<Vec<FormulaValue>>)> = Vec::new();
    for (k, arg) in args.iter().enumerate() {
        if k != position && def.is_some_and(|def| def.scalar_args.includes_in(k, args.len())) {
            if let FormulaValue::Array { data, .. } = evaluate(arg, ctx)? {
                criteria.push((k, data));
            }
        }
    }
    let shape = criteria
        .iter()
        .map(|(_, data)| dims(data))
        .fold((1, 1), |(rows, cols), (r, c)| (rows.max(r), cols.max(c)));
    let shape = values
        .iter()
        .flatten()
        .fold(shape, |(rows, cols), value| match value {
            FormulaValue::Array { data, .. } => {
                let (r, c) = dims(data);
                (rows.max(r), cols.max(c))
            }
            _ => (rows, cols),
        });
    let mut out = Vec::with_capacity(shape.0);
    for i in 0..shape.0 {
        let mut line = Vec::with_capacity(shape.1);
        for j in 0..shape.1 {
            let numbers: Vec<Result<Option<i64>, CellError>> = values
                .iter()
                .map(|value| match value {
                    None => Ok(None),
                    Some(FormulaValue::Array { data, .. }) => {
                        match broadcast_at(data, dims(data), i, j) {
                            Some(element) => whole_of(element).map(Some),
                            None => Err(CellError::Na),
                        }
                    }
                    Some(value) => whole_of(value).map(Some),
                })
                .collect();
            // As OFFSET itself: an error in rows or cols comes before the base's.
            let target = match numbers.iter().take(2).find_map(|n| n.as_ref().err()) {
                Some(e) => Err(*e),
                None => base.clone().and_then(|base| offset_range(base, &numbers)),
            };
            let mut call = args.to_vec();
            call[position] = match target {
                Ok((sheet, range)) => FormulaExpr::RangeRef(RangeReference { sheet, range }),
                Err(e) => FormulaExpr::Error(e),
            };
            for (k, data) in &criteria {
                call[*k] = match broadcast_at(data, dims(data), i, j) {
                    Some(FormulaValue::Number(n)) => FormulaExpr::Number(*n),
                    Some(FormulaValue::String(text)) => FormulaExpr::String(text.clone()),
                    Some(FormulaValue::Boolean(b)) => FormulaExpr::Boolean(*b),
                    Some(FormulaValue::Error(e)) => FormulaExpr::Error(*e),
                    // A blank criteria cell is 0.
                    Some(_) => FormulaExpr::Number(0.0),
                    None => FormulaExpr::Error(CellError::Na),
                };
            }
            line.push(
                match crate::evaluator::evaluate_function(name, &call, ctx)? {
                    FormulaValue::Array { data, .. } if shape != (1, 1) => data
                        .first()
                        .and_then(|r| r.first())
                        .cloned()
                        .unwrap_or(FormulaValue::Error(CellError::Value)),
                    value => value,
                },
            );
        }
        out.push(line);
    }
    Ok(Some(match shape {
        (1, 1) => out.remove(0).remove(0),
        _ => FormulaValue::Array {
            data: out,
            source: None,
        },
    }))
}

/// INDEX given a reference, a union of them included (`INDEX((A1:B2,C1:D2),1,1,2)`), read; or
/// given an array, the element or row or column it picks.
pub(crate) fn evaluate_index(
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    Ok(match index(args, ctx, true)? {
        Some(Resolved::Reference(target)) => read(ctx, &target),
        Some(Resolved::Value(value)) => value,
        None => FormulaValue::Error(CellError::Value),
    })
}

/// SUMIF and AVERAGEIF read as many cells of sum_range as range has, in the same shape, from
/// sum_range's top-left cell, as Microsoft documents: SUMIF(A1:E17,"x",B1:B17) adds from B1:F17.
/// `None` when the two are the same shape, or aren't references.
fn resized_sum_range(
    name: &str,
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> Option<FormulaResult<FormulaValue>> {
    let Ok(Some(Ok((sheet, sum)))) = reference_of(&args[2], ctx) else {
        return None;
    };
    let Ok(Some(Ok((_, range)))) = reference_of(&args[0], ctx) else {
        return None;
    };
    let (rows, cols) = (
        range.end.row - range.start.row,
        range.end.col - range.start.col,
    );
    if (sum.end.row - sum.start.row, sum.end.col - sum.start.col) == (rows, cols)
        || sum.start.row + rows >= MAX_ROWS
        || sum.start.col + cols >= MAX_COLS
    {
        return None;
    }
    let resized = CellRange::from_indices(
        sum.start.row,
        sum.start.col,
        sum.start.row + rows,
        sum.start.col + cols,
    );
    let sum_range = FormulaExpr::RangeRef(RangeReference {
        sheet,
        range: resized,
    });
    Some(crate::evaluator::evaluate_function(
        name,
        &[args[0].clone(), args[1].clone(), sum_range],
        ctx,
    ))
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
