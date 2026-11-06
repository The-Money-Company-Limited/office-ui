//! LET, LAMBDA and the helpers that call a LAMBDA for each element, row or column (MAP, BYROW,
//! BYCOL, REDUCE, SCAN, MAKEARRAY). They are evaluated from the formula itself, before the
//! function registry, because their names and bodies are not values.

use std::sync::Arc;

use crate::ast::FormulaExpr;
use crate::error::FormulaResult;
use crate::evaluator::{evaluate, to_array, EvaluationContext, FormulaValue};
use duke_sheets_core::{CellError, MAX_COLS, MAX_ROWS};

/// A name bound by LET or a LAMBDA parameter, visible to the expressions inside it.
pub struct Local<'a> {
    name: String,
    binding: Binding,
    outer: Option<&'a Local<'a>>,
}

/// What a local name stands for: a value, a reference (`LET(r,A1:A3,ROW(r))`), or a LAMBDA that
/// can be called by that name.
#[derive(Clone)]
pub enum Binding {
    Value(FormulaValue),
    Reference(crate::reference::Target),
    Lambda(Arc<Lambda>),
}

/// A LAMBDA with the names it could see where it was written. A LET name is bound after its value
/// is evaluated, so a LAMBDA can't call itself by its own name (#NAME? in Excel).
pub struct Lambda {
    params: Vec<String>,
    body: FormulaExpr,
    captured: Vec<(String, Binding)>,
}

/// Excel writes LET and LAMBDA names as `_xlpm.name`; they compare without case.
pub(crate) fn local_name(name: &str) -> String {
    let upper = name.to_ascii_uppercase();
    upper.strip_prefix("_XLPM.").map(str::to_string).unwrap_or(upper)
}

impl<'a> EvaluationContext<'a> {
    /// The LET or LAMBDA binding of `name` in scope, innermost first.
    pub fn local(&self, name: &str) -> Option<&Binding> {
        let name = local_name(name);
        let mut scope = self.locals;
        while let Some(local) = scope {
            if local.name == name {
                return Some(&local.binding);
            }
            scope = local.outer;
        }
        None
    }

    /// This context with `locals` as the names in scope.
    fn with_locals<'b>(&'b self, locals: Option<&'b Local<'b>>) -> EvaluationContext<'b> {
        EvaluationContext { locals, ..*self }
    }

    /// The names in scope, outermost first.
    fn scope(&self) -> Vec<(String, Binding)> {
        let mut names = Vec::new();
        let mut scope = self.locals;
        while let Some(local) = scope {
            names.push((local.name.clone(), local.binding.clone()));
            scope = local.outer;
        }
        names.reverse();
        names
    }
}

/// Evaluates `body` with `bindings` in scope on top of `outer`, the last binding innermost.
fn evaluate_with(
    bindings: &[(String, Binding)],
    outer: Option<&Local>,
    body: &FormulaExpr,
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    match bindings.split_first() {
        Some(((name, binding), rest)) => {
            let local = Local { name: name.clone(), binding: binding.clone(), outer };
            evaluate_with(rest, Some(&local), body, ctx)
        }
        None => evaluate(body, &ctx.with_locals(outer)),
    }
}

/// LET(name1, value1, …, calculation): each value can use the names before it, and a value that
/// is a LAMBDA makes its name callable.
pub(crate) fn evaluate_let(args: &[FormulaExpr], ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.len() < 3 || args.len() % 2 == 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let FormulaExpr::NameRef(name) = &args[0] else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let binding = match lambda_of(&args[1], ctx) {
        Some(lambda) => Binding::Lambda(lambda),
        None => match crate::reference::reference_of(&args[1], ctx)? {
            Some(Ok(target)) => Binding::Reference(target),
            _ => Binding::Value(evaluate(&args[1], ctx)?),
        },
    };
    let local = Local { name: local_name(name), binding, outer: ctx.locals };
    let inner = ctx.with_locals(Some(&local));
    if args.len() == 3 {
        evaluate(&args[2], &inner)
    } else {
        evaluate_let(&args[2..], &inner)
    }
}

/// A LAMBDA written in place, capturing the names in scope, or one bound to a LET name.
pub(crate) fn lambda_of(expr: &FormulaExpr, ctx: &EvaluationContext) -> Option<Arc<Lambda>> {
    match expr {
        FormulaExpr::Function { name, args } if crate::evaluator::function_name(name) == "LAMBDA" => {
            let (body, params) = args.split_last()?;
            let params = params
                .iter()
                .map(|p| match p {
                    FormulaExpr::NameRef(n) => Some(local_name(n)),
                    _ => None,
                })
                .collect::<Option<Vec<_>>>()?;
            Some(Arc::new(Lambda { params, body: body.clone(), captured: ctx.scope() }))
        }
        FormulaExpr::NameRef(name) => match ctx.local(name) {
            Some(Binding::Lambda(lambda)) => Some(lambda.clone()),
            _ => None,
        },
        _ => None,
    }
}

/// Calls a LAMBDA: its parameters on top of the names it captured. A wrong number of values is
/// #VALUE!.
pub(crate) fn call_lambda(
    lambda: &Lambda,
    values: &[FormulaValue],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    if lambda.params.len() != values.len() {
        return Ok(FormulaValue::Error(CellError::Value));
    }
    let mut bindings = lambda.captured.clone();
    bindings.extend(lambda.params.iter().cloned().zip(values.iter().cloned().map(Binding::Value)));
    evaluate_with(&bindings, None, &lambda.body, ctx)
}

/// A single result for one element: the top-left value of an array result.
fn single(value: FormulaValue) -> FormulaValue {
    match value {
        FormulaValue::Array { data, .. } => data
            .into_iter()
            .next()
            .and_then(|row| row.into_iter().next())
            .unwrap_or(FormulaValue::Error(CellError::Calc)),
        other => other,
    }
}

/// MAP, BYROW, BYCOL, REDUCE, SCAN and MAKEARRAY call their LAMBDA for each element, row or column.
pub(crate) fn evaluate_lambda_helper(
    name: &str,
    args: &[FormulaExpr],
    ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    let Some((last, inputs)) = args.split_last() else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let Some(lambda) = lambda_of(last, ctx) else {
        return Ok(FormulaValue::Error(CellError::Value));
    };
    let call = |values: &[FormulaValue]| call_lambda(&lambda, values, ctx).map(single);
    let array = |data: Vec<Vec<FormulaValue>>| Ok(FormulaValue::Array { data, source: None });
    let grid = |expr: &FormulaExpr| evaluate(expr, ctx).map(|v| to_array(&v));
    match name {
        "MAP" => {
            let grids = inputs.iter().map(grid).collect::<FormulaResult<Vec<_>>>()?;
            let Some(first) = grids.first() else {
                return Ok(FormulaValue::Error(CellError::Value));
            };
            let mut out = Vec::with_capacity(first.len());
            for (i, row) in first.iter().enumerate() {
                let mut line = Vec::with_capacity(row.len());
                for j in 0..row.len() {
                    let values: Vec<FormulaValue> = grids
                        .iter()
                        .map(|g| g.get(i).and_then(|r| r.get(j)).cloned().unwrap_or(FormulaValue::Error(CellError::Na)))
                        .collect();
                    line.push(call(&values)?);
                }
                out.push(line);
            }
            array(out)
        }
        "BYROW" => {
            let Some(input) = inputs.first() else {
                return Ok(FormulaValue::Error(CellError::Value));
            };
            let mut out = Vec::new();
            for row in grid(input)? {
                out.push(vec![call(&[FormulaValue::Array { data: vec![row], source: None }])?]);
            }
            array(out)
        }
        "BYCOL" => {
            let Some(input) = inputs.first() else {
                return Ok(FormulaValue::Error(CellError::Value));
            };
            let data = grid(input)?;
            let cols = data.first().map_or(0, Vec::len);
            let mut out = Vec::with_capacity(cols);
            for j in 0..cols {
                let column = data.iter().map(|row| vec![row.get(j).cloned().unwrap_or(FormulaValue::Empty)]).collect();
                out.push(call(&[FormulaValue::Array { data: column, source: None }])?);
            }
            array(vec![out])
        }
        "REDUCE" | "SCAN" => {
            let (Some(init), Some(input)) = (inputs.first(), inputs.get(1)) else {
                return Ok(FormulaValue::Error(CellError::Value));
            };
            let mut accumulator = evaluate(init, ctx)?;
            let mut out = Vec::new();
            for row in grid(input)? {
                let mut line = Vec::with_capacity(row.len());
                for value in row {
                    accumulator = call(&[accumulator, value])?;
                    line.push(accumulator.clone());
                }
                out.push(line);
            }
            if name == "REDUCE" {
                Ok(accumulator)
            } else {
                array(out)
            }
        }
        "MAKEARRAY" => {
            // MAKEARRAY(rows, cols, LAMBDA(r, c, …)), at most a sheet's rows and columns
            let size = |expr: Option<&FormulaExpr>, max: f64| -> FormulaResult<Option<usize>> {
                Ok(match expr {
                    Some(e) => evaluate(e, ctx)?.as_number().filter(|n| (1.0..=max).contains(n)).map(|n| n as usize),
                    None => None,
                })
            };
            let (Some(rows), Some(cols)) = (size(inputs.first(), MAX_ROWS as f64)?, size(inputs.get(1), MAX_COLS as f64)?)
            else {
                return Ok(FormulaValue::Error(CellError::Value));
            };
            let mut out = Vec::with_capacity(rows);
            for r in 1..=rows {
                let mut line = Vec::with_capacity(cols);
                for c in 1..=cols {
                    line.push(call(&[FormulaValue::Number(r as f64), FormulaValue::Number(c as f64)])?);
                }
                out.push(line);
            }
            array(out)
        }
        _ => Ok(FormulaValue::Error(CellError::Value)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        evaluate(&ast, &EvaluationContext::simple())
    }

    #[test]
    fn let_binds_names() {
        // https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
        assert_eq!(eval("=LET(x,5,SUM(x,1))").unwrap(), FormulaValue::Number(6.0));
        // Later names see earlier ones; names ignore case; Excel's file form is _xlpm.name
        assert_eq!(eval("=LET(a,10,b,a*2,A+b)").unwrap(), FormulaValue::Number(30.0));
        assert_eq!(eval("=LET(_xlpm.x,2,_xlpm.x^3)").unwrap(), FormulaValue::Number(8.0));
        // A name shadows an outer one
        assert_eq!(eval("=LET(x,1,LET(x,2,x)+x)").unwrap(), FormulaValue::Number(3.0));
        // Errors in a value come through where the name is used
        assert_eq!(eval("=LET(x,1/0,x+1)").unwrap(), FormulaValue::Error(CellError::Div0));
        assert_eq!(eval("=LET(x,1/0,42)").unwrap(), FormulaValue::Number(42.0));
        // A quoted name, or a missing calculation, is #VALUE!
        assert_eq!(eval("=LET(\"x\",5,5+1)").unwrap(), FormulaValue::Error(CellError::Value));
        assert_eq!(eval("=LET(a,1,b,2)").unwrap(), FormulaValue::Error(CellError::Value));
    }

    #[test]
    fn lambdas_close_over_names_and_cannot_recurse() {
        assert_eq!(eval("=LET(sq,LAMBDA(n,n*n),sq(4)+sq(2))").unwrap(), FormulaValue::Number(20.0));
        assert_eq!(eval("=LET(k,3,f,LAMBDA(x,x*k),f(2))").unwrap(), FormulaValue::Number(6.0));
        // A LAMBDA on its own is #CALC!; the wrong number of values is #VALUE!
        assert_eq!(eval("=LET(sq,LAMBDA(n,n*n),sq)").unwrap(), FormulaValue::Error(CellError::Calc));
        assert_eq!(eval("=LET(sq,LAMBDA(n,n*n),sq(1,2))").unwrap(), FormulaValue::Error(CellError::Value));
        // Its own name isn't in scope inside it, so recursion is #NAME? rather than a stack overflow
        assert_eq!(
            eval("=LET(f,LAMBDA(n,IF(n<2,1,n*f(n-1))),f(5))").unwrap(),
            FormulaValue::Error(CellError::Name)
        );
    }

    #[test]
    fn lambda_helpers_call_per_element() {
        let sum = |f: &str| match eval(f).unwrap() {
            FormulaValue::Array { data, .. } => data.iter().flatten().filter_map(|v| v.as_number()).sum::<f64>(),
            other => other.as_number().unwrap(),
        };
        assert_eq!(sum("=MAP({1,2,3},LAMBDA(v,v*2))"), 12.0);
        assert_eq!(sum("=MAP({1,2},{10,20},LAMBDA(a,b,a+b))"), 33.0);
        assert_eq!(sum("=BYROW({1,2;3,4},LAMBDA(r,MAX(r)))"), 6.0);
        assert_eq!(sum("=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))"), 10.0);
        assert_eq!(sum("=REDUCE(0,{1,2,3},LAMBDA(a,v,a+v))"), 6.0);
        assert_eq!(sum("=SCAN(0,{1,2,3},LAMBDA(a,v,a+v))"), 10.0);
        assert_eq!(sum("=MAKEARRAY(2,2,LAMBDA(r,c,r*c))"), 9.0);
        assert_eq!(eval("=MAKEARRAY(1E+20,1,LAMBDA(r,c,1))").unwrap(), FormulaValue::Error(CellError::Value));
        // A LAMBDA called where it is written
        assert_eq!(eval("=LAMBDA(x,x+1)(2)").unwrap(), FormulaValue::Number(3.0));
        assert_eq!(eval("=LAMBDA(a,b,a*b)(3,4)").unwrap(), FormulaValue::Number(12.0));
    }
}
