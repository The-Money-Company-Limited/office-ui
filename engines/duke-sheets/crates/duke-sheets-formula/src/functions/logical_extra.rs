use crate::error::FormulaResult;
use crate::evaluator::{EvaluationContext, FormulaValue};
use duke_sheets_core::CellError;

pub fn fn_let(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // LET(name1, value1, ..., nameN, valueN, calculation)
    // Args come in pre-evaluated. The last arg is the calculation result.
    // Must have odd number of args >= 3 (pairs + final calculation)
    if args.len() < 3 || args.len() % 2 == 0 {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    for (i, arg) in args.iter().enumerate() {
        if i % 2 == 1 && i < args.len() - 1 {
            if let FormulaValue::Error(e) = arg {
                return Ok(FormulaValue::Error(*e));
            }
        }
    }

    Ok(args.last().unwrap().clone())
}

pub fn fn_lambda(args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    if args.is_empty() {
        return Ok(FormulaValue::Error(CellError::Value));
    }

    Ok(args.last().unwrap().clone())
}

pub fn fn_map(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_reduce(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_scan(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_bycol(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_byrow(_args: &[FormulaValue], _ctx: &EvaluationContext) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

pub fn fn_makearray(
    _args: &[FormulaValue],
    _ctx: &EvaluationContext,
) -> FormulaResult<FormulaValue> {
    // TODO: Requires evaluator-level lazy evaluation / lambda capture support.
    // Currently returns #N/A. Full implementation needs FormulaExpr-level changes
    // to pass unevaluated lambda bodies and invoke them per-element.
    Ok(FormulaValue::Error(CellError::Na))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn eval(formula: &str) -> FormulaResult<FormulaValue> {
        let ast = crate::parser::parse_formula(formula)?;
        crate::evaluator::evaluate(&ast, &EvaluationContext::simple())
    }

    #[test]
    fn test_eval_helper_works() {
        let result = eval("=1+1").unwrap();
        assert_eq!(result, FormulaValue::Number(2.0));
    }

    #[test]
    fn test_let_passthrough() {
        let ctx = EvaluationContext::simple();
        let result = fn_let(
            &[
                FormulaValue::String("x".into()),
                FormulaValue::Number(5.0),
                FormulaValue::Number(5.0),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Number(5.0));
    }

    #[test]
    fn test_lambda_passthrough() {
        let ctx = EvaluationContext::simple();
        let result = fn_lambda(
            &[FormulaValue::String("x".into()), FormulaValue::Number(42.0)],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Number(42.0));
    }

    #[test]
    fn test_map_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_map(
            &[FormulaValue::Number(1.0), FormulaValue::Number(2.0)],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_reduce_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_reduce(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Array {
                    data: vec![vec![FormulaValue::Number(1.0)]],
                    source: None,
                },
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_scan_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_scan(
            &[
                FormulaValue::Number(0.0),
                FormulaValue::Array {
                    data: vec![vec![FormulaValue::Number(1.0)]],
                    source: None,
                },
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_bycol_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_bycol(
            &[
                FormulaValue::Array {
                    data: vec![vec![FormulaValue::Number(1.0)]],
                    source: None,
                },
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_byrow_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_byrow(
            &[
                FormulaValue::Array {
                    data: vec![vec![FormulaValue::Number(1.0)]],
                    source: None,
                },
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_makearray_returns_na() {
        let ctx = EvaluationContext::simple();
        let result = fn_makearray(
            &[
                FormulaValue::Number(2.0),
                FormulaValue::Number(3.0),
                FormulaValue::String("lambda".into()),
            ],
            &ctx,
        )
        .unwrap();
        assert_eq!(result, FormulaValue::Error(CellError::Na));
    }

    #[test]
    fn test_let_docs() {
        // https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999
        // Docs Example 1: =LET(x, 5, SUM(x, 1)) → 6
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
        // A LAMBDA bound by LET is callable by its name
        assert_eq!(eval("=LET(sq,LAMBDA(n,n*n),sq(4)+sq(2))").unwrap(), FormulaValue::Number(20.0));
    }

    #[test]
    fn test_lambda_helpers() {
        let sum = |f: &str| match eval(f).unwrap() {
            FormulaValue::Array { data, .. } => data.iter().flatten().filter_map(|v| v.as_number()).sum::<f64>(),
            other => other.as_number().unwrap(),
        };
        assert_eq!(sum("=MAP({1,2,3},LAMBDA(v,v*2))"), 12.0);
        assert_eq!(sum("=BYROW({1,2;3,4},LAMBDA(r,MAX(r)))"), 6.0);
        assert_eq!(sum("=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))"), 10.0);
        assert_eq!(sum("=REDUCE(0,{1,2,3},LAMBDA(a,v,a+v))"), 6.0);
        assert_eq!(sum("=SCAN(0,{1,2,3},LAMBDA(a,v,a+v))"), 10.0);
        assert_eq!(sum("=MAKEARRAY(2,2,LAMBDA(r,c,r*c))"), 9.0);
    }

    #[test]
    fn test_lambda_docs() {
        // https://support.microsoft.com/en-us/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67
        // LAMBDA is a pass-through: returns the last argument.
        let ctx = EvaluationContext::simple();

        // Docs Step 2: =LAMBDA(number, number + 1)(1) => 2
        // Pass-through returns last arg
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("number".into()),
                    FormulaValue::Number(6.0)
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Number(6.0)
        );

        // Example 1 (ToCelsius): two args => returns calculation (last arg)
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("temp".into()),
                    FormulaValue::Number(40.0)
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Number(40.0)
        );

        // Example 2 (Hypotenuse): three args => returns last
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("a".into()),
                    FormulaValue::String("b".into()),
                    FormulaValue::Number(5.0),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Number(5.0)
        );

        // Single arg = just the calculation, no parameters
        assert_eq!(
            fn_lambda(&[FormulaValue::Number(42.0)], &ctx).unwrap(),
            FormulaValue::Number(42.0)
        );

        // Boolean calculation result
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("x".into()),
                    FormulaValue::Boolean(true)
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Boolean(true)
        );

        // String calculation result
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("text".into()),
                    FormulaValue::String("hello".into()),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::String("hello".into())
        );

        // Error propagation: if calculation is an error, return it
        assert_eq!(
            fn_lambda(
                &[
                    FormulaValue::String("x".into()),
                    FormulaValue::Error(CellError::Value),
                ],
                &ctx,
            )
            .unwrap(),
            FormulaValue::Error(CellError::Value)
        );

        // Empty args => #VALUE!
        assert_eq!(
            fn_lambda(&[], &ctx).unwrap(),
            FormulaValue::Error(CellError::Value)
        );
    }
}
