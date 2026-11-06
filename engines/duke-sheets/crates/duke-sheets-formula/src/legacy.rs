//! Formulas from before dynamic arrays. Excel saves them without the dynamic-array mark and still
//! evaluates them the old way: a range where one value is expected, at an operator or a value
//! argument, reads only its cell in the formula's own row or column, and the formula shows one
//! value. Excel 365 shows such a formula with `@` at those places (`=SUM(@A1:A3*@B1:B3)`), and so
//! does the engine: a reader puts the `@` into the formula text (`read_formulas`), which then keeps
//! its meaning through every edit. Saving takes them out again where the old rules put them back
//! (`saved_formula`). Either way only the `@` change; the rest of the text stays as written.
//!
//! Which arguments take one value comes from Excel's argument classes in the function registry
//! (`FunctionDef::arg_class`). Array arguments keep arrays, as they always did: SUMPRODUCT's,
//! INDEX's first, LOOKUP's vectors (`LOOKUP(2,1/(A1:A9=x),B1:B9)` needs no Ctrl+Shift+Enter).
//! Reference arguments (SUM's, COUNTIF's range) keep the range itself, while the operators inside
//! them still read one cell: `SUM(A1:A3*B1:B3)` multiplies one cell of each.

use std::collections::HashSet;
use std::mem::discriminant;

use crate::ast::{BinaryOperator, FormulaExpr, StructuredRefSpecifier, UnaryOperator};
use crate::evaluator::function_name;
use crate::functions::{registry, ArgClass, FunctionDef, Returns, ScalarArgs};
use crate::parser::{parse_formula, tokens, Token};
use crate::print::{file_formula_text, formula_text};
use duke_sheets_core::{FormulaKind, Workbook};

/// The defined names that hold more than one value: a range of cells, an array, or a calculation.
/// A name for one cell or a constant reads the same either way, so `=TaxRate*A1` needs no `@`.
/// LET and LAMBDA names, and names that aren't defined, aren't among them.
#[derive(Default)]
pub struct Names(HashSet<String>);

impl Names {
    pub fn of(workbook: &Workbook) -> Self {
        let many = workbook
            .named_ranges()
            .iter()
            .filter(|name| {
                let one = match parse_formula(&format!("={}", name.refers_to.trim_start_matches('='))) {
                    Ok(FormulaExpr::RangeRef(r)) => r.range.start == r.range.end,
                    Ok(
                        FormulaExpr::CellRef(_)
                        | FormulaExpr::Number(_)
                        | FormulaExpr::String(_)
                        | FormulaExpr::Boolean(_)
                        | FormulaExpr::Error(_),
                    ) => true,
                    _ => false,
                };
                !one
            })
            .map(|name| name.name.to_ascii_uppercase())
            .collect();
        Names(many)
    }

    fn many(&self, name: &str) -> bool {
        !self.0.is_empty() && self.0.contains(&name.to_ascii_uppercase())
    }
}

/// A formula cell as a file holds it.
pub struct FileFormula {
    pub sheet: usize,
    pub row: u32,
    pub col: u16,
    /// Saved with the dynamic-array mark (`cm="1"`).
    pub dynamic: bool,
    /// The shared formula it belongs to (`si`): its cells' formulas have the same shape.
    pub shared: Option<u32>,
}

/// Spells the formulas a file holds as the editor shows them, after the workbook's names are read.
/// One saved without the dynamic-array mark gets Excel's `@` where the old rules read one value;
/// `_xlfn.SINGLE(x)` and `_xlfn.ANCHORARRAY(x)` become `@x` and `x#`. A Ctrl+Shift+Enter array
/// formula, evaluated as an array, gets no `@`.
pub fn read_formulas(workbook: &mut Workbook, cells: &[FileFormula]) {
    let names = Names::of(workbook);
    // A legacy shared formula whose first cell reads the same needs nothing in its other cells.
    let mut unchanged: HashSet<(usize, u32)> = HashSet::new();
    for cell in cells {
        let Some(formula) = workbook.worksheet_mut(cell.sheet).and_then(|ws| ws.formula_data_at_mut(cell.row, cell.col))
        else {
            continue;
        };
        let legacy = !cell.dynamic && formula.kind == FormulaKind::Dynamic;
        let group = cell.shared.filter(|_| legacy).map(|si| (cell.sheet, si));
        if group.is_some_and(|group| unchanged.contains(&group)) {
            continue;
        }
        match shown(&formula.text, legacy, &names) {
            Some(text) => formula.text = text,
            None => {
                if let Some(group) = group {
                    unchanged.insert(group);
                }
            }
        }
    }
}

/// `read_formulas` for every formula of a .xls file, which has no dynamic-array mark. Its array
/// formulas are kept as `{=…}` text, so only text starting with `=` is a formula here.
pub fn read_all_as_legacy(workbook: &mut Workbook) {
    let cells: Vec<FileFormula> = (0..workbook.sheet_count())
        .filter_map(|sheet| workbook.worksheet(sheet).map(|ws| (sheet, ws)))
        .flat_map(|(sheet, ws)| {
            ws.formula_cells()
                .filter(|(_, _, text)| text.starts_with('='))
                .map(move |(row, col, _)| FileFormula { sheet, row, col, dynamic: false, shared: None })
                .collect::<Vec<_>>()
        })
        .collect();
    read_formulas(workbook, &cells);
}

/// The text a file's formula shows in the editor, or `None` when that is the text as it is.
fn shown(text: &str, legacy: bool, names: &Names) -> Option<String> {
    // Nothing in a formula without ranges, names, calls, arrays, `@` or `#` can need an `@`.
    if names.0.is_empty() && !text.contains([':', '[', '#', '@', '{', '(']) {
        return None;
    }
    let parsed = parse_formula(text).ok()?;
    let marked = has_marks(&parsed);
    if !legacy && !marked {
        return None;
    }
    let mut changed = false;
    let target = if legacy { walk(parsed, Place::One, names, &mut changed) } else { parsed };
    if !changed && !marked {
        return None;
    }
    let from = parse_formula(text).ok()?;
    // `_xlfn.SINGLE(x)` in the text doesn't line up with `@x`: write the formula anew.
    Some(respell(text, &from, &target).unwrap_or_else(|| formula_text(&target))).filter(|shown| shown != text)
}

/// A formula as a file saves it, given with its `=`: the text, and whether it is marked as a
/// dynamic-array formula (`cm="1"`). One whose `@` are all where the old rules put them is saved
/// without them, as a formula Excel 2019 and earlier read the same; any other `@` is saved as
/// `_xlfn.SINGLE` (and `#` as `_xlfn.ANCHORARRAY`), and the mark is set when reading the formula
/// the old way would put in more `@`: a range where one value is expected, or a result that can
/// be an array (FILTER, `A1:A3*2`, INDEX). An `@` on what can't be an array (`=@A1`) is left out,
/// as Excel does.
pub fn saved_formula(text: &str, names: &Names) -> (String, bool) {
    // Nothing in a formula without ranges, names, calls, arrays, `@` or `#` reads differently.
    if names.0.is_empty() && !text.contains([':', '[', '#', '@', '{', '(']) {
        return (text.to_string(), false);
    }
    let Ok(parsed) = parse_formula(text) else {
        return (text.to_string(), false);
    };
    let mut more = false;
    if !has_marks(&parsed) {
        walk(parsed, Place::One, names, &mut more);
        return (text.to_string(), more);
    }
    let expr = without_redundant_at(parsed.clone(), names);
    let plain = without_at(expr.clone());
    let mut unused = false;
    if !has_marks(&plain) && walk(plain.clone(), Place::One, names, &mut unused) == expr {
        return (spelled(text, &parsed, &plain), false);
    }
    walk(expr.clone(), Place::One, names, &mut more);
    let text = if has_marks(&expr) { file_formula_text(&expr) } else { spelled(text, &parsed, &expr) };
    (text, more)
}

/// A Ctrl+Shift+Enter array formula's text as a file saves it: `@` and `#` as `_xlfn.SINGLE` and
/// `_xlfn.ANCHORARRAY`, and nothing else changed.
pub fn saved_array_formula(text: &str) -> String {
    match parse_formula(text) {
        Ok(expr) if has_marks(&expr) => file_formula_text(&expr),
        _ => text.to_string(),
    }
}

/// `text`, parsed as `from`, spelled as `to`, which differs from it only by `@`; written anew when
/// the two don't line up that way.
fn spelled(text: &str, from: &FormulaExpr, to: &FormulaExpr) -> String {
    respell(text, from, to).unwrap_or_else(|| formula_text(to))
}

/// `text` (parsed as `from`) with the `@` that `to` adds put in before their tokens and those it
/// drops taken out, the rest left as written: spaces, line breaks, names in their own case,
/// numbers as typed. `None` when `to` differs from `from` by more than `@`.
fn respell(text: &str, from: &FormulaExpr, to: &FormulaExpr) -> Option<String> {
    let original = tokens(text)?;
    let before = tokens(&formula_text(from))?;
    let after = tokens(&formula_text(to))?;
    let same = |a: &Token, b: &Token| discriminant(a) == discriminant(b);
    if original.len() != before.len() || original.iter().zip(&before).any(|(o, b)| !same(&o.1, &b.1)) {
        return None;
    }
    let mut out = String::with_capacity(text.len() + 8);
    let mut copied = 0;
    let (mut i, mut j) = (0, 0);
    while i < before.len() || j < after.len() {
        match (before.get(i).map(|t| &t.1), after.get(j).map(|t| &t.1)) {
            (b, Some(Token::At)) if b != Some(&Token::At) => {
                let at = original.get(i).map_or(text.len(), |t| t.0.start);
                out.push_str(&text[copied..at]);
                out.push('@');
                copied = at;
                j += 1;
            }
            (Some(Token::At), a) if a != Some(&Token::At) => {
                let span = &original[i].0;
                out.push_str(&text[copied..span.start]);
                copied = span.end;
                i += 1;
            }
            (Some(b), Some(a)) if same(a, b) => {
                i += 1;
                j += 1;
            }
            _ => return None,
        }
    }
    out.push_str(&text[copied..]);
    Some(out)
}

fn has_marks(expr: &FormulaExpr) -> bool {
    let mut found = false;
    visit(expr, &mut |e| {
        found |= matches!(e, FormulaExpr::UnaryOp { op: UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange, .. });
    });
    found
}

fn visit(expr: &FormulaExpr, f: &mut dyn FnMut(&FormulaExpr)) {
    f(expr);
    match expr {
        FormulaExpr::BinaryOp { left, right, .. } => {
            visit(left, f);
            visit(right, f);
        }
        FormulaExpr::UnaryOp { operand, .. } => visit(operand, f),
        FormulaExpr::Function { args, .. } | FormulaExpr::ExternalFunction { args, .. } => {
            args.iter().for_each(|a| visit(a, f))
        }
        FormulaExpr::Array(rows) => rows.iter().flatten().for_each(|a| visit(a, f)),
        _ => {}
    }
}

/// `expr` with `f` applied to each of its direct parts.
fn map_parts(expr: FormulaExpr, f: &mut dyn FnMut(FormulaExpr) -> FormulaExpr) -> FormulaExpr {
    match expr {
        FormulaExpr::UnaryOp { op, operand } => FormulaExpr::UnaryOp { op, operand: Box::new(f(*operand)) },
        FormulaExpr::BinaryOp { op, left, right } => {
            FormulaExpr::BinaryOp { op, left: Box::new(f(*left)), right: Box::new(f(*right)) }
        }
        FormulaExpr::Function { name, args } => FormulaExpr::Function { name, args: args.into_iter().map(&mut *f).collect() },
        FormulaExpr::ExternalFunction { book, name, args } => {
            FormulaExpr::ExternalFunction { book, name, args: args.into_iter().map(&mut *f).collect() }
        }
        FormulaExpr::Array(rows) => {
            FormulaExpr::Array(rows.into_iter().map(|row| row.into_iter().map(&mut *f).collect()).collect())
        }
        other => other,
    }
}

fn without_at(expr: FormulaExpr) -> FormulaExpr {
    match expr {
        FormulaExpr::UnaryOp { op: UnaryOperator::ImplicitIntersection, operand } => without_at(*operand),
        other => map_parts(other, &mut without_at),
    }
}

/// Leaves out each `@` on what can only be one value: `@A1`, `@SUM(A1:A3)`.
fn without_redundant_at(expr: FormulaExpr, names: &Names) -> FormulaExpr {
    match expr {
        FormulaExpr::UnaryOp { op: UnaryOperator::ImplicitIntersection, operand } if !could_be_array(&operand, names) => {
            without_redundant_at(*operand, names)
        }
        other => map_parts(other, &mut |part| without_redundant_at(part, names)),
    }
}

/// The formula as Excel 365 shows it when it was saved without the dynamic-array mark: `@`
/// where the old rules read one value. The `@` already there stay.
pub fn intersections(expr: FormulaExpr, names: &Names) -> FormulaExpr {
    let mut changed = false;
    walk(expr, Place::One, names, &mut changed)
}

/// What a place in a formula takes.
#[derive(Clone, Copy, PartialEq)]
enum Place {
    /// One value: a range there reads one cell.
    One,
    /// A reference, which stays whole; the operators inside it read one value.
    Whole,
    /// An array, as in a dynamic-array formula. A condition still reads one value there:
    /// `SUMPRODUCT(IF(A1:A3>1,1,0))` needs Ctrl+Shift+Enter.
    Array,
}

/// The place a function's argument is, the call itself being at `call`: one its result takes
/// (IF's value_if_true is where IF is), else from the argument's class.
fn place_of(def: &FunctionDef, position: usize, call: Place) -> Place {
    if def.result_args.includes(position) {
        return call;
    }
    // An argument that takes each element of an array in a dynamic-array formula takes one
    // value in an old one.
    if def.scalar_args.includes(position) {
        return Place::One;
    }
    match def.arg_class(position) {
        ArgClass::V => Place::One,
        ArgClass::R => Place::Whole,
        ArgClass::A | ArgClass::RA => Place::Array,
    }
}

fn at(expr: FormulaExpr, changed: &mut bool) -> FormulaExpr {
    *changed = true;
    FormulaExpr::UnaryOp { op: UnaryOperator::ImplicitIntersection, operand: Box::new(expr) }
}

fn walk(expr: FormulaExpr, place: Place, names: &Names, changed: &mut bool) -> FormulaExpr {
    if place == Place::Array {
        return array_place(expr, names, changed);
    }
    match expr {
        FormulaExpr::BinaryOp { op: op @ (BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect), left, right } => {
            let reference = FormulaExpr::BinaryOp {
                op,
                left: Box::new(walk(*left, Place::Whole, names, changed)),
                right: Box::new(walk(*right, Place::Whole, names, changed)),
            };
            one_value(reference, place, names, changed)
        }
        FormulaExpr::BinaryOp { op, left, right } => FormulaExpr::BinaryOp {
            op,
            left: Box::new(walk(*left, Place::One, names, changed)),
            right: Box::new(walk(*right, Place::One, names, changed)),
        },
        FormulaExpr::UnaryOp { op: op @ (UnaryOperator::Negate | UnaryOperator::Plus | UnaryOperator::Percent), operand } => {
            FormulaExpr::UnaryOp { op, operand: Box::new(walk(*operand, Place::One, names, changed)) }
        }
        // `(A1,B1)` and `(A1:A3)` take the `@` outside their parentheses.
        FormulaExpr::UnaryOp { op: UnaryOperator::Paren, operand } => {
            let inner = FormulaExpr::UnaryOp {
                op: UnaryOperator::Paren,
                operand: Box::new(walk(*operand, Place::Whole, names, changed)),
            };
            one_value(inner, place, names, changed)
        }
        FormulaExpr::Function { name, args } => {
            let def = registry().get(function_name(&name));
            let args = args
                .into_iter()
                .enumerate()
                .map(|(position, arg)| {
                    let takes = def.map_or(Place::Whole, |def| place_of(def, position, place));
                    walk(arg, takes, names, changed)
                })
                .collect();
            one_value(FormulaExpr::Function { name, args }, place, names, changed)
        }
        reference @ (FormulaExpr::RangeRef(_)
        | FormulaExpr::NameRef(_)
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::UnaryOp { op: UnaryOperator::SpillRange, .. }) => one_value(reference, place, names, changed),
        other => other,
    }
}

/// `@expr` where one value is expected and `expr` can give more.
fn one_value(expr: FormulaExpr, place: Place, names: &Names, changed: &mut bool) -> FormulaExpr {
    if place == Place::One && could_be_array(&expr, names) {
        at(expr, changed)
    } else {
        expr
    }
}

/// An array place leaves its contents as they are, but for the arguments that decide which
/// result a function hands back (IF's condition, CHOOSE's index), which read one value.
fn array_place(expr: FormulaExpr, names: &Names, changed: &mut bool) -> FormulaExpr {
    match expr {
        FormulaExpr::Function { name, args } => {
            let decides = registry()
                .get(function_name(&name))
                .map_or(ScalarArgs::None, |def| def.result_args);
            let args = args
                .into_iter()
                .enumerate()
                .map(|(position, arg)| {
                    if decides != ScalarArgs::None && !decides.includes(position) {
                        walk(arg, Place::One, names, changed)
                    } else {
                        array_place(arg, names, changed)
                    }
                })
                .collect();
            FormulaExpr::Function { name, args }
        }
        other => map_parts(other, &mut |part| array_place(part, names, changed)),
    }
}

/// Whether `expr` can give more than one value: a range, a name for one, an array, a function
/// that returns arrays or references (FILTER, INDEX, OFFSET), or one given such a value where it
/// takes one (`LEN(A1:A3)`). `[@Qty]`, `@A1:A3`, a cell and SUM(A1:A3) can't.
pub fn could_be_array(expr: &FormulaExpr, names: &Names) -> bool {
    match expr {
        FormulaExpr::RangeRef(r) => r.range.start != r.range.end,
        FormulaExpr::NameRef(name) => names.many(name),
        FormulaExpr::StructuredRef(sr) => !sr.specifiers.contains(&StructuredRefSpecifier::ThisRow),
        FormulaExpr::Array(rows) => rows.len() > 1 || rows.first().is_some_and(|row| row.len() > 1),
        FormulaExpr::BinaryOp { op: BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect, .. } => true,
        FormulaExpr::BinaryOp { left, right, .. } => could_be_array(left, names) || could_be_array(right, names),
        FormulaExpr::UnaryOp { op: UnaryOperator::ImplicitIntersection, .. } => false,
        FormulaExpr::UnaryOp { op: UnaryOperator::SpillRange, .. } => true,
        FormulaExpr::UnaryOp { operand, .. } => could_be_array(operand, names),
        FormulaExpr::ExternalFunction { .. } => true,
        FormulaExpr::Function { name, args } => match registry().get(function_name(name)) {
            // A LAMBDA called in place, or a function Excel doesn't have (an add-in's).
            None => true,
            Some(def) => match def.returns {
                Returns::Array | Returns::Reference => true,
                Returns::PerCell => args.iter().any(|arg| could_be_array(arg, names)),
                Returns::Last => args.last().is_some_and(|last| could_be_array(last, names)),
                Returns::Value => args.iter().enumerate().any(|(position, arg)| {
                    place_of(def, position, Place::One) == Place::One && could_be_array(arg, names)
                }),
            },
        },
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn names() -> Names {
        Names(["PRICES".to_string()].into_iter().collect())
    }

    fn shown_legacy(text: &str) -> String {
        shown(text, true, &names()).unwrap_or_else(|| text.to_string())
    }

    #[test]
    fn ranges_where_one_value_is_expected_get_an_at() {
        assert_eq!(shown_legacy("=A1:A3"), "=@A1:A3");
        assert_eq!(shown_legacy("=SUM(A1:A3*B1:B3)"), "=SUM(@A1:A3*@B1:B3)");
        assert_eq!(shown_legacy("=SUMIFS(B:B,A:A,C1:C3)"), "=SUMIFS(B:B,A:A,@C1:C3)");
        assert_eq!(shown_legacy("=NETWORKDAYS(A:A,B:B)"), "=NETWORKDAYS(@A:A,@B:B)");
        assert_eq!(shown_legacy("=MAX(IF(A1:A9=1,B1:B9))"), "=MAX(IF(@A1:A9=1,B1:B9))");
        assert_eq!(shown_legacy("=IF(A1>1,B1:B3,C1)+0"), "=IF(A1>1,@B1:B3,C1)+0");
        assert_eq!(shown_legacy("=INDEX(A:A,5)"), "=@INDEX(A:A,5)");
        assert_eq!(shown_legacy("=Prices*2"), "=@Prices*2");
        assert_eq!(shown_legacy("=(A1,B1)"), "=@(A1,B1)");
        assert_eq!(shown_legacy("=SUMPRODUCT(IF(A1:A3>1,1,0))"), "=SUMPRODUCT(IF(@A1:A3>1,1,0))");
        assert_eq!(shown_legacy("=SUM(MID(A1,ROW(A1:A9),1)*1)"), "=SUM(MID(A1,@ROW(A1:A9),1)*1)");
        assert_eq!(shown_legacy("=SUM(A2#*B2#)"), "=SUM(@A2#*@B2#)");
        assert_eq!(shown_legacy("=A1:A3%"), "=@A1:A3%");
    }

    #[test]
    fn only_the_at_change_in_the_text() {
        assert_eq!(shown_legacy("= SUM( A1:A3 * b1:b3 )\n+MyUdf(1E+20,[1]$A$1)"), "= SUM( @A1:A3 * @b1:b3 )\n+@MyUdf(1E+20,[1]$A$1)");
        let (saved, dynamic) = saved_formula("= SUM( @A1:A3 * @b1:b3 )", &names());
        assert_eq!((saved.as_str(), dynamic), ("= SUM( A1:A3 * b1:b3 )", false));
    }

    #[test]
    fn references_and_arrays_stay_whole() {
        for text in [
            "=SUM(A1:A3)+B1",
            "=SUMPRODUCT((A1:A9>1)*B1:B9)",
            "=LOOKUP(2,1/(A1:A9=1),B1:B9)",
            "=VLOOKUP(A1,C:D,2,FALSE)",
            "=SUM(IF(A1>1,B1:B3))",
            "=SUM(SUMIF(A1:A9,{1,2},B1:B9))",
            "=@INDEX(B:B,MATCH(1,INDEX((A1:A9=1)*(C1:C9=2),0),0))",
            "=SUM(FILTER(A1:A9,A1:A9>1))",
            "=[@Qty]*[@Price]",
            "=TaxRate*A1",
            "=LET(x,A1:A3,SUM(IF(x>1,1)))",
        ] {
            assert_eq!(shown_legacy(text), text);
        }
    }

    #[test]
    fn what_the_rules_build_prints_and_parses_back() {
        for text in ["=A1:A3%", "=SUM(A2#*B2#)", "=-A1:A3", "=(A1:A3)*2+(B1:B3*2)", "=A1:B3 B2:C4", "=INDEX(A:A,5)%"] {
            let legacy = intersections(parse_formula(text).unwrap(), &names());
            assert_eq!(parse_formula(&formula_text(&legacy)).unwrap(), legacy, "{text}");
        }
    }

    #[test]
    fn saving_marks_formulas_that_read_differently() {
        let names = names();
        let saved = |text: &str| saved_formula(text, &names);
        assert_eq!(saved("=SUM(A1:A3*B1:B3)"), ("=SUM(A1:A3*B1:B3)".to_string(), true));
        assert_eq!(saved("=SUM(@A1:A3*@B1:B3)"), ("=SUM(A1:A3*B1:B3)".to_string(), false));
        assert_eq!(saved("=SUM(@A1:A3)"), ("=SUM(_xlfn.SINGLE(A1:A3))".to_string(), false));
        assert_eq!(saved("=@A1+@SUM(A1:A3)"), ("=A1+SUM(A1:A3)".to_string(), false));
        assert_eq!(saved("=FILTER(A1:A3,A1:A3>1)"), ("=FILTER(A1:A3,A1:A3>1)".to_string(), true));
        assert!(saved("=SUM(OFFSET(A1,0,0,3,1)*B1)").1);
        assert!(saved("=SUM(--(A2#>5))").1);
        assert!(saved("=MAX(A2#-B1)").1);
        assert!(!saved("=[@Qty]*[@Price]").1);
        assert!(!saved("=TaxRate*A1").1);
        assert!(!saved("=VLOOKUP(A1,C:D,2,FALSE)").1);
    }
}
