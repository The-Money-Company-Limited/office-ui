//! Formulas from before dynamic arrays. Excel saves them without the dynamic-array mark and still
//! evaluates them the old way: a range where one value is expected, at an operator or a value
//! argument, reads only its cell in the formula's own row or column, and the formula shows one
//! value. Excel 365 shows such a formula with `@` at those places (`=SUM(@A1:A3*@B1:B3)`), and so
//! does the engine: a reader puts the `@` into the formula text (`read_formulas`), which then keeps
//! its meaning through every edit. Saving takes them out again where the old rules put them back
//! (`Saving`). Either way only the `@` change; the rest of the text stays as written. Both work
//! out each shape of formula once (`shape`): a shared formula's cells, or a formula filled down,
//! change alike.
//!
//! Which arguments take one value comes from Excel's argument classes in the function registry
//! (crate::functions::classes). Array arguments keep arrays, as they always did: SUMPRODUCT's,
//! INDEX's first, LOOKUP's vectors (`LOOKUP(2,1/(A1:A9=x),B1:B9)` needs no Ctrl+Shift+Enter).
//! Reference arguments (SUM's, COUNTIF's range) keep the range itself, while the operators inside
//! them still read one cell: `SUM(A1:A3*B1:B3)` multiplies one cell of each.

use std::collections::{HashMap, HashSet};
use std::ops::Range;
use std::mem::discriminant;

use crate::ast::{BinaryOperator, FormulaExpr, StructuredRefSpecifier, UnaryOperator};
use crate::evaluator::function_name;
use crate::functions::{registry, ArgClass, ArgPositions, FunctionDef, Returns};
use crate::parser::{parse_formula, tokens, Token};
use crate::print::{file_formula_text, formula_text};
use duke_sheets_core::{CellAddress, CellRange, FormulaKind, Workbook};

/// The workbook's defined names by what they hold: more than one value (a range of cells, an
/// array, a calculation), or one (a cell, a constant). `=TaxRate*A1` needs no `@`; LET and LAMBDA
/// names, and names in other workbooks, are neither, so nothing is assumed about them.
#[derive(Default)]
pub(crate) struct Names {
    many: HashSet<String>,
    one: HashSet<String>,
}

impl Names {
    fn of(workbook: &Workbook) -> Self {
        let mut names = Names::default();
        for name in workbook.named_ranges().iter() {
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
            let key = name.name.to_ascii_uppercase();
            if one {
                names.one.insert(key)
            } else {
                names.many.insert(key)
            };
        }
        names
    }
}

/// What may be said of a name nobody defined: in the rules (`Inserting`), a LET name or another
/// workbook's name holds one value; before dropping an `@` (`Dropping`), it may hold more.
#[derive(Clone, Copy, PartialEq)]
enum Unknown {
    Inserting,
    Dropping,
}

impl Names {
    fn many(&self, name: &str, unknown: Unknown) -> bool {
        let key = name.to_ascii_uppercase();
        self.many.contains(&key) || (unknown == Unknown::Dropping && !self.one.contains(&key))
    }
}

/// A formula cell as a file holds it.
pub struct FileFormula {
    pub sheet: usize,
    pub row: u32,
    pub col: u16,
    /// Saved with the dynamic-array mark (`cm="1"`).
    pub dynamic: bool,
}

/// Spells the formulas a file holds as the editor shows them, after the workbook's names are read.
/// One saved without the dynamic-array mark gets Excel's `@` where the old rules read one value;
/// `_xlfn.SINGLE(x)` and `_xlfn.ANCHORARRAY(x)` become `@x` and `x#`. A Ctrl+Shift+Enter array
/// formula, evaluated as an array, and a data table's TABLE get no `@`.
pub fn read_formulas(workbook: &mut Workbook, cells: &[FileFormula]) {
    let names = Names::of(workbook);
    // What was worked out for each shape of formula (a shared formula's cells are one shape).
    let (mut legacy_shapes, mut dynamic_shapes) = (Shapes::new(), Shapes::new());
    for cell in cells {
        let Some(formula) = workbook
            .worksheet_mut(cell.sheet)
            .and_then(|ws| ws.formula_data_at_mut(cell.row, cell.col))
        else {
            continue;
        };
        if formula.kind == FormulaKind::DataTable {
            continue;
        }
        let legacy = !cell.dynamic && formula.kind == FormulaKind::Dynamic;
        let shapes = if legacy {
            &mut legacy_shapes
        } else {
            &mut dynamic_shapes
        };
        if let Some(text) = shown(&formula.text, legacy, &names, shapes) {
            formula.text = text;
        }
    }
}

/// `read_formulas` for every formula of a file without the dynamic-array mark: all legacy in a
/// .xls file, all dynamic in a .xlsb one, which this reader doesn't read the mark of.
pub fn read_all_formulas(workbook: &mut Workbook, dynamic: bool) {
    let cells: Vec<FileFormula> = (0..workbook.sheet_count())
        .filter_map(|sheet| workbook.worksheet(sheet).map(|ws| (sheet, ws)))
        .flat_map(|(sheet, ws)| {
            ws.formula_cells()
                .map(move |(row, col, _)| FileFormula {
                    sheet,
                    row,
                    col,
                    dynamic,
                })
                .collect::<Vec<_>>()
        })
        .collect();
    read_formulas(workbook, &cells);
}

/// Whether nothing in `text` can take or have an `@`: no range, name, call, array, `@` or `#`.
fn nothing_to_intersect(text: &str, names: &Names) -> bool {
    names.many.is_empty() && names.one.is_empty() && !text.contains([':', '[', '#', '@', '{', '('])
}

/// The text a file's formula shows in the editor, or `None` when that is the text as it is.
/// `shapes` holds what the formulas read before showed.
fn shown(text: &str, legacy: bool, names: &Names, shapes: &mut Shapes) -> Option<String> {
    if nothing_to_intersect(text, names) {
        return None;
    }
    // A dynamic-array formula changes only where it spells `@` or `#` as a function.
    if !legacy && !spells_marks(text) {
        return None;
    }
    let (shown, _) = per_shape(shapes, text, || {
        let Ok(parsed) = parse_formula(text) else {
            return (Change::Same, false);
        };
        let marked = has_marks(&parsed);
        if !legacy && !marked {
            return (Change::Same, false);
        }
        let (target, changed) = if legacy {
            walk(parsed.clone(), Place::One, names)
        } else {
            (parsed.clone(), false)
        };
        if !changed && !marked {
            return (Change::Same, false);
        }
        // `_xlfn.SINGLE(x)` in the text doesn't line up with `@x`: written anew.
        (spelled(text, &parsed, &target), false)
    });
    shown
}

/// Saves formulas as a file holds them (`formula`), for one workbook's names.
pub struct Saving {
    names: Names,
    /// What the formulas saved before were saved as, array formulas apart.
    shapes: Shapes,
    arrays: Shapes,
}

impl Saving {
    pub fn of(workbook: &Workbook) -> Self {
        Self::with_names(Names::of(workbook))
    }

    fn with_names(names: Names) -> Self {
        Saving {
            names,
            shapes: Shapes::new(),
            arrays: Shapes::new(),
        }
    }

    /// A formula as a file saves it, given with its `=`: the text, and whether it is marked as a
    /// dynamic-array formula (`cm="1"`). One whose `@` are all where the old rules put them is
    /// saved without them, as a formula Excel 2019 and earlier read the same; any other `@` is
    /// saved as `_xlfn.SINGLE` (and `#` as `_xlfn.ANCHORARRAY`), and the mark is set when reading
    /// the formula the old way would put in more `@`: a range where one value is expected, or a
    /// result that can be an array (FILTER, `A1:A3*2`, INDEX). An `@` on what can only be one
    /// value (`=@A1`) is left out, as Excel does.
    pub fn formula(&mut self, text: &str) -> (String, bool) {
        let names = &self.names;
        if nothing_to_intersect(text, names) {
            return (text.to_string(), false);
        }
        let (saved, dynamic) = per_shape(&mut self.shapes, text, || {
            let Ok(parsed) = parse_formula(text) else {
                return (Change::Same, false);
            };
            if !has_marks(&parsed) {
                let (_, more) = walk(parsed, Place::One, names);
                return (Change::Same, more);
            }
            let expr = without_redundant_at(parsed.clone(), names);
            let plain = without_at(expr.clone());
            if !has_marks(&plain) && walk(plain.clone(), Place::One, names).0 == expr {
                return (spelled(text, &parsed, &plain), false);
            }
            let (_, more) = walk(expr.clone(), Place::One, names);
            let change = if has_marks(&expr) {
                Change::Text(file_formula_text(&expr))
            } else {
                spelled(text, &parsed, &expr)
            };
            (change, more)
        });
        (saved.unwrap_or_else(|| text.to_string()), dynamic)
    }

    /// A Ctrl+Shift+Enter array formula's text as a file saves it: `@` and `#` as `_xlfn.SINGLE`
    /// and `_xlfn.ANCHORARRAY`, and nothing else changed.
    pub fn array_formula(&mut self, text: &str) -> String {
        if !text.contains(['@', '#']) && !spells_marks(text) {
            return text.to_string();
        }
        let (saved, _) = per_shape(&mut self.arrays, text, || match parse_formula(text) {
            Ok(expr) if has_marks(&expr) => (Change::Text(file_formula_text(&expr)), false),
            _ => (Change::Same, false),
        });
        saved.unwrap_or_else(|| text.to_string())
    }
}

/// Whether `text` spells `@` or `#` as a function, as a file does: `_xlfn.SINGLE(x)`,
/// `_xlfn.ANCHORARRAY(x)`.
fn spells_marks(text: &str) -> bool {
    let upper = text.to_ascii_uppercase();
    upper.contains("SINGLE") || upper.contains("ANCHORARRAY")
}

/// A new text for a formula: `@` put in or taken out, the rest as written, or written anew.
enum Change {
    Same,
    At(String),
    Text(String),
}

/// What reading or saving made of each shape of formula (`shape`): the shape of its new text
/// (`None`: the text as it was), and when saving whether it is marked dynamic.
type Shapes = HashMap<String, (Option<String>, bool)>;

/// What `work` makes of a formula, its new text (`None`: as it is) and its mark, worked out once
/// for all the formulas of its shape. Only `@` change, so another formula of the shape gets the
/// same new shape with its own cells; a formula written anew is its own.
fn per_shape(
    shapes: &mut Shapes,
    text: &str,
    work: impl FnOnce() -> (Change, bool),
) -> (Option<String>, bool) {
    let known = shape(text);
    if let Some((key, cells)) = &known {
        if let Some((new, dynamic)) = shapes.get(key) {
            return (new.as_ref().map(|new| filled(new, text, cells)), *dynamic);
        }
    }
    let (change, dynamic) = work();
    let (new, alike) = match change {
        Change::Same => (None, true),
        Change::At(new) => (Some(new).filter(|new| new != text), true),
        Change::Text(new) => (Some(new).filter(|new| new != text), false),
    };
    let Some((key, cells)) = known.filter(|_| alike) else {
        return (new, dynamic);
    };
    let new_shape = match new.as_deref().map(|new| (new, shape(new))) {
        None => None,
        // Putting in or taking out `@` keeps the text's cells as they are.
        Some((new, Some((new_shape, new_cells))))
            if new_cells.iter().map(|at| &new[at.clone()]).eq(cells.iter().map(|at| &text[at.clone()])) =>
        {
            Some(new_shape)
        }
        Some(_) => {
            debug_assert!(false, "{text} respelled with other cells: {new:?}");
            return (new, dynamic);
        }
    };
    shapes.insert(key, (new_shape, dynamic));
    (new, dynamic)
}

/// A formula's shape: its text with each cell reference left out, and for each after the first
/// whether a range from the one before it to it is one cell (`A1:A1`); with where the references
/// are. Formulas of one shape (a shared formula's cells, a formula filled down) differ only in
/// which cells they name, which doesn't move their `@`. A reference here is a run of a name's
/// characters that reads as a cell (`$B$2`), as the parser reads one: not a call, a table, a sheet
/// (`A1!`, `A1:B2!`) or an error, and not in a string, a quoted sheet name or brackets, which stay
/// as written. `None` for a text holding the characters the shape marks cells with.
fn shape(text: &str) -> Option<(String, Vec<Range<usize>>)> {
    if text.contains(['\0', '\u{1}', '\u{2}']) {
        return None;
    }
    let name_char = |b: u8| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'$' | b'.');
    let bytes = text.as_bytes();
    let mut shape = String::with_capacity(text.len());
    let mut cells = Vec::new();
    let mut previous: Option<CellAddress> = None;
    let (mut copied, mut i) = (0, 0);
    while i < bytes.len() {
        match bytes[i] {
            // A doubled quote inside is one quote: the text goes on to the next lone one.
            quote @ (b'"' | b'\'') => {
                i += 1;
                while i < bytes.len() {
                    i += 1;
                    if bytes[i - 1] == quote {
                        if bytes.get(i) != Some(&quote) {
                            break;
                        }
                        i += 1;
                    }
                }
                continue;
            }
            b'[' => {
                let mut depth = 0;
                while i < bytes.len() {
                    depth += match bytes[i] {
                        b'[' => 1,
                        b']' => -1,
                        _ => 0,
                    };
                    i += 1;
                    if depth == 0 {
                        break;
                    }
                }
                continue;
            }
            b if !name_char(b) => {
                i += 1;
                continue;
            }
            _ => {}
        }
        let start = i;
        while i < bytes.len() && name_char(bytes[i]) {
            i += 1;
        }
        let sheet_span = bytes.get(i) == Some(&b':') && {
            let span = bytes[i + 1..]
                .iter()
                .take_while(|&&b| b.is_ascii_alphanumeric() || matches!(b, b'_' | b'.'))
                .count();
            span > 0 && bytes.get(i + 1 + span) == Some(&b'!')
        };
        if sheet_span
            || matches!(bytes.get(i), Some(b'(' | b'[' | b'!'))
            || (start > 0 && bytes[start - 1] == b'#')
        {
            continue;
        }
        let Some(cell) = cell_reference(&text[start..i]) else {
            continue;
        };
        shape.push_str(&text[copied..start]);
        shape.push('\0');
        if let Some(previous) = previous {
            let range = CellRange::new(previous, cell);
            shape.push(if range.start == range.end { '\u{1}' } else { '\u{2}' });
        }
        previous = Some(cell);
        cells.push(start..i);
        copied = i;
    }
    shape.push_str(&text[copied..]);
    Some((shape, cells))
}

/// `run` as a cell reference, `[$]letters[$]digits` as the parser reads one (checked before
/// parsing it, so that a name is turned down without building an error).
fn cell_reference(run: &str) -> Option<CellAddress> {
    let rest = run.strip_prefix('$').unwrap_or(run);
    let letters = rest.bytes().take_while(u8::is_ascii_alphabetic).count();
    let rest = &rest[letters..];
    let rest = rest.strip_prefix('$').unwrap_or(rest);
    let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
    if letters == 0 || digits == 0 || digits != rest.len() {
        return None;
    }
    CellAddress::parse(run).ok()
}

/// The text of `shape` with the cells of `text` (at `cells`, in order) put back in.
fn filled(shape: &str, text: &str, cells: &[Range<usize>]) -> String {
    let mut out = String::with_capacity(text.len() + 8);
    let mut cells = cells.iter();
    let mut copied = 0;
    for (at, c) in shape.char_indices() {
        match c {
            '\0' => {
                out.push_str(&shape[copied..at]);
                out.push_str(&text[cells.next().expect("one cell per mark").clone()]);
                copied = at + 1;
            }
            '\u{1}' | '\u{2}' => copied = at + 1,
            _ => {}
        }
    }
    out.push_str(&shape[copied..]);
    out
}

/// `text`, parsed as `from`, spelled as `to`, which differs from it only by `@`; written anew when
/// the two don't line up that way.
fn spelled(text: &str, from: &FormulaExpr, to: &FormulaExpr) -> Change {
    respell(text, from, to).map_or_else(|| Change::Text(formula_text(to)), Change::At)
}

/// `text` (parsed as `from`) with the `@` that `to` adds put in before their tokens and those it
/// drops taken out, the rest left as written: spaces, line breaks, names in their own case,
/// numbers as typed. `None` when `to` differs from `from` by more than `@`.
fn respell(text: &str, from: &FormulaExpr, to: &FormulaExpr) -> Option<String> {
    let original = tokens(text)?;
    let before = tokens(&formula_text(from))?;
    let after = tokens(&formula_text(to))?;
    let same = |a: &Token, b: &Token| discriminant(a) == discriminant(b);
    if original.len() != before.len()
        || original.iter().zip(&before).any(|(o, b)| !same(&o.1, &b.1))
    {
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
        found |= matches!(
            e,
            FormulaExpr::UnaryOp {
                op: UnaryOperator::ImplicitIntersection | UnaryOperator::SpillRange,
                ..
            }
        );
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
        FormulaExpr::UnaryOp { op, operand } => FormulaExpr::UnaryOp {
            op,
            operand: Box::new(f(*operand)),
        },
        FormulaExpr::BinaryOp { op, left, right } => FormulaExpr::BinaryOp {
            op,
            left: Box::new(f(*left)),
            right: Box::new(f(*right)),
        },
        FormulaExpr::Function { name, args } => FormulaExpr::Function {
            name,
            args: args.into_iter().map(&mut *f).collect(),
        },
        FormulaExpr::ExternalFunction { book, name, args } => FormulaExpr::ExternalFunction {
            book,
            name,
            args: args.into_iter().map(&mut *f).collect(),
        },
        FormulaExpr::Array(rows) => FormulaExpr::Array(
            rows.into_iter()
                .map(|row| row.into_iter().map(&mut *f).collect())
                .collect(),
        ),
        other => other,
    }
}

fn without_at(expr: FormulaExpr) -> FormulaExpr {
    match expr {
        FormulaExpr::UnaryOp {
            op: UnaryOperator::ImplicitIntersection,
            operand,
        } => without_at(*operand),
        other => map_parts(other, &mut without_at),
    }
}

/// Leaves out each `@` on what can only be one value: `@A1`, `@SUM(A1:A3)`, a name for one cell.
/// One on a LET name or another workbook's name stays.
fn without_redundant_at(expr: FormulaExpr, names: &Names) -> FormulaExpr {
    match expr {
        FormulaExpr::UnaryOp {
            op: UnaryOperator::ImplicitIntersection,
            operand,
        } if !could_be_array(&operand, names, Unknown::Dropping) => {
            without_redundant_at(*operand, names)
        }
        other => map_parts(other, &mut |part| without_redundant_at(part, names)),
    }
}

/// The formula as Excel 365 shows it when it was saved without the dynamic-array mark: `@`
/// where the old rules read one value. The `@` already there stay.
#[cfg(test)]
fn intersections(expr: FormulaExpr, names: &Names) -> FormulaExpr {
    walk(expr, Place::One, names).0
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

/// The place of the argument at `position` of `count`, the call itself being at `call`: one its
/// result takes (IF's value_if_true is where IF is), else from the argument's class.
fn place_of(def: &FunctionDef, position: usize, count: usize, call: Place) -> Place {
    if def.result_args.includes_in(position, count) {
        return call;
    }
    match def.arg_class(position) {
        ArgClass::V => Place::One,
        ArgClass::R => Place::Whole,
        ArgClass::A | ArgClass::RA => Place::Array,
    }
}

fn at(expr: FormulaExpr) -> FormulaExpr {
    FormulaExpr::UnaryOp {
        op: UnaryOperator::ImplicitIntersection,
        operand: Box::new(expr),
    }
}

/// `expr` with the `@` the old rules put at `place`, and whether it gained any.
fn walk(expr: FormulaExpr, place: Place, names: &Names) -> (FormulaExpr, bool) {
    let mut changed = false;
    let expr = walk_into(expr, place, names, &mut changed);
    (expr, changed)
}

fn walk_into(expr: FormulaExpr, place: Place, names: &Names, changed: &mut bool) -> FormulaExpr {
    if place == Place::Array {
        return array_place(expr, names, changed);
    }
    match expr {
        FormulaExpr::BinaryOp {
            op: op @ (BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect),
            left,
            right,
        } => {
            let reference = FormulaExpr::BinaryOp {
                op,
                left: Box::new(walk_into(*left, Place::Whole, names, changed)),
                right: Box::new(walk_into(*right, Place::Whole, names, changed)),
            };
            one_value(reference, place, names, changed)
        }
        FormulaExpr::BinaryOp { op, left, right } => FormulaExpr::BinaryOp {
            op,
            left: Box::new(walk_into(*left, Place::One, names, changed)),
            right: Box::new(walk_into(*right, Place::One, names, changed)),
        },
        FormulaExpr::UnaryOp {
            op: op @ (UnaryOperator::Negate | UnaryOperator::Plus | UnaryOperator::Percent),
            operand,
        } => FormulaExpr::UnaryOp {
            op,
            operand: Box::new(walk_into(*operand, Place::One, names, changed)),
        },
        // `(A1,B1)` and `(A1:A3)` take the `@` outside their parentheses.
        FormulaExpr::UnaryOp {
            op: UnaryOperator::Paren,
            operand,
        } => {
            let inner = FormulaExpr::UnaryOp {
                op: UnaryOperator::Paren,
                operand: Box::new(walk_into(*operand, Place::Whole, names, changed)),
            };
            one_value(inner, place, names, changed)
        }
        FormulaExpr::Function { name, args } => {
            let def = registry().get(function_name(&name));
            let count = args.len();
            let args = args
                .into_iter()
                .enumerate()
                .map(|(position, arg)| {
                    let takes =
                        def.map_or(Place::Whole, |def| place_of(def, position, count, place));
                    walk_into(arg, takes, names, changed)
                })
                .collect();
            one_value(FormulaExpr::Function { name, args }, place, names, changed)
        }
        reference @ (FormulaExpr::RangeRef(_)
        | FormulaExpr::NameRef(_)
        | FormulaExpr::StructuredRef(_)
        | FormulaExpr::UnaryOp {
            op: UnaryOperator::SpillRange,
            ..
        }) => one_value(reference, place, names, changed),
        other => other,
    }
}

/// `@expr` where one value is expected and `expr` can give more.
fn one_value(expr: FormulaExpr, place: Place, names: &Names, changed: &mut bool) -> FormulaExpr {
    if place == Place::One && could_be_array(&expr, names, Unknown::Inserting) {
        *changed = true;
        at(expr)
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
                .map_or(ArgPositions::None, |def| def.result_args);
            let count = args.len();
            let args = args
                .into_iter()
                .enumerate()
                .map(|(position, arg)| {
                    if decides != ArgPositions::None && !decides.includes_in(position, count) {
                        walk_into(arg, Place::One, names, changed)
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
fn could_be_array(expr: &FormulaExpr, names: &Names, unknown: Unknown) -> bool {
    let could = |e: &FormulaExpr| could_be_array(e, names, unknown);
    match expr {
        FormulaExpr::RangeRef(r) => r.range.start != r.range.end,
        FormulaExpr::NameRef(name) => names.many(name, unknown),
        FormulaExpr::StructuredRef(sr) => !sr.specifiers.contains(&StructuredRefSpecifier::ThisRow),
        FormulaExpr::Array(rows) => rows.len() > 1 || rows.first().is_some_and(|row| row.len() > 1),
        FormulaExpr::BinaryOp {
            op: BinaryOperator::Range | BinaryOperator::Union | BinaryOperator::Intersect,
            ..
        } => true,
        FormulaExpr::BinaryOp { left, right, .. } => could(left) || could(right),
        FormulaExpr::UnaryOp {
            op: UnaryOperator::ImplicitIntersection,
            ..
        } => false,
        FormulaExpr::UnaryOp {
            op: UnaryOperator::SpillRange,
            ..
        } => true,
        FormulaExpr::UnaryOp { operand, .. } => could(operand),
        FormulaExpr::ExternalFunction { .. } => true,
        FormulaExpr::Function { name, args } => match registry().get(function_name(name)) {
            // A LAMBDA called in place, or a function Excel doesn't have (an add-in's).
            None => true,
            Some(def) => match def.returns {
                Returns::Array | Returns::Reference => true,
                Returns::PerCell => args.iter().any(could),
                Returns::Last => args.last().is_some_and(could),
                Returns::Value => args.iter().enumerate().any(|(position, arg)| {
                    place_of(def, position, args.len(), Place::One) == Place::One && could(arg)
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
        Names {
            many: ["PRICES".to_string()].into_iter().collect(),
            one: ["TAXRATE".to_string()].into_iter().collect(),
        }
    }

    fn shown_legacy(text: &str) -> String {
        shown(text, true, &names(), &mut Shapes::new()).unwrap_or_else(|| text.to_string())
    }

    fn saving() -> Saving {
        Saving::with_names(names())
    }

    #[test]
    fn ranges_where_one_value_is_expected_get_an_at() {
        assert_eq!(shown_legacy("=A1:A3"), "=@A1:A3");
        assert_eq!(shown_legacy("=SUM(A1:A3*B1:B3)"), "=SUM(@A1:A3*@B1:B3)");
        assert_eq!(
            shown_legacy("=SUMIFS(B:B,A:A,C1:C3)"),
            "=SUMIFS(B:B,A:A,@C1:C3)"
        );
        assert_eq!(
            shown_legacy("=NETWORKDAYS(A:A,B:B)"),
            "=NETWORKDAYS(@A:A,@B:B)"
        );
        assert_eq!(
            shown_legacy("=MAX(IF(A1:A9=1,B1:B9))"),
            "=MAX(IF(@A1:A9=1,B1:B9))"
        );
        assert_eq!(
            shown_legacy("=IF(A1>1,B1:B3,C1)+0"),
            "=IF(A1>1,@B1:B3,C1)+0"
        );
        assert_eq!(shown_legacy("=INDEX(A:A,5)"), "=@INDEX(A:A,5)");
        assert_eq!(shown_legacy("=Prices*2"), "=@Prices*2");
        assert_eq!(shown_legacy("=(A1,B1)"), "=@(A1,B1)");
        assert_eq!(
            shown_legacy("=SUMPRODUCT(IF(A1:A3>1,1,0))"),
            "=SUMPRODUCT(IF(@A1:A3>1,1,0))"
        );
        assert_eq!(
            shown_legacy("=SUM(MID(A1,ROW(A1:A9),1)*1)"),
            "=SUM(MID(A1,@ROW(A1:A9),1)*1)"
        );
        assert_eq!(shown_legacy("=SUM(A2#*B2#)"), "=SUM(@A2#*@B2#)");
        assert_eq!(shown_legacy("=A1:A3%"), "=@A1:A3%");
        assert_eq!(
            shown_legacy("=SUM(IFERROR(1/0,A1:A3))"),
            "=SUM(IFERROR(1/0,@A1:A3))"
        );
        assert_eq!(
            shown_legacy("=SUM(SWITCH(A1,1,B1:B3,C1:C3))"),
            "=SUM(SWITCH(A1,1,B1:B3,C1:C3))"
        );
    }

    #[test]
    fn a_dynamic_files_formulas_change_only_where_they_spell_at() {
        let mut workbook = Workbook::new();
        let sheet = workbook.worksheet_mut(0).unwrap();
        sheet
            .set_formula_with_cached_value_at(
                0,
                0,
                "=_xlfn.SINGLE(B1:B3)+_xlfn.ANCHORARRAY(C1)",
                0.0.into(),
            )
            .unwrap();
        sheet
            .set_formula_with_cached_value_at(1, 0, "=SUM(B1:B3*2)", 0.0.into())
            .unwrap();
        read_all_formulas(&mut workbook, true);
        let sheet = workbook.worksheet(0).unwrap();
        assert_eq!(sheet.get_formula_at(0, 0), Some("=@B1:B3+C1#"));
        assert_eq!(sheet.get_formula_at(1, 0), Some("=SUM(B1:B3*2)"));
    }

    #[test]
    fn only_the_at_change_in_the_text() {
        assert_eq!(
            shown_legacy("= SUM( A1:A3 * b1:b3 )\n+MyUdf(1E+20,[1]$A$1)"),
            "= SUM( @A1:A3 * @b1:b3 )\n+@MyUdf(1E+20,[1]$A$1)"
        );
        let (saved, dynamic) = saving().formula("= SUM( @A1:A3 * @b1:b3 )");
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
            "=N(A1:A3)+T(A1:A3)",
            "=XNPV(0.1,B2:B6,A2:A6)",
            "=GCD(A1:A3)+FORECAST.ETS(A1,B1:B9,C1:C9)",
        ] {
            assert_eq!(shown_legacy(text), text);
        }
    }

    #[test]
    fn what_the_rules_build_prints_and_parses_back() {
        for text in [
            "=A1:A3%",
            "=SUM(A2#*B2#)",
            "=-A1:A3",
            "=(A1:A3)*2+(B1:B3*2)",
            "=A1:B3 B2:C4",
            "=INDEX(A:A,5)%",
        ] {
            let legacy = intersections(parse_formula(text).unwrap(), &names());
            assert_eq!(
                parse_formula(&formula_text(&legacy)).unwrap(),
                legacy,
                "{text}"
            );
        }
    }

    #[test]
    fn saving_marks_formulas_that_read_differently() {
        let saved = |text: &str| saving().formula(text);
        assert_eq!(
            saved("=SUM(A1:A3*B1:B3)"),
            ("=SUM(A1:A3*B1:B3)".to_string(), true)
        );
        assert_eq!(
            saved("=SUM(@A1:A3*@B1:B3)"),
            ("=SUM(A1:A3*B1:B3)".to_string(), false)
        );
        assert_eq!(
            saved("=SUM(@A1:A3)"),
            ("=SUM(_xlfn.SINGLE(A1:A3))".to_string(), false)
        );
        assert_eq!(
            saved("=@A1+@SUM(A1:A3)"),
            ("=A1+SUM(A1:A3)".to_string(), false)
        );
        assert_eq!(
            saved("=FILTER(A1:A3,A1:A3>1)"),
            ("=FILTER(A1:A3,A1:A3>1)".to_string(), true)
        );
        assert!(saved("=SUM(OFFSET(A1,0,0,3,1)*B1)").1);
        assert!(saved("=SUM(--(A2#>5))").1);
        assert!(saved("=MAX(A2#-B1)").1);
        assert!(!saved("=[@Qty]*[@Price]").1);
        assert!(!saved("=TaxRate*A1").1);
        assert!(!saved("=VLOOKUP(A1,C:D,2,FALSE)").1);
        // An `@` on a LET name stays: the name may hold a range.
        assert_eq!(
            saved("=LET(r,A1:A3,@r*2)").0,
            "=LET(r,A1:A3,_xlfn.SINGLE(r)*2)"
        );
        assert_eq!(saved("=@TaxRate*2"), ("=TaxRate*2".to_string(), false));
    }

    #[test]
    fn formulas_of_one_shape_change_alike() {
        let names = names();
        let mut shapes = Shapes::new();
        let mut read = |text: &str| shown(text, true, &names, &mut shapes);
        assert_eq!(read("=SUM(A1:A3*B1)").as_deref(), Some("=SUM(@A1:A3*B1)"));
        assert_eq!(read("=SUM(A2:A4*B2)").as_deref(), Some("=SUM(@A2:A4*B2)"));
        assert_eq!(read("=SUM($A$1:$A$9*B10)").as_deref(), Some("=SUM(@$A$1:$A$9*B10)"));
        // A range of one cell takes no `@`, whichever cell it is.
        assert_eq!(read("=A1:A1+1"), None);
        assert_eq!(read("=A1:A2+1").as_deref(), Some("=@A1:A2+1"));
        assert_eq!(read("=Sheet1!B2:B2"), None);
        assert_eq!(read("=Sheet1!B2:B5").as_deref(), Some("=@Sheet1!B2:B5"));
        assert_eq!(read("=A3:A3+1"), None);
        // Text that reads as a cell, in a string or a sheet's name, is left as it is.
        assert_eq!(read("=\"A1\"&B1:B3").as_deref(), Some("=\"A1\"&@B1:B3"));
        assert_eq!(read("=\"Q9\"&C1:C3").as_deref(), Some("=\"Q9\"&@C1:C3"));
        // A sheet's name between a range's two cells doesn't stand for a cell.
        let fresh = |text: &str| shown(text, true, &self::names(), &mut Shapes::new());
        assert!(fresh("=C1:'B2'!A1+1").is_some());
        for text in ["=A1:'B2'!A1+1", "=C1:'B2'!A1+1", "=A1:Q1:Q4!A1+1", "=C1:Q1:Q4!A1+1"] {
            assert_eq!(read(text), fresh(text), "{text}");
        }
        assert_eq!(
            read("='Q1 A1'!B1:B3*2").as_deref(),
            Some("=@'Q1 A1'!B1:B3*2")
        );
        assert_eq!(
            read("='Q1 B7'!C1:C3*2").as_deref(),
            Some("=@'Q1 B7'!C1:C3*2")
        );

        let mut saving = saving();
        for _ in 0..2 {
            assert_eq!(
                saving.formula("=SUM(@A1:A3*B1)"),
                ("=SUM(A1:A3*B1)".to_string(), false)
            );
            assert_eq!(
                saving.formula("=SUM(@A2:A4*@B2:B4)"),
                ("=SUM(A2:A4*B2:B4)".to_string(), false)
            );
            assert_eq!(
                saving.formula("=SUM(A2:A4*B2)"),
                ("=SUM(A2:A4*B2)".to_string(), true)
            );
            assert_eq!(
                saving.formula("=SUM(@A2:A4)"),
                ("=SUM(_xlfn.SINGLE(A2:A4))".to_string(), false)
            );
        }
    }

    #[test]
    fn shapes_find_the_cells_the_parser_reads() {
        for text in [
            "=SUM(A1:B2,$C$3)*LOG10(D4)+LOG10 (E5)",
            "=A1:Sheet2!B2+Sheet1:Sheet3!C1+Q1:Q4!A1",
            "=IF(x1>0,A1#,#N/A)+#REF!+R1C1+TRUE",
            "=[1]Sheet1!A1+Table1[[#This Row],[B2]]+'Q1 A1'!C3&\"D4\"",
            "=1E5+.5+A1048576+XFE1",
        ] {
            let (_, cells) = shape(text).unwrap();
            let read: Vec<_> = tokens(text)
                .unwrap()
                .into_iter()
                .filter(|(_, token)| matches!(token, Token::CellRef(_)))
                .map(|(at, _)| at)
                .collect();
            let valid: Vec<_> = read
                .into_iter()
                .filter(|at| CellAddress::parse(&text[at.clone()]).is_ok())
                .collect();
            assert_eq!(cells, valid, "{text}");
        }
    }
}
