//! Formula text from a parsed formula. The editor's text writes `@x` and `A1#` as typed; a file
//! writes them as `_xlfn.SINGLE(x)` and `_xlfn.ANCHORARRAY(A1)`, as Excel does.

use crate::ast::{BinaryOperator, FormulaExpr, StructuredRefSpecifier, StructuredReference, UnaryOperator};
use duke_sheets_core::{CellAddress, CellRange, MAX_COLS, MAX_ROWS};

/// The formula as the editor shows it, with its `=`.
pub fn formula_text(expr: &FormulaExpr) -> String {
    let mut out = String::from("=");
    write(expr, false, &mut out);
    out
}

/// The formula as a file holds it, with its `=`.
pub fn file_formula_text(expr: &FormulaExpr) -> String {
    let mut out = String::from("=");
    write(expr, true, &mut out);
    out
}

fn write(expr: &FormulaExpr, file: bool, out: &mut String) {
    match expr {
        FormulaExpr::Number(n) => out.push_str(&n.to_string()),
        FormulaExpr::String(s) => {
            out.push('"');
            out.push_str(&s.replace('"', "\"\""));
            out.push('"');
        }
        FormulaExpr::Boolean(b) => out.push_str(if *b { "TRUE" } else { "FALSE" }),
        FormulaExpr::Error(e) => out.push_str(e.as_str()),
        FormulaExpr::CellRef(c) => {
            sheet(c.sheet.as_deref(), out);
            out.push_str(&c.address.to_a1_string());
        }
        FormulaExpr::RangeRef(r) => {
            sheet(r.sheet.as_deref(), out);
            range(&r.range, out);
        }
        FormulaExpr::NameRef(name) => out.push_str(name),
        FormulaExpr::StructuredRef(sr) => structured(sr, out),
        FormulaExpr::ExternalRef(r) => {
            out.push('[');
            out.push_str(&r.book);
            out.push(']');
            if let Some(name) = &r.sheet {
                out.push_str(&quoted(name));
                out.push('!');
            }
            out.push_str(&r.address.to_a1_string());
        }
        FormulaExpr::BinaryOp { op, left, right } => {
            write(left, file, out);
            out.push_str(match op {
                BinaryOperator::Add => "+",
                BinaryOperator::Subtract => "-",
                BinaryOperator::Multiply => "*",
                BinaryOperator::Divide => "/",
                BinaryOperator::Power => "^",
                BinaryOperator::Equal => "=",
                BinaryOperator::NotEqual => "<>",
                BinaryOperator::LessThan => "<",
                BinaryOperator::LessEqual => "<=",
                BinaryOperator::GreaterThan => ">",
                BinaryOperator::GreaterEqual => ">=",
                BinaryOperator::Concat => "&",
                BinaryOperator::Range => ":",
                BinaryOperator::Union => ",",
                BinaryOperator::Intersect => " ",
            });
            write(right, file, out);
        }
        FormulaExpr::UnaryOp { op, operand } => match op {
            UnaryOperator::Negate | UnaryOperator::Plus => {
                out.push(if *op == UnaryOperator::Negate { '-' } else { '+' });
                write(operand, file, out);
            }
            UnaryOperator::Percent => {
                write(operand, file, out);
                out.push('%');
            }
            UnaryOperator::Paren => {
                out.push('(');
                write(operand, file, out);
                out.push(')');
            }
            UnaryOperator::ImplicitIntersection if file => call("_xlfn.SINGLE", std::slice::from_ref(&**operand), file, out),
            UnaryOperator::ImplicitIntersection => {
                out.push('@');
                operand_of_mark(operand, file, out);
            }
            UnaryOperator::SpillRange if file => call("_xlfn.ANCHORARRAY", std::slice::from_ref(&**operand), file, out),
            UnaryOperator::SpillRange => {
                operand_of_mark(operand, file, out);
                out.push('#');
            }
        },
        FormulaExpr::Function { name, args } if name == crate::parser::LAMBDA_CALL => {
            if let Some((lambda, values)) = args.split_first() {
                write(lambda, file, out);
                call("", values, file, out);
            }
        }
        FormulaExpr::Function { name, args } => call(&file_prefixes(name), args, file, out),
        FormulaExpr::ExternalFunction { book, name, args } => {
            out.push('[');
            out.push_str(book);
            out.push_str("]!");
            call(name, args, file, out);
        }
        FormulaExpr::Array(rows) => {
            out.push('{');
            for (i, row) in rows.iter().enumerate() {
                if i > 0 {
                    out.push(';');
                }
                for (j, value) in row.iter().enumerate() {
                    if j > 0 {
                        out.push(',');
                    }
                    write(value, file, out);
                }
            }
            out.push('}');
        }
        FormulaExpr::Empty => {}
    }
}

/// The operand of `@` or `#`: in parentheses unless it is a reference or a single term, so
/// `_xlfn.SINGLE(A1:A3*2)` is `@(A1:A3*2)`.
fn operand_of_mark(operand: &FormulaExpr, file: bool, out: &mut String) {
    let bare = matches!(
        operand,
        FormulaExpr::CellRef(_)
            | FormulaExpr::RangeRef(_)
            | FormulaExpr::NameRef(_)
            | FormulaExpr::StructuredRef(_)
            | FormulaExpr::ExternalRef(_)
            | FormulaExpr::Function { .. }
            | FormulaExpr::ExternalFunction { .. }
            | FormulaExpr::Array(_)
            | FormulaExpr::Number(_)
            | FormulaExpr::String(_)
            | FormulaExpr::Boolean(_)
            | FormulaExpr::Error(_)
            | FormulaExpr::UnaryOp { op: UnaryOperator::Paren | UnaryOperator::SpillRange, .. }
            | FormulaExpr::BinaryOp { op: BinaryOperator::Range | BinaryOperator::Intersect, .. }
    );
    if bare {
        write(operand, file, out);
    } else {
        out.push('(');
        write(operand, file, out);
        out.push(')');
    }
}

fn call(name: &str, args: &[FormulaExpr], file: bool, out: &mut String) {
    out.push_str(name);
    out.push('(');
    for (i, arg) in args.iter().enumerate() {
        if i > 0 {
            out.push(',');
        }
        write(arg, file, out);
    }
    out.push(')');
}

/// The parser keeps names in capitals; the prefixes Excel writes in a file are lower case
/// (`_xlfn._xlws.FILTER`).
fn file_prefixes(name: &str) -> String {
    let mut rest = name;
    let mut out = String::new();
    for prefix in ["_XLFN.", "_XLWS."] {
        if let Some(after) = rest.strip_prefix(prefix) {
            out.push_str(&prefix.to_ascii_lowercase());
            rest = after;
        }
    }
    out.push_str(rest);
    out
}

fn sheet(name: Option<&str>, out: &mut String) {
    if let Some(name) = name {
        out.push_str(&quoted(name));
        out.push('!');
    }
}

/// A sheet name as written before `!`: in quotes unless it is a plain ASCII word that can't be
/// read as a cell or a boolean (`Data`, `Q1_2024`; but `'Données'`, `'My Sheet'`, `'A1'`).
pub fn quoted(name: &str) -> String {
    if sheet_needs_quotes(name) {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

/// Whether a sheet name needs quotes before `!` (see `quoted`).
pub fn sheet_needs_quotes(name: &str) -> bool {
    let plain = name.chars().next().is_some_and(|c| c.is_ascii_alphabetic() || c == '_')
        && name.chars().all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.')
        && CellAddress::parse(name).is_err()
        && !name.eq_ignore_ascii_case("TRUE")
        && !name.eq_ignore_ascii_case("FALSE");
    !plain
}

/// A range as written: `A1:B2`, whole columns `A:C` and whole rows `1:3`.
fn range(range: &CellRange, out: &mut String) {
    let (a, b) = (range.start, range.end);
    let dollar = |absolute: bool| if absolute { "$" } else { "" };
    if a.row == 0 && b.row == MAX_ROWS - 1 && a.row_absolute && b.row_absolute {
        out.push_str(&format!(
            "{}{}:{}{}",
            dollar(a.col_absolute),
            CellAddress::column_to_letters(a.col),
            dollar(b.col_absolute),
            CellAddress::column_to_letters(b.col)
        ));
    } else if a.col == 0 && b.col == MAX_COLS - 1 && a.col_absolute && b.col_absolute {
        out.push_str(&format!("{}{}:{}{}", dollar(a.row_absolute), a.row + 1, dollar(b.row_absolute), b.row + 1));
    } else {
        out.push_str(&format!("{}:{}", a.to_a1_string(), b.to_a1_string()));
    }
}

/// `Table1[Col]`, `[@Col]`, `Table1[[#Headers],[Col]]`, `Table1[[A]:[C]]`.
fn structured(sr: &StructuredReference, out: &mut String) {
    if let Some(table) = &sr.table {
        out.push_str(table);
    }
    let column = |name: &str| {
        let mut escaped = String::new();
        for c in name.chars() {
            if matches!(c, '[' | ']' | '#' | '\'') {
                escaped.push('\'');
            }
            escaped.push(c);
        }
        escaped
    };
    let simple = |name: &str| name.chars().all(|c| c.is_alphanumeric() || c == '_');
    let columns = match (&sr.column, &sr.last_column) {
        (Some(first), Some(last)) => Some(format!("[{}]:[{}]", column(first), column(last))),
        (Some(first), None) => Some(format!("[{}]", column(first))),
        _ => None,
    };
    if sr.specifiers == [StructuredRefSpecifier::ThisRow] && sr.last_column.is_none() {
        match &sr.column {
            Some(name) if simple(name) => out.push_str(&format!("[@{}]", column(name))),
            _ => out.push_str(&format!("[@{}]", columns.unwrap_or_default())),
        }
        return;
    }
    let mut groups: Vec<String> = sr
        .specifiers
        .iter()
        .map(|spec| {
            match spec {
                StructuredRefSpecifier::All => "[#All]",
                StructuredRefSpecifier::Data => "[#Data]",
                StructuredRefSpecifier::Headers => "[#Headers]",
                StructuredRefSpecifier::Totals => "[#Totals]",
                StructuredRefSpecifier::ThisRow => "[#This Row]",
            }
            .to_string()
        })
        .collect();
    groups.extend(columns);
    match groups.as_slice() {
        [] => out.push_str("[]"),
        [one] if sr.last_column.is_none() => out.push_str(one),
        _ => {
            out.push('[');
            out.push_str(&groups.join(","));
            out.push(']');
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::parse_formula;

    #[test]
    fn text_parses_back_to_the_same_formula() {
        for text in [
            "=SUM(A1:B2,$C$3)*-2%",
            "=IF(A1>=1,\"say \"\"hi\"\"\",FALSE)&#N/A",
            "='My Sheet'!A1+Data!$A:$C+Sheet1!1:3+'Données'!A1",
            "=SUM((A1:B2,C1:C3))+A1:B3 B2:C4",
            "=Table1[Col]+[@Qty]*Table1[[#Headers],[My Col]]+Table1[[A]:[C]]+[@[My Col]]",
            "=Table2[[#This Row],[Description]:[Description]]+Table2[@[Description]:[Description]]",
            "=_xlfn.XLOOKUP(A1,{1,2;3,4},B:B,,1)",
            "=@A1:A3+A1#+_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)",
            "=[1]!TBLink(\"a\")+[Book.xlsx]Sheet1!A1",
        ] {
            let expr = parse_formula(text).unwrap();
            assert_eq!(parse_formula(&formula_text(&expr)).unwrap(), expr, "{text} -> {}", formula_text(&expr));
            assert_eq!(parse_formula(&file_formula_text(&expr)).unwrap(), expr, "{text}");
        }
    }

    #[test]
    fn at_on_a_calculation_keeps_its_parentheses() {
        let expr = parse_formula("=_xlfn.SINGLE(A1:A3*2)+_xlfn.ANCHORARRAY(B1)").unwrap();
        assert_eq!(formula_text(&expr), "=@(A1:A3*2)+B1#");
    }

    #[test]
    fn a_file_writes_at_and_spill_as_functions() {
        let expr = parse_formula("=SUM(@A1:A3)+B1#").unwrap();
        assert_eq!(file_formula_text(&expr), "=SUM(_xlfn.SINGLE(A1:A3))+_xlfn.ANCHORARRAY(B1)");
        assert_eq!(formula_text(&parse_formula(&file_formula_text(&expr)).unwrap()), "=SUM(@A1:A3)+B1#");
    }
}
