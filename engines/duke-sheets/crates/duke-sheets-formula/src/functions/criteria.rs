//! Criteria matching for SUMIF, COUNTIF, AVERAGEIF and related functions
//!
//! Excel criteria can be:
//! - A number: exact match (e.g., 5)
//! - A text string: case-insensitive match (e.g., "apple")
//! - A comparison expression: ">5", ">=10", "<100", "<=50", "<>0", "=5"
//! - Wildcards: "*" matches any characters, "?" matches single character
//! - Empty string: matches empty cells
//! - An error, or its text ("#N/A"): matches cells holding that error

use crate::evaluator::FormulaValue;

/// Criteria matcher for SUMIF/COUNTIF/AVERAGEIF and related functions
/// Handles comparison operators, wildcards, and exact matching
#[derive(Debug)]
pub struct CriteriaMatcher {
    criteria_type: CriteriaType,
}

#[derive(Debug)]
enum CriteriaType {
    /// Exact number match
    Number(f64),
    /// A number written as text (`"5"`): matches the number and the same text
    NumberOrText(f64, String),
    /// TRUE or FALSE: matches only that boolean
    Boolean(bool),
    /// Comparison with number (operator, value)
    Comparison(ComparisonOp, f64),
    /// Text match (case-insensitive, with wildcards)
    Text(String),
    TextComparison(ComparisonOp, String),
    /// Match empty values
    Empty,
    /// An error (`NA()`, `"#N/A"`): matches cells holding that error
    Error(duke_sheets_core::CellError),
}

#[derive(Debug, Clone, Copy)]
enum ComparisonOp {
    Equal,
    NotEqual,
    LessThan,
    LessEqual,
    GreaterThan,
    GreaterEqual,
}

impl CriteriaMatcher {
    /// Create a new criteria matcher from a FormulaValue
    pub fn new(criteria: &FormulaValue) -> Self {
        let criteria_type = match criteria {
            FormulaValue::Number(n) => CriteriaType::Number(*n),
            FormulaValue::Boolean(b) => CriteriaType::Boolean(*b),
            FormulaValue::String(s) => Self::parse_string_criteria(s),
            FormulaValue::Empty => CriteriaType::Empty,
            FormulaValue::Error(e) => CriteriaType::Error(*e),
            FormulaValue::Array { .. } => CriteriaType::Empty, // Arrays as criteria not supported
        };

        Self { criteria_type }
    }

    fn parse_string_criteria(s: &str) -> CriteriaType {
        // Text is matched as written, spaces included: "  2E" finds only "  2E". Numbers, percents
        // and dates in it are read as Excel reads them (">=1/1/2024", "5%").
        if s.is_empty() {
            return CriteriaType::Empty;
        }

        if let Some(error) = duke_sheets_core::CellError::parse(s) {
            return CriteriaType::Error(error);
        }

        // Try to parse as comparison operator + number
        if let Some(ct) = Self::try_parse_comparison(s) {
            return ct;
        }

        // Try to parse as plain number
        if let Some(n) = crate::coerce::text_to_number(s) {
            return CriteriaType::NumberOrText(n, s.trim().to_lowercase());
        }
        match s.to_ascii_uppercase().as_str() {
            "TRUE" => return CriteriaType::Boolean(true),
            "FALSE" => return CriteriaType::Boolean(false),
            _ => {}
        }

        // Text match (case-insensitive, supports wildcards)
        CriteriaType::Text(s.to_lowercase())
    }

    fn try_parse_comparison(s: &str) -> Option<CriteriaType> {
        // Check for comparison operators (order matters - check longer ones first)
        let (op, rest) = if let Some(rest) = s.strip_prefix(">=") {
            (ComparisonOp::GreaterEqual, rest)
        } else if let Some(rest) = s.strip_prefix("<=") {
            (ComparisonOp::LessEqual, rest)
        } else if let Some(rest) = s.strip_prefix("<>") {
            (ComparisonOp::NotEqual, rest)
        } else if let Some(rest) = s.strip_prefix('>') {
            (ComparisonOp::GreaterThan, rest)
        } else if let Some(rest) = s.strip_prefix('<') {
            (ComparisonOp::LessThan, rest)
        } else if let Some(rest) = s.strip_prefix('=') {
            (ComparisonOp::Equal, rest)
        } else {
            return None;
        };

        // Try to parse the number part
        if let Some(n) = crate::coerce::text_to_number(rest) {
            Some(CriteriaType::Comparison(op, n))
        } else {
            Some(CriteriaType::TextComparison(op, rest.to_lowercase()))
        }
    }

    /// Check if a value matches the criteria
    pub fn matches(&self, value: &FormulaValue) -> bool {
        match &self.criteria_type {
            CriteriaType::Number(criteria_num) => is_number(value, *criteria_num),

            CriteriaType::NumberOrText(criteria_num, text) => {
                is_number(value, *criteria_num)
                    || matches!(value, FormulaValue::String(s) if s.trim().to_lowercase() == *text)
            }

            CriteriaType::Boolean(b) => matches!(value, FormulaValue::Boolean(v) if v == b),

            CriteriaType::Comparison(op, criteria_num) => {
                // Comparisons only work on numeric values: ">0" doesn't count TRUE, but "<>5"
                // counts everything that isn't the number 5, text and blanks included
                let n = match value {
                    FormulaValue::Number(n) => *n,
                    _ => return matches!(op, ComparisonOp::NotEqual),
                };
                match op {
                    ComparisonOp::Equal => (n - criteria_num).abs() < 1e-10,
                    ComparisonOp::NotEqual => (n - criteria_num).abs() >= 1e-10,
                    ComparisonOp::LessThan => n < *criteria_num,
                    ComparisonOp::LessEqual => n <= *criteria_num,
                    ComparisonOp::GreaterThan => n > *criteria_num,
                    ComparisonOp::GreaterEqual => n >= *criteria_num,
                }
            }

            // Text criteria test text only: "*" counts no blanks or numbers, while "<>x" counts
            // everything that isn't x. "=" alone matches blanks.
            CriteriaType::Text(pattern) => match value {
                FormulaValue::String(s) => super::wildcard::matches(pattern, &s.to_lowercase()),
                _ => false,
            },

            CriteriaType::TextComparison(op, pattern) => {
                let text = match value {
                    FormulaValue::String(s) => s.to_lowercase(),
                    FormulaValue::Empty if pattern.is_empty() => String::new(),
                    _ => return matches!(op, ComparisonOp::NotEqual),
                };
                let matched = super::wildcard::matches(pattern, &text);
                match op {
                    ComparisonOp::Equal => matched,
                    ComparisonOp::NotEqual => !matched,
                    ComparisonOp::LessThan => text < *pattern,
                    ComparisonOp::LessEqual => text <= *pattern,
                    ComparisonOp::GreaterThan => text > *pattern,
                    ComparisonOp::GreaterEqual => text >= *pattern,
                }
            }

            CriteriaType::Empty => {
                matches!(value, FormulaValue::Empty)
                    || matches!(value, FormulaValue::String(s) if s.is_empty())
            }

            CriteriaType::Error(error) => matches!(value, FormulaValue::Error(e) if e == error),
        }
    }
}

/// A number, or text that reads as it: SUMIFS(K:K,A:A,DATE(2021,3,1)) adds the rows whose date
/// is the text "3-1-21". Booleans never match.
fn is_number(value: &FormulaValue, number: f64) -> bool {
    let same = |n: f64| (n - number).abs() < 1e-10;
    match value {
        FormulaValue::Number(n) => same(*n),
        FormulaValue::String(s) => crate::coerce::text_to_number(s).is_some_and(same),
        _ => false,
    }

}


#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_criteria() {
        use duke_sheets_core::CellError;
        for criteria in [FormulaValue::Error(CellError::Na), FormulaValue::String("#N/A".into())] {
            let matcher = CriteriaMatcher::new(&criteria);
            assert!(matcher.matches(&FormulaValue::Error(CellError::Na)));
            assert!(!matcher.matches(&FormulaValue::Error(CellError::Value)));
            assert!(!matcher.matches(&FormulaValue::Empty));
        }
    }

    #[test]
    fn test_number_criteria() {
        let matcher = CriteriaMatcher::new(&FormulaValue::Number(5.0));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));
        assert!(matcher.matches(&FormulaValue::String("5".into())));
        assert!(!matcher.matches(&FormulaValue::String("x".into())));
        assert!(!matcher.matches(&FormulaValue::Boolean(true)));
    }

    #[test]
    fn test_comparison_criteria() {
        // Greater than
        let matcher = CriteriaMatcher::new(&FormulaValue::String(">5".into()));
        assert!(matcher.matches(&FormulaValue::Number(6.0)));
        assert!(!matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));

        // Greater than or equal
        let matcher = CriteriaMatcher::new(&FormulaValue::String(">=5".into()));
        assert!(matcher.matches(&FormulaValue::Number(6.0)));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));

        // Less than
        let matcher = CriteriaMatcher::new(&FormulaValue::String("<5".into()));
        assert!(!matcher.matches(&FormulaValue::Number(6.0)));
        assert!(!matcher.matches(&FormulaValue::Number(5.0)));
        assert!(matcher.matches(&FormulaValue::Number(4.0)));

        // Less than or equal
        let matcher = CriteriaMatcher::new(&FormulaValue::String("<=5".into()));
        assert!(!matcher.matches(&FormulaValue::Number(6.0)));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(matcher.matches(&FormulaValue::Number(4.0)));

        // Not equal
        let matcher = CriteriaMatcher::new(&FormulaValue::String("<>5".into()));
        assert!(matcher.matches(&FormulaValue::Number(6.0)));
        assert!(!matcher.matches(&FormulaValue::Number(5.0)));
        assert!(matcher.matches(&FormulaValue::Number(4.0)));

        // Equal
        let matcher = CriteriaMatcher::new(&FormulaValue::String("=5".into()));
        assert!(!matcher.matches(&FormulaValue::Number(6.0)));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));

        let matcher = CriteriaMatcher::new(&FormulaValue::String("<>fruit".into()));
        assert!(matcher.matches(&FormulaValue::String("veg".into())));
        assert!(!matcher.matches(&FormulaValue::String("fruit".into())));
    }

    #[test]
    fn test_text_criteria() {
        // Case insensitive
        let matcher = CriteriaMatcher::new(&FormulaValue::String("apple".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(matcher.matches(&FormulaValue::String("APPLE".into())));
        assert!(matcher.matches(&FormulaValue::String("Apple".into())));
        assert!(!matcher.matches(&FormulaValue::String("banana".into())));
    }

    #[test]
    fn test_wildcard_criteria() {
        // Asterisk wildcard
        let matcher = CriteriaMatcher::new(&FormulaValue::String("a*".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(matcher.matches(&FormulaValue::String("a".into())));
        assert!(!matcher.matches(&FormulaValue::String("banana".into())));

        // Asterisk in middle
        let matcher = CriteriaMatcher::new(&FormulaValue::String("a*e".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(matcher.matches(&FormulaValue::String("ae".into())));
        assert!(!matcher.matches(&FormulaValue::String("apples".into())));

        // Question mark wildcard
        let matcher = CriteriaMatcher::new(&FormulaValue::String("a?ple".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(!matcher.matches(&FormulaValue::String("aple".into())));
        assert!(!matcher.matches(&FormulaValue::String("axxple".into())));

        // Combined wildcards
        let matcher = CriteriaMatcher::new(&FormulaValue::String("a?p*".into()));
        assert!(matcher.matches(&FormulaValue::String("apple".into())));
        assert!(matcher.matches(&FormulaValue::String("app".into())));
        assert!(!matcher.matches(&FormulaValue::String("ap".into())));
    }

    #[test]
    fn test_text_criteria_skip_blanks_and_numbers() {
        let any = CriteriaMatcher::new(&FormulaValue::String("*".into()));
        assert!(any.matches(&FormulaValue::String("x".into())));
        assert!(!any.matches(&FormulaValue::Empty));
        assert!(!any.matches(&FormulaValue::Number(1.0)));
        let before_m = CriteriaMatcher::new(&FormulaValue::String("<m".into()));
        assert!(before_m.matches(&FormulaValue::String("a".into())));
        assert!(!before_m.matches(&FormulaValue::Empty));
        let not_x = CriteriaMatcher::new(&FormulaValue::String("<>x".into()));
        assert!(not_x.matches(&FormulaValue::Empty));
        assert!(not_x.matches(&FormulaValue::Number(1.0)));
        assert!(!not_x.matches(&FormulaValue::String("X".into())));
        let equals = CriteriaMatcher::new(&FormulaValue::String("=".into()));
        assert!(equals.matches(&FormulaValue::Empty));
        assert!(!equals.matches(&FormulaValue::String("x".into())));
    }

    #[test]
    fn test_number_text_criteria_reads_text_as_numbers() {
        let matcher = CriteriaMatcher::new(&FormulaValue::String("5".into()));
        assert!(matcher.matches(&FormulaValue::String("5.0".into())));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
    }

    #[test]
    fn test_empty_criteria() {
        let matcher = CriteriaMatcher::new(&FormulaValue::String("".into()));
        assert!(matcher.matches(&FormulaValue::Empty));
        assert!(matcher.matches(&FormulaValue::String("".into())));
        assert!(!matcher.matches(&FormulaValue::String("text".into())));
        assert!(!matcher.matches(&FormulaValue::Number(0.0)));
    }
}
