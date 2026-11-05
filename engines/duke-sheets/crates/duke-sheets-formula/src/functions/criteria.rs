//! Criteria matching for SUMIF, COUNTIF, AVERAGEIF and related functions
//!
//! Excel criteria can be:
//! - A number: exact match (e.g., 5)
//! - A text string: case-insensitive match (e.g., "apple")
//! - A comparison expression: ">5", ">=10", "<100", "<=50", "<>0", "=5"
//! - Wildcards: "*" matches any characters, "?" matches single character
//! - Empty string: matches empty cells

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
            FormulaValue::Error(_) => CriteriaType::Empty, // Errors don't match anything
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
            CriteriaType::Number(criteria_num) => {
                // Only match actual numeric values, not booleans or strings that look like
                // numbers: SUMIF(A:A, 5) won't match TRUE or the text "5"
                matches!(value, FormulaValue::Number(n) if (n - criteria_num).abs() < 1e-10)
            }

            CriteriaType::NumberOrText(criteria_num, text) => match value {
                FormulaValue::Number(n) => (n - criteria_num).abs() < 1e-10,
                FormulaValue::String(s) => s.trim().to_lowercase() == *text,
                _ => false,
            },

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

            CriteriaType::Text(pattern) => {
                let text = value.as_string().to_lowercase();
                Self::wildcard_match(pattern, &text)
            }

            CriteriaType::TextComparison(op, pattern) => {
                let text = value.as_string().to_lowercase();
                let matched = Self::wildcard_match(pattern, &text);
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
        }
    }

    /// Match with wildcards: * = any characters, ? = single character, ~ makes the next
    /// character literal (`a~*` matches the text "a*").
    fn wildcard_match(pattern: &str, text: &str) -> bool {
        if !pattern.contains(['*', '?', '~']) {
            return pattern == text;
        }

        let mut tokens = Vec::new();
        let mut chars = pattern.chars();
        while let Some(c) = chars.next() {
            tokens.push(match c {
                '*' => Wild::Run,
                '?' => Wild::One,
                '~' => Wild::Char(chars.next().unwrap_or('~')),
                other => Wild::Char(other),
            });
        }
        let text_chars: Vec<char> = text.chars().collect();

        Self::wildcard_match_impl(&tokens, &text_chars)
    }

    fn wildcard_match_impl(pattern: &[Wild], text: &[char]) -> bool {
        let mut pi = 0; // pattern index
        let mut ti = 0; // text index
        let mut star_pi = None; // position of last * in pattern
        let mut star_ti = 0; // position in text when we matched last *

        while ti < text.len() {
            let matches_here = match pattern.get(pi) {
                Some(Wild::One) => true,
                Some(Wild::Char(c)) => *c == text[ti],
                _ => false,
            };
            if matches_here {
                pi += 1;
                ti += 1;
            } else if matches!(pattern.get(pi), Some(Wild::Run)) {
                // Star found, record position
                star_pi = Some(pi);
                star_ti = ti;
                pi += 1; // Try matching * with empty string first
            } else if let Some(sp) = star_pi {
                // Mismatch, but we have a star to backtrack to
                pi = sp + 1;
                star_ti += 1;
                ti = star_ti;
            } else {
                // No match and no star to backtrack to
                return false;
            }
        }

        // Check remaining pattern characters (must all be *)
        while matches!(pattern.get(pi), Some(Wild::Run)) {
            pi += 1;
        }

        pi == pattern.len()
    }
}

/// One element of a wildcard pattern.
#[derive(Clone, Copy)]
enum Wild {
    One,
    Run,
    Char(char),
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number_criteria() {
        let matcher = CriteriaMatcher::new(&FormulaValue::Number(5.0));
        assert!(matcher.matches(&FormulaValue::Number(5.0)));
        assert!(!matcher.matches(&FormulaValue::Number(4.0)));
        assert!(!matcher.matches(&FormulaValue::String("5".into())));
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
    fn test_empty_criteria() {
        let matcher = CriteriaMatcher::new(&FormulaValue::String("".into()));
        assert!(matcher.matches(&FormulaValue::Empty));
        assert!(matcher.matches(&FormulaValue::String("".into())));
        assert!(!matcher.matches(&FormulaValue::String("text".into())));
        assert!(!matcher.matches(&FormulaValue::Number(0.0)));
    }
}
