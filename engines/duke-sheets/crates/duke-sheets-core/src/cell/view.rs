//! Cell view - a lightweight borrow wrapper that provides access to
//! a cell's value, style, and formatted display string.

use crate::cell::value::CellValue;
use crate::style::{NumberFormat, Style};
use crate::number_text::{number_text, with_decimal_separator};
use crate::style::FormatKind;
use crate::{RegionalSettings, WorkbookSettings};

/// A read-only view of a cell, providing access to its value, style,
/// and formatted display string.
///
/// This is a lightweight borrow - no allocation or copying.
/// Obtain one via [`Worksheet::cell_view_at`] or [`Worksheet::cell_view`].
///
/// # Example
///
/// ```ignore
/// let cell = sheet.cell_view_at(0, 0);
/// println!("Raw: {}", cell.value());
/// println!("Formatted: {}", cell.formatted());
/// ```
pub struct CellView<'a> {
    value: &'a CellValue,
    style: Option<&'a Style>,
    date_1904: bool,
    regional: RegionalSettings,
}

impl<'a> CellView<'a> {
    /// Create a new cell view, formatting in en-US.
    pub(crate) fn new(value: &'a CellValue, style: Option<&'a Style>, date_1904: bool) -> Self {
        Self {
            value,
            style,
            date_1904,
            regional: RegionalSettings::default(),
        }
    }

    /// The view formatting in the workbook's settings: its regional settings (separators, month
    /// and day names, the short date) and its date system.
    pub fn with_settings(self, settings: &WorkbookSettings) -> Self {
        Self { regional: settings.regional, date_1904: settings.date_1904, ..self }
    }

    /// The cell's value.
    pub fn value(&self) -> &CellValue {
        self.value
    }

    /// The cell's style, if non-default.
    pub fn style(&self) -> Option<&Style> {
        self.style
    }

    /// The cell's number format (from the style, or `General` if no style).
    pub fn number_format(&self) -> &NumberFormat {
        self.style
            .map(|s| &s.number_format)
            .unwrap_or(&NumberFormat::General)
    }

    /// Format the cell value for display, applying the cell's number format.
    ///
    /// - Numbers are formatted according to the cell's `NumberFormat`
    ///   (percentages, currencies, dates, times, scientific notation, etc.)
    /// - Date serial numbers are converted to human-readable dates when the
    ///   format is a date/time format
    /// - Strings, booleans, and errors display as-is
    /// - Empty cells return an empty string
    pub fn formatted(&self) -> String {
        format_cell_value(self.value, self.number_format(), self.date_1904, &self.regional)
    }

    /// The value as text to edit, as Excel's formula bar shows it, which typed back unchanged
    /// gives the same value (to the 15 digits Excel keeps): a number with the regional decimal
    /// separator, in a percent format with `%` (`12.5%`), and in a date or time format as the
    /// short date and the time (`05/11/2026 13:30:00`, `1:30:00 PM` where the format has AM/PM,
    /// `25:30:00` in a duration format). Other values show as they are.
    pub fn input_text(&self) -> String {
        let CellValue::Number(n) = *self.value.effective_value() else {
            return format_cell_value(self.value, &NumberFormat::General, self.date_1904, &self.regional);
        };
        let format = self.number_format();
        let separator = self.regional.decimal_separator;
        // The last day a date format shows, 31 December 9999.
        let last_day = if self.date_1904 { 2_957_003.0 } else { 2_958_465.0 };
        let kind = format.kind();
        match kind {
            FormatKind::Percent => return format!("{}%", number_text(n * 100.0, separator)),
            FormatKind::Date | FormatKind::Time | FormatKind::DateTime | FormatKind::Duration
                if (0.0..=last_day).contains(&n) => {}
            _ => return number_text(n, separator),
        }
        let twelve_hour = ssfmt::NumberFormat::parse(format.format_string()).is_ok_and(|f| f.has_am_pm());
        let time = if twelve_hour { "h:mm:ss AM/PM" } else { "h:mm:ss" };
        // Day 0 is a date in the 1904 system (1 January 1904) but not in the 1900 one.
        let with_date = n >= 1.0 || (self.date_1904 && kind != FormatKind::Time);
        let code = if kind == FormatKind::Duration {
            "[h]:mm:ss".to_string()
        } else if !with_date {
            time.to_string()
        } else if n.fract() == 0.0 {
            self.regional.short_date_code()
        } else {
            format!("{} {time}", self.regional.short_date_code())
        };
        format_number(n, &NumberFormat::Custom(code), self.date_1904, &self.regional)
    }
}

/// Format a cell value using the given number format, date system, and regional settings.
///
/// This is the standalone version of [`CellView::formatted`], useful when
/// you have the value and format separately.
pub fn format_cell_value(
    value: &CellValue,
    format: &NumberFormat,
    date_1904: bool,
    regional: &RegionalSettings,
) -> String {
    let effective = value.effective_value();

    match effective {
        CellValue::Number(n) => format_number(*n, format, date_1904, regional),
        CellValue::String(s) => {
            // Apply the format's text section, if it has one (`0;-0;0;"x "@`, `"Code: "@`)
            let code = format.format_string();
            if code == "General" || code == "@" {
                return s.to_string();
            }
            match ssfmt::NumberFormat::parse(&code) {
                Ok(parsed) => parsed.format_text(s.as_str(), &ssfmt::FormatOptions::default()),
                Err(_) => s.to_string(),
            }
        }
        CellValue::Boolean(b) => {
            if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        CellValue::Error(e) => e.to_string(),
        CellValue::Empty => String::new(),
        CellValue::SpillTarget { .. } => String::new(),
        CellValue::RichText(runs) => crate::rich_text::rich_text_to_plain(runs),
    }
}

/// Format a numeric value using an Excel number format code.
fn format_number(
    value: f64,
    format: &NumberFormat,
    date_1904: bool,
    regional: &RegionalSettings,
) -> String {
    let opts = ssfmt::FormatOptions {
        date_system: if date_1904 {
            ssfmt::DateSystem::Date1904
        } else {
            ssfmt::DateSystem::Date1900
        },
        locale: regional.format_locale(),
    };

    match format {
        NumberFormat::General => format_general(value, regional.decimal_separator),
        NumberFormat::BuiltIn(_) | NumberFormat::Custom(_) => {
            // Built-in IDs come from our builtin_format_string() table, and the short date (14) and
            // date with time (22) from the regional settings (display_code). ssfmt::format_with_id
            // uses SheetJS's table, which diverges from ECMA-376 on locale-dependent IDs like 14.
            let code = format.display_code(regional);
            match ssfmt::format(value, &code, &opts) {
                Ok(s) => s,
                Err(_) => format_general(value, regional.decimal_separator),
            }
        }
    }
}

/// General format: display numbers without trailing zeros, integers
/// without decimal point, with the regional decimal separator.
fn format_general(value: f64, decimal_separator: char) -> String {
    if value.is_nan() || value.is_infinite() {
        return value.to_string();
    }
    // Match Excel's General format: up to 11 significant digits,
    // no trailing zeros, scientific notation for very large/small values.
    let abs = value.abs();
    let text = if value == 0.0 {
        "0".to_string()
    } else if abs >= 1e11 || abs < 1e-4 {
        match ssfmt::format_default(value, "0.#####E+00") {
            Ok(s) => s,
            Err(_) => format!("{:E}", value),
        }
    } else if value.fract() == 0.0 && abs < 1e15 {
        format!("{}", value as i64)
    } else {
        format!("{:.10}", value).trim_end_matches('0').trim_end_matches('.').to_string()
    };
    with_decimal_separator(text, decimal_separator)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::value::CellError;

    // Helper
    fn fmt(value: &CellValue, format: &NumberFormat, date_1904: bool) -> String {
        format_cell_value(value, format, date_1904, &RegionalSettings::default())
    }

    // General format

    #[test]
    fn general_integer() {
        assert_eq!(
            fmt(&CellValue::Number(42.0), &NumberFormat::General, false),
            "42"
        );
    }

    #[test]
    fn general_zero() {
        assert_eq!(
            fmt(&CellValue::Number(0.0), &NumberFormat::General, false),
            "0"
        );
    }

    #[test]
    fn general_negative_integer() {
        assert_eq!(
            fmt(&CellValue::Number(-7.0), &NumberFormat::General, false),
            "-7"
        );
    }

    #[test]
    fn general_decimal() {
        assert_eq!(
            fmt(&CellValue::Number(3.14), &NumberFormat::General, false),
            "3.14"
        );
    }

    #[test]
    fn general_no_trailing_zeros() {
        assert_eq!(
            fmt(&CellValue::Number(1.50), &NumberFormat::General, false),
            "1.5"
        );
    }

    #[test]
    fn general_large_number_scientific() {
        let result = fmt(&CellValue::Number(1e12), &NumberFormat::General, false);
        // Should use scientific notation for numbers >= 1e11
        assert!(
            result.contains('E') || result.contains('e'),
            "expected scientific: {}",
            result
        );
    }

    #[test]
    fn general_small_number_scientific() {
        let result = fmt(&CellValue::Number(0.00001), &NumberFormat::General, false);
        assert!(
            result.contains('E') || result.contains('e'),
            "expected scientific: {}",
            result
        );
    }

    #[test]
    fn general_nan() {
        assert_eq!(
            fmt(&CellValue::Number(f64::NAN), &NumberFormat::General, false),
            "NaN"
        );
    }

    #[test]
    fn general_infinity() {
        assert_eq!(
            fmt(
                &CellValue::Number(f64::INFINITY),
                &NumberFormat::General,
                false
            ),
            "inf"
        );
    }

    // Built-in number formats

    #[test]
    fn builtin_integer_format() {
        // Format ID 1: "0"
        let result = fmt(
            &CellValue::Number(1234.567),
            &NumberFormat::BuiltIn(1),
            false,
        );
        assert_eq!(result, "1235");
    }

    #[test]
    fn builtin_decimal_format() {
        // Format ID 2: "0.00"
        let result = fmt(&CellValue::Number(1234.5), &NumberFormat::BuiltIn(2), false);
        assert_eq!(result, "1234.50");
    }

    #[test]
    fn builtin_thousands_separator() {
        // Format ID 3: "#,##0"
        let result = fmt(
            &CellValue::Number(1234567.0),
            &NumberFormat::BuiltIn(3),
            false,
        );
        assert_eq!(result, "1,234,567");
    }

    #[test]
    fn builtin_thousands_decimal() {
        // Format ID 4: "#,##0.00"
        let result = fmt(
            &CellValue::Number(1234567.89),
            &NumberFormat::BuiltIn(4),
            false,
        );
        assert_eq!(result, "1,234,567.89");
    }

    // Percentage formats

    #[test]
    fn builtin_percent_integer() {
        // Format ID 9: "0%"
        let result = fmt(&CellValue::Number(0.75), &NumberFormat::BuiltIn(9), false);
        assert_eq!(result, "75%");
    }

    #[test]
    fn builtin_percent_decimal() {
        // Format ID 10: "0.00%"
        let result = fmt(
            &CellValue::Number(0.1234),
            &NumberFormat::BuiltIn(10),
            false,
        );
        assert_eq!(result, "12.34%");
    }

    // Scientific notation format

    #[test]
    fn builtin_scientific() {
        // Format ID 11: "0.00E+00"
        let result = fmt(
            &CellValue::Number(12345.0),
            &NumberFormat::BuiltIn(11),
            false,
        );
        assert_eq!(result, "1.23E+04");
    }

    // Date formats (1900 date system)

    #[test]
    fn builtin_date_short_1900() {
        // Format ID 14: the short date, m/d/yyyy in en-US
        // Serial 44927 = 2023-01-01 in 1900 system
        let result = fmt(
            &CellValue::Number(44927.0),
            &NumberFormat::BuiltIn(14),
            false,
        );
        assert_eq!(result, "1/1/2023");
    }

    #[test]
    fn builtin_date_medium() {
        // Format ID 15: "d-mmm-yy"
        let result = fmt(
            &CellValue::Number(44927.0),
            &NumberFormat::BuiltIn(15),
            false,
        );
        assert_eq!(result, "1-Jan-23");
    }

    #[test]
    fn builtin_date_day_month() {
        // Format ID 16: "d-mmm"
        let result = fmt(
            &CellValue::Number(44927.0),
            &NumberFormat::BuiltIn(16),
            false,
        );
        assert_eq!(result, "1-Jan");
    }

    #[test]
    fn builtin_date_month_year() {
        // Format ID 17: "mmm-yy"
        let result = fmt(
            &CellValue::Number(44927.0),
            &NumberFormat::BuiltIn(17),
            false,
        );
        assert_eq!(result, "Jan-23");
    }

    // Date formats (1904 date system)

    #[test]
    fn builtin_date_short_1904() {
        // In 1904 system, serial 43465 = 2023-01-01
        // (1904 system starts 1462 days later than 1900 system)
        let result = fmt(
            &CellValue::Number(43465.0),
            &NumberFormat::BuiltIn(14),
            true,
        );
        assert_eq!(result, "1/1/2023");
    }

    // Time formats

    #[test]
    fn builtin_time_ampm() {
        // Format ID 18: "h:mm AM/PM"
        // 0.75 = 18:00 = 6:00 PM
        let result = fmt(&CellValue::Number(0.75), &NumberFormat::BuiltIn(18), false);
        assert_eq!(result, "6:00 PM");
    }

    #[test]
    fn builtin_time_24h() {
        // Format ID 20: "h:mm"
        let result = fmt(&CellValue::Number(0.75), &NumberFormat::BuiltIn(20), false);
        assert_eq!(result, "18:00");
    }

    #[test]
    fn builtin_time_24h_seconds() {
        // Format ID 21: "h:mm:ss"
        // 0.5 = 12:00:00 noon
        let result = fmt(&CellValue::Number(0.5), &NumberFormat::BuiltIn(21), false);
        assert_eq!(result, "12:00:00");
    }

    // Custom format strings

    #[test]
    fn custom_percent_format() {
        let fmt_code = NumberFormat::Custom("0.0%".to_string());
        let result = fmt(&CellValue::Number(0.1234), &fmt_code, false);
        assert_eq!(result, "12.3%");
    }

    #[test]
    fn custom_thousands_format() {
        let fmt_code = NumberFormat::Custom("#,##0.00".to_string());
        let result = fmt(&CellValue::Number(9876.5), &fmt_code, false);
        assert_eq!(result, "9,876.50");
    }

    #[test]
    fn custom_date_format() {
        let fmt_code = NumberFormat::Custom("yyyy-mm-dd".to_string());
        // Serial 44927 = 2023-01-01
        let result = fmt(&CellValue::Number(44927.0), &fmt_code, false);
        assert_eq!(result, "2023-01-01");
    }

    #[test]
    fn custom_currency_format() {
        let fmt_code = NumberFormat::Custom("$#,##0.00".to_string());
        let result = fmt(&CellValue::Number(1234.5), &fmt_code, false);
        assert_eq!(result, "$1,234.50");
    }

    // Non-numeric cell types (passthrough)

    #[test]
    fn string_passthrough() {
        let val = CellValue::string("Hello World");
        assert_eq!(fmt(&val, &NumberFormat::General, false), "Hello World");
    }

    #[test]
    fn string_ignores_number_format() {
        let val = CellValue::string("Not a number");
        let fmt_code = NumberFormat::Custom("#,##0.00".to_string());
        assert_eq!(fmt(&val, &fmt_code, false), "Not a number");
    }

    #[test]
    fn boolean_true() {
        assert_eq!(
            fmt(&CellValue::Boolean(true), &NumberFormat::General, false),
            "TRUE"
        );
    }

    #[test]
    fn boolean_false() {
        assert_eq!(
            fmt(&CellValue::Boolean(false), &NumberFormat::General, false),
            "FALSE"
        );
    }

    #[test]
    fn error_value() {
        assert_eq!(
            fmt(
                &CellValue::Error(CellError::Value),
                &NumberFormat::General,
                false
            ),
            "#VALUE!"
        );
    }

    #[test]
    fn error_ref() {
        assert_eq!(
            fmt(
                &CellValue::Error(CellError::Ref),
                &NumberFormat::General,
                false
            ),
            "#REF!"
        );
    }

    #[test]
    fn error_div0() {
        assert_eq!(
            fmt(
                &CellValue::Error(CellError::Div0),
                &NumberFormat::General,
                false
            ),
            "#DIV/0!"
        );
    }

    #[test]
    fn empty_cell() {
        assert_eq!(fmt(&CellValue::Empty, &NumberFormat::General, false), "");
    }

    #[test]
    fn cached_formula_number_formats_as_number() {
        let val = CellValue::Number(42.0);
        assert_eq!(fmt(&val, &NumberFormat::General, false), "42");
    }

    #[test]
    fn cached_formula_number_uses_cell_format() {
        let val = CellValue::Number(0.85);
        let fmt_code = NumberFormat::BuiltIn(9); // "0%"
        assert_eq!(fmt(&val, &fmt_code, false), "85%");
    }

    #[test]
    fn cached_formula_string_formats_as_string() {
        let val = CellValue::string("yes");
        assert_eq!(fmt(&val, &NumberFormat::General, false), "yes");
    }

    #[test]
    fn cached_formula_error_formats_as_error() {
        let val = CellValue::Error(CellError::Div0);
        assert_eq!(fmt(&val, &NumberFormat::General, false), "#DIV/0!");
    }

    #[test]
    fn uncached_formula_formats_as_empty_cell_value() {
        let val = CellValue::Empty;
        assert_eq!(fmt(&val, &NumberFormat::General, false), "");
    }

    // CellView API

    #[test]
    fn cell_view_value_accessor() {
        let val = CellValue::Number(3.14);
        let view = CellView::new(&val, None, false);
        assert_eq!(view.value(), &CellValue::Number(3.14));
    }

    #[test]
    fn cell_view_default_format() {
        let val = CellValue::Number(3.14);
        let view = CellView::new(&val, None, false);
        assert_eq!(view.number_format(), &NumberFormat::General);
    }

    #[test]
    fn cell_view_with_style() {
        let val = CellValue::Number(0.75);
        let style = Style {
            number_format: NumberFormat::BuiltIn(9), // "0%"
            ..Default::default()
        };
        let view = CellView::new(&val, Some(&style), false);
        assert_eq!(view.formatted(), "75%");
    }

    #[test]
    fn cell_view_style_accessor() {
        let val = CellValue::Number(1.0);
        let style = Style {
            number_format: NumberFormat::BuiltIn(2),
            ..Default::default()
        };
        let view = CellView::new(&val, Some(&style), false);
        assert!(view.style().is_some());

        let view_no_style = CellView::new(&val, None, false);
        assert!(view_no_style.style().is_none());
    }

    // Accounting / currency built-in formats

    #[test]
    fn builtin_currency_format_5() {
        // Format ID 5: "$#,##0_);($#,##0)"
        let result = fmt(&CellValue::Number(1234.0), &NumberFormat::BuiltIn(5), false);
        assert!(result.contains("1,234"), "expected currency: {}", result);
        assert!(result.contains('$'), "expected dollar sign: {}", result);
    }

    #[test]
    fn builtin_accounting_negative() {
        // Format ID 37: "#,##0_);(#,##0)" - negative in parens
        let result = fmt(
            &CellValue::Number(-1234.0),
            &NumberFormat::BuiltIn(37),
            false,
        );
        assert!(result.contains("1,234"), "expected number: {}", result);
        assert!(
            result.contains('('),
            "expected parens for negative: {}",
            result
        );
    }

    // Edge cases

    #[test]
    fn negative_zero() {
        // -0.0 should display as "0" under General format
        let result = format_general(-0.0, '.');
        assert_eq!(result, "0");
    }

    #[test]
    fn very_small_decimal() {
        let result = fmt(&CellValue::Number(0.1 + 0.2), &NumberFormat::General, false);
        // Should produce "0.3" or close to it, not "0.30000000000000004"
        assert!(result.starts_with("0.3"), "expected ~0.3: {}", result);
    }

    #[test]
    fn spill_target_empty() {
        let val = CellValue::SpillTarget {
            source_row: 0,
            source_col: 0,
            offset_row: 1,
            offset_col: 0,
        };
        assert_eq!(fmt(&val, &NumberFormat::General, false), "");
    }

    #[test]
    fn text_format_id_49() {
        // Format ID 49: "@" - text format, numbers should still format
        let result = fmt(&CellValue::Number(42.0), &NumberFormat::BuiltIn(49), false);
        // ssfmt with "@" format on a number typically returns the number as text
        assert!(!result.is_empty(), "expected non-empty for text format");
    }

    #[test]
    fn date_format_with_underscore_paren_no_trailing_paren() {
        // Real-world Excel files often use formats like m/d/yyyy_) where _)
        // means "pad with width of ')'". The ')' must NOT appear in output.
        // Serial 46022 = 12/31/2025 in the 1900 date system.
        let val = CellValue::Number(46022.0);
        let fmt_code = NumberFormat::Custom("m/d/yyyy_)".to_string());
        let result = fmt(&val, &fmt_code, false);
        assert_eq!(result, "12/31/2025 ", "underscore-paren should produce a space, not a literal ')'");
    }

    fn fmt_in(tag: &str, value: f64, format: NumberFormat) -> String {
        format_cell_value(&CellValue::Number(value), &format, false, &RegionalSettings::from_tag(tag))
    }

    #[test]
    fn short_dates_show_in_the_date_order() {
        // 46331.5625 is 5 November 2026, 13:30.
        for (tag, short, with_time) in [
            ("en-US", "11/5/2026", "11/5/2026 13:30"),
            ("en-GB", "05/11/2026", "05/11/2026 13:30"),
            ("de-DE", "05.11.2026", "05.11.2026 13:30"),
            ("sv-SE", "2026-11-05", "2026-11-05 13:30"),
        ] {
            assert_eq!(fmt_in(tag, 46331.0, NumberFormat::BuiltIn(14)), short, "{tag}");
            assert_eq!(fmt_in(tag, 46331.5625, NumberFormat::BuiltIn(22)), with_time, "{tag}");
        }
    }

    #[test]
    fn numbers_show_with_the_regional_separators() {
        assert_eq!(fmt_in("de-DE", 1.5, NumberFormat::General), "1,5");
        assert_eq!(fmt_in("de-DE", 1234.5, NumberFormat::Custom("#,##0.00".into())), "1.234,50");
        assert_eq!(fmt_in("de-DE", 0.125, NumberFormat::BuiltIn(10)), "12,50%");
        assert_eq!(fmt_in("en-US", 1234.5, NumberFormat::Custom("#,##0.00".into())), "1,234.50");
        // Fractional seconds take the regional decimal separator (90.5 seconds).
        assert_eq!(fmt_in("de-DE", 90.5 / 86400.0, NumberFormat::BuiltIn(47)), "01:30,5");
        // In a date format a period or comma is the character itself.
        assert_eq!(fmt_in("de-DE", 46331.0, NumberFormat::Custom("dd.mm.yyyy".into())), "05.11.2026");
        assert_eq!(fmt_in("de-DE", 46331.0, NumberFormat::Custom("d. mmmm yyyy".into())), "5. November 2026");
        assert_eq!(fmt_in("fr-FR", 46331.0, NumberFormat::Custom("d mmm, yyyy".into())), "5 nov., 2026");
    }

    #[test]
    fn input_text_reads_back_in_the_regional_settings() {
        let input = |tag: &str, value: f64, format: NumberFormat| {
            let style = Style { number_format: format, ..Default::default() };
            let value = CellValue::Number(value);
            let settings = WorkbookSettings { regional: RegionalSettings::from_tag(tag), ..Default::default() };
            CellView::new(&value, Some(&style), false).with_settings(&settings).input_text()
        };
        assert_eq!(input("en-US", 1.5, NumberFormat::General), "1.5");
        assert_eq!(input("de-DE", 1234.5, NumberFormat::Custom("#,##0.00".into())), "1234,5");
        assert_eq!(input("en-US", 0.1 + 0.2, NumberFormat::General), "0.3");
        assert_eq!(input("en-US", 1e300, NumberFormat::General), "1E+300");
        assert_eq!(input("en-US", 1.5e-12, NumberFormat::General), "1.5E-12");
        assert_eq!(input("en-US", 0.125, NumberFormat::Custom("0.##########%".into())), "12.5%");
        assert_eq!(input("de-DE", 0.125, NumberFormat::BuiltIn(10)), "12,5%");
        assert_eq!(input("en-US", 0.5625, NumberFormat::BuiltIn(18)), "1:30:00 PM");
        assert_eq!(input("en-US", 46331.5625, NumberFormat::Custom("m/d/yyyy h:mm AM/PM".into())), "11/5/2026 1:30:00 PM");
        assert_eq!(input("en-US", 25.5 / 24.0, NumberFormat::BuiltIn(46)), "25:30:00");
        assert_eq!(input("en-US", 60.0, NumberFormat::BuiltIn(14)), "2/29/1900");
        assert_eq!(input("en-US", 3_000_000.0, NumberFormat::BuiltIn(14)), "3000000");
        let in_1904 = |value: f64| {
            let style = Style { number_format: NumberFormat::BuiltIn(14), ..Default::default() };
            let value = CellValue::Number(value);
            let settings = WorkbookSettings { date_1904: true, ..Default::default() };
            CellView::new(&value, Some(&style), false).with_settings(&settings).input_text()
        };
        assert_eq!(in_1904(0.0), "1/1/1904");
        assert_eq!(input("en-US", 46331.0, NumberFormat::BuiltIn(14)), "11/5/2026");
        assert_eq!(input("en-GB", 46331.0, NumberFormat::Custom("d mmm yyyy".into())), "05/11/2026");
        assert_eq!(input("en-GB", 46331.5625, NumberFormat::BuiltIn(22)), "05/11/2026 13:30:00");
        assert_eq!(input("en-GB", 0.5625, NumberFormat::BuiltIn(20)), "13:30:00");
        assert_eq!(input("en-US", 0.0, NumberFormat::BuiltIn(14)), "0:00:00");
        let text = CellValue::string("5/11/2026");
        assert_eq!(CellView::new(&text, None, false).input_text(), "5/11/2026");
    }

    #[test]
    fn formats_have_kinds() {
        use crate::style::FormatKind::*;
        let kind = |code: &str| NumberFormat::Custom(code.into()).kind();
        assert_eq!(NumberFormat::General.kind(), General);
        assert_eq!(NumberFormat::BuiltIn(14).kind(), Date);
        assert_eq!(NumberFormat::BuiltIn(22).kind(), DateTime);
        assert_eq!(NumberFormat::BuiltIn(21).kind(), Time);
        assert_eq!(NumberFormat::BuiltIn(18).kind(), Time);
        assert_eq!(NumberFormat::BuiltIn(46).kind(), Duration);
        assert_eq!(NumberFormat::BuiltIn(10).kind(), Percent);
        assert_eq!(NumberFormat::BuiltIn(49).kind(), Text);
        assert_eq!(kind("[m]:ss"), Duration);
        assert_eq!(kind("[s]"), Duration);
        assert_eq!(kind("dd mmm yyyy"), Date);
        assert_eq!(kind("#,##0.00"), Number);
        assert_eq!(kind("0.0%"), Percent);
    }
}
