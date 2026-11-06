//! Numbers as Excel writes them as text, at the precision it keeps.

/// A number rounded to 15 significant digits, the precision Excel shows, compares and rounds at:
/// `=0.1+0.2=0.3` is TRUE and ROUND(1.005,2) is 1.01.
pub fn round_15(n: f64) -> f64 {
    if n == 0.0 || !n.is_finite() {
        return n;
    }
    format!("{:.14e}", n).parse().unwrap_or(n)
}

/// A number as Excel writes it as text (`&`, CONCATENATE, the formula bar): at most 15
/// significant digits, so `1/3` is "0.333333333333333" and `0.1+0.2` is "0.3", in scientific
/// notation from 1E+15 up and below 1E-09, with `decimal_separator`.
pub fn number_text(n: f64, decimal_separator: char) -> String {
    if n == 0.0 || !n.is_finite() {
        return "0".to_string();
    }
    let rounded = round_15(n);
    let exponent = rounded.abs().log10().floor() as i32;
    let trim = |text: String| -> String {
        if text.contains('.') {
            text.trim_end_matches('0').trim_end_matches('.').to_string()
        } else {
            text
        }
    };
    let text = if !(-9..15).contains(&exponent) {
        let scientific = format!("{:.14e}", rounded);
        let (mantissa, power) = scientific.split_once('e').unwrap_or((&scientific, "0"));
        let power: i32 = power.parse().unwrap_or(0);
        let sign = if power < 0 { '-' } else { '+' };
        format!("{}E{}{:02}", trim(mantissa.to_string()), sign, power.abs())
    } else {
        trim(format!("{:.*}", (14 - exponent).max(0) as usize, rounded))
    };
    with_decimal_separator(text, decimal_separator)
}

/// `text`, a number written with a `.` decimal point, with `separator` instead.
pub fn with_decimal_separator(text: String, separator: char) -> String {
    match separator {
        '.' => text,
        separator => text.replace('.', &separator.to_string()),
    }
}
