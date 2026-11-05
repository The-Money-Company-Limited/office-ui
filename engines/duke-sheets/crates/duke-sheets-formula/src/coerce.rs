//! Text read as a number, the way Excel reads it wherever a number is needed (`="12%"*2`,
//! `VALUE("1,234")`, `MONTH("2024-03-15")`, `="6:30 PM"+0`): numbers with thousands separators,
//! a currency sign, a percent sign or parentheses for negatives, dates, times, and dates with
//! times. Dates are 1900-system serials. The workbook's regional settings (crate::regional) decide
//! the order of day and month, the month names besides English ones, and the decimal separator.

use chrono::{Datelike, NaiveDate};
use duke_sheets_core::DateOrder;

use crate::regional::month_named;

/// The number `text` stands for, or `None` when Excel would give #VALUE!.
pub fn text_to_number(text: &str) -> Option<f64> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }
    plain_number(text)
        .or_else(|| date_time(text))
        .filter(|n| n.is_finite())
}

/// `1234`, `-1,234.5`, `1.2E+3`, `$1,234`, `(5)`, `12%`, `+3`.
fn plain_number(text: &str) -> Option<f64> {
    let (negated, text) = match text.strip_prefix('(').and_then(|t| t.strip_suffix(')')) {
        Some(inner) => (true, inner.trim()),
        None => (false, text),
    };
    let (percent, text) = match text.strip_suffix('%') {
        Some(rest) => (true, rest.trim_end()),
        None => (false, text),
    };
    let (sign, text) = match text.as_bytes().first() {
        Some(b'-') => (-1.0, text[1..].trim_start()),
        Some(b'+') => (1.0, text[1..].trim_start()),
        _ => (1.0, text),
    };
    let text = text.strip_prefix('$').map(str::trim_start).unwrap_or(text);
    let settings = crate::regional::current();
    let local: String;
    let text = if settings.decimal_separator == '.' {
        text
    } else {
        // Where 1.234,5 (or 1 234,5) is 1234.5, read it as 1,234.5.
        let space_grouped = settings.thousands_separator == '\u{a0}';
        local = text
            .chars()
            .map(|c| match c {
                c if c == settings.decimal_separator => '.',
                c if c == settings.thousands_separator || (space_grouped && matches!(c, ' ' | '\u{202f}')) => ',',
                '.' | ',' => '\0',
                c => c,
            })
            .collect();
        &local
    };
    if text.is_empty() || !text.starts_with(|c: char| c.is_ascii_digit() || c == '.') {
        return None;
    }
    let (mantissa, exponent) = match text.find(['e', 'E']) {
        Some(i) => (&text[..i], Some(&text[i + 1..])),
        None => (text, None),
    };
    let (whole, fraction) = match mantissa.split_once('.') {
        Some((whole, fraction)) => (whole, Some(fraction)),
        None => (mantissa, None),
    };
    if whole.contains(',') {
        // Thousands separators sit between groups of three digits.
        let mut groups = whole.split(',');
        let first = groups.next()?;
        if first.is_empty() || first.len() > 3 || !groups.all(|g| g.len() == 3) {
            return None;
        }
    }
    let digits = whole.replace(',', "");
    if !digits.chars().all(|c| c.is_ascii_digit())
        || !fraction.is_none_or(|f| f.chars().all(|c| c.is_ascii_digit()))
        || (digits.is_empty() && fraction.is_none_or(str::is_empty))
    {
        return None;
    }
    let exponent: i32 = match exponent {
        Some(e) => e.strip_prefix('+').unwrap_or(e).parse().ok()?,
        None => 0,
    };
    let mut number: f64 =
        format!("{}.{}e{}", if digits.is_empty() { "0" } else { &digits }, fraction.unwrap_or("0"), exponent)
            .parse()
            .ok()?;
    if percent {
        number /= 100.0;
    }
    Some(if negated { -number } else { sign * number })
}

/// The date in text that is a date, or a date followed by a time (DATEVALUE).
pub fn date_text(text: &str) -> Option<f64> {
    let text = text.trim();
    date(text).or_else(|| split_date_time(text).map(|(day, _)| day))
}

/// The time in text that is a time, or a date followed by a time (TIMEVALUE).
pub fn time_text(text: &str) -> Option<f64> {
    let text = text.trim();
    time_of_day(text).or_else(|| split_date_time(text).map(|(_, time)| time))
}

/// "2024-03-15 18:00", "3/15/2024 6:30 PM": the time starts at the first space that leaves a
/// valid date before it.
fn split_date_time(text: &str) -> Option<(f64, f64)> {
    text.match_indices(' ').find_map(|(i, _)| Some((date(text[..i].trim())?, time_of_day(text[i..].trim())?)))
}

/// A date, a time, or a date followed by a time.
fn date_time(text: &str) -> Option<f64> {
    if let Some(time) = time_of_day(text) {
        return Some(time);
    }
    if let Some(date) = date(text) {
        return Some(date);
    }
    split_date_time(text).map(|(day, time)| day + time)
}

/// `18:30`, `18:30:15`, `6:30 PM`, `6 PM`, `1:05:00.5 am`: a fraction of a day.
fn time_of_day(text: &str) -> Option<f64> {
    let lower = text.to_ascii_lowercase();
    // "8 a" is 8 AM, but "8A" (a room, a grade) is text.
    let letter = |suffix: char| lower.strip_suffix(suffix).filter(|rest| rest.ends_with(char::is_whitespace));
    let (clock, meridiem) = if let Some(rest) = lower.strip_suffix("am").or_else(|| letter('a')) {
        (rest.trim_end(), Some(false))
    } else if let Some(rest) = lower.strip_suffix("pm").or_else(|| letter('p')) {
        (rest.trim_end(), Some(true))
    } else {
        (lower.as_str(), None)
    };
    let parts: Vec<&str> = clock.split(':').collect();
    if parts.len() > 3 || (parts.len() == 1 && meridiem.is_none()) {
        return None;
    }
    let number = |s: &str| (!s.is_empty() && s.chars().all(|c| c.is_ascii_digit())).then(|| s.parse::<u32>().ok()).flatten();
    let mut hours = number(parts[0])?;
    let minutes = parts.get(1).map_or(Some(0), |s| number(s))?;
    let seconds: f64 = match parts.get(2) {
        Some(s) => {
            let (whole, fraction) = s.split_once('.').unwrap_or((s, ""));
            if number(whole).is_none() || !fraction.chars().all(|c| c.is_ascii_digit()) {
                return None;
            }
            s.parse().ok()?
        }
        None => 0.0,
    };
    if minutes > 59 || seconds >= 60.0 {
        return None;
    }
    if let Some(pm) = meridiem {
        if hours == 0 || hours > 12 {
            return None;
        }
        hours = hours % 12 + if pm { 12 } else { 0 };
    }
    Some((hours as f64 * 3600.0 + minutes as f64 * 60.0 + seconds) / 86400.0)
}

/// `2024-03-15`, `2024/3/15`, `3/15/2024`, `3-15-24`, `15-Mar-2024`, `15 March 2024`,
/// `Mar 15, 2024`, `March 15 2024`, `Mar-2024` (the 1st), `1Jan2016` (a month name needs no separators).
/// Three numbers are month first, day first (`15/3/2024`, and `15.3.2024` where the decimal
/// separator is a comma) or year first, as the regional settings order them; a four-digit first
/// number is always the year.
fn date(text: &str) -> Option<f64> {
    let settings = crate::regional::current();
    let dotted = settings.date_order == DateOrder::DayFirst && settings.decimal_separator == ',';
    let boundary = |a: u8, b: u8| a.is_ascii_alphanumeric() && b.is_ascii_alphanumeric() && a.is_ascii_digit() != b.is_ascii_digit();
    let spaced: std::borrow::Cow<str> = if text.as_bytes().windows(2).any(|w| boundary(w[0], w[1])) {
        let mut spaced = String::with_capacity(text.len() + 4);
        let mut last = 0u8;
        for c in text.chars() {
            if c.is_ascii() && boundary(last, c as u8) {
                spaced.push(' ');
            }
            last = if c.is_ascii() { c as u8 } else { 0 };
            spaced.push(c);
        }
        spaced.into()
    } else {
        text.into()
    };
    let parts: Vec<&str> = spaced
        .split(|c: char| c == '/' || c == '-' || c == ' ' || c == ',' || (dotted && c == '.'))
        .filter(|p| !p.is_empty())
        .collect();
    let numeric = |s: &str| (s.chars().all(|c| c.is_ascii_digit()) && !s.is_empty()).then(|| s.parse::<i32>().ok()).flatten();
    let (year, month, day) = match parts.as_slice() {
        [a, b, c] => match (numeric(a), numeric(b), numeric(c)) {
            (Some(y), Some(m), Some(d)) if a.len() == 4 => (y, m as u32, d),
            (Some(x), Some(y), Some(z)) => match settings.date_order {
                DateOrder::MonthFirst => (full_year(z, c.len())?, x as u32, y),
                DateOrder::DayFirst => (full_year(z, c.len())?, y as u32, x),
                DateOrder::YearFirst => (full_year(x, a.len())?, y as u32, z),
            },
            (Some(d), None, Some(y)) => (full_year(y, c.len())?, month_named(b)?, d),
            (None, Some(d), Some(y)) => (full_year(y, c.len())?, month_named(a)?, d),
            _ => return None,
        },
        [a, b] => match (numeric(a), numeric(b)) {
            (None, Some(y)) if b.len() == 4 => (y, month_named(a)?, 1),
            _ => return None,
        },
        _ => return None,
    };
    serial(year, month, day)
}

fn full_year(year: i32, digits: usize) -> Option<i32> {
    match digits {
        4 => Some(year),
        // Excel's two-digit window: 00-29 are 2000s, 30-99 are 1900s.
        1 | 2 => Some(if year < 30 { 2000 + year } else { 1900 + year }),
        _ => None,
    }
}

fn serial(year: i32, month: u32, day: i32) -> Option<f64> {
    let date = NaiveDate::from_ymd_opt(year, month, u32::try_from(day).ok()?)?;
    (date.year() >= 1900).then(|| crate::functions::date::excel1900_serial_from_ymd(year, month, day) as f64)
}

/// A number rounded to 15 significant digits, the precision Excel shows, compares and rounds at:
/// `=0.1+0.2=0.3` is TRUE and ROUND(1.005,2) is 1.01.
pub fn round_15(n: f64) -> f64 {
    if n == 0.0 || !n.is_finite() {
        return n;
    }
    format!("{:.14e}", n).parse().unwrap_or(n)
}

/// A number as Excel writes it as text (`&`, CONCATENATE, a text argument): at most 15
/// significant digits, so `=1/3&""` is "0.333333333333333" and `=0.1+0.2&""` is "0.3", in
/// scientific notation from 1E+15 up and below 1E-09. The decimal separator is the regional one.
pub fn number_text(n: f64) -> String {
    let text = number_text_with_point(n);
    match crate::regional::current().decimal_separator {
        '.' => text,
        separator => text.replace('.', &separator.to_string()),
    }
}

fn number_text_with_point(n: f64) -> String {
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
    if !(-9..15).contains(&exponent) {
        let scientific = format!("{:.14e}", rounded);
        let (mantissa, power) = scientific.split_once('e').unwrap_or((&scientific, "0"));
        let power: i32 = power.parse().unwrap_or(0);
        let sign = if power < 0 { '-' } else { '+' };
        return format!("{}E{}{:02}", trim(mantissa.to_string()), sign, power.abs());
    }
    trim(format!("{:.*}", (14 - exponent).max(0) as usize, rounded))
}

#[cfg(test)]
mod tests {
    use super::text_to_number;

    #[test]
    fn reads_numbers_like_excel() {
        for (text, number) in [
            ("1234", 1234.0),
            (" 1,234.5 ", 1234.5),
            ("-1,234", -1234.0),
            ("$1,234", 1234.0),
            ("(5)", -5.0),
            ("12%", 0.12),
            ("1.2E+3", 1200.0),
            (".5", 0.5),
            ("+3", 3.0),
        ] {
            assert_eq!(text_to_number(text), Some(number), "{text}");
        }
        for text in ["", "abc", "1,23", "12,34,567", "1.2.3", "inf", "NaN", "--1", "$", "%"] {
            assert_eq!(text_to_number(text), None, "{text}");
        }
    }

    #[test]
    fn reads_dates_and_times_like_excel() {
        for (text, number) in [
            ("2024-03-15", 45366.0),
            ("3/15/2024", 45366.0),
            ("15-Mar-2024", 45366.0),
            ("Mar 15, 2024", 45366.0),
            ("15 March 2024", 45366.0),
            ("3/15/24", 45366.0),
            ("Mar-2024", 45352.0),
            ("1900-01-01", 1.0),
            ("1900-03-01", 61.0),
            ("18:00", 0.75),
            ("6:30 PM", 18.5 / 24.0),
            ("12:00 AM", 0.0),
            ("2024-03-15 18:00", 45366.75),
            ("25:00", 25.0 / 24.0),
        ] {
            assert_eq!(text_to_number(text), Some(number), "{text}");
        }
        for text in ["2024-02-30", "13/1/2024", "12:60", "13:00 PM", "Marchy 5, 2024"] {
            assert_eq!(text_to_number(text), None, "{text}");
        }
    }

    #[test]
    fn reads_text_in_the_regional_settings() {
        use crate::regional::with_settings;
        with_settings("en-GB", || {
            assert_eq!(text_to_number("13/3/2024"), Some(45364.0));
            assert_eq!(text_to_number("2024-03-13"), Some(45364.0));
            assert_eq!(text_to_number("3/13/2024"), None);
            assert_eq!(text_to_number("1,234.5"), Some(1234.5));
        });
        with_settings("de-DE", || {
            assert_eq!(text_to_number("13.3.2024"), Some(45364.0));
            assert_eq!(text_to_number("1.234,5"), Some(1234.5));
            assert_eq!(text_to_number("1,5"), Some(1.5));
            assert_eq!(text_to_number("1.5"), None);
            assert_eq!(super::number_text(1.5), "1,5");
        });
        with_settings("sv-SE", || {
            assert_eq!(text_to_number("24-03-13"), Some(45364.0));
            assert_eq!(text_to_number("1 234,5"), Some(1234.5));
        });
        with_settings("es-ES", || assert_eq!(text_to_number("13 marzo 2024"), Some(45364.0)));
    }
}
