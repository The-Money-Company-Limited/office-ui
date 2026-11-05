//! Text read as a number, the way Excel reads it wherever a number is needed (`="12%"*2`,
//! `VALUE("1,234")`, `MONTH("2024-03-15")`, `="6:30 PM"+0`): numbers with thousands separators,
//! a currency sign, a percent sign or parentheses for negatives, dates, times, and dates with
//! times. Dates are 1900-system serials and en-US month-first, the engine's default locale.

use chrono::{Datelike, NaiveDate};

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
    let mut number: f64 = format!("{}.{}", if digits.is_empty() { "0" } else { &digits }, fraction.unwrap_or("0"))
        .parse()
        .ok()?;
    if let Some(exponent) = exponent {
        let exponent: i32 = exponent.strip_prefix('+').unwrap_or(exponent).parse().ok()?;
        number *= 10f64.powi(exponent);
    }
    if percent {
        number /= 100.0;
    }
    Some(if negated { -number } else { sign * number })
}

const MONTHS: [&str; 12] = [
    "january", "february", "march", "april", "may", "june", "july", "august", "september", "october",
    "november", "december",
];

fn month_named(word: &str) -> Option<u32> {
    let word = word.trim_end_matches('.').to_ascii_lowercase();
    if word.len() < 3 {
        return None;
    }
    MONTHS
        .iter()
        .position(|m| m.starts_with(&word) && (word.len() == 3 || *m == word))
        .map(|i| i as u32 + 1)
}

/// A date, a time, or a date followed by a time.
fn date_time(text: &str) -> Option<f64> {
    if let Some(time) = time_of_day(text) {
        return Some(time);
    }
    if let Some(date) = date(text) {
        return Some(date);
    }
    // "2024-03-15 18:00", "3/15/2024 6:30 PM": the time starts at the first space that leaves a
    // valid date before it.
    text.match_indices(' ').find_map(|(i, _)| {
        let day = date(text[..i].trim())?;
        let time = time_of_day(text[i..].trim())?;
        Some(day + time)
    })
}

/// `18:30`, `18:30:15`, `6:30 PM`, `6 PM`, `1:05:00.5 am`: a fraction of a day.
fn time_of_day(text: &str) -> Option<f64> {
    let lower = text.to_ascii_lowercase();
    let (clock, meridiem) = if let Some(rest) = lower.strip_suffix("am").or_else(|| lower.strip_suffix('a')) {
        (rest.trim_end(), Some(false))
    } else if let Some(rest) = lower.strip_suffix("pm").or_else(|| lower.strip_suffix('p')) {
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
/// `Mar 15, 2024`, `March 15 2024`, `Mar-2024` (the 1st).
fn date(text: &str) -> Option<f64> {
    let parts: Vec<&str> = text
        .split(|c: char| c == '/' || c == '-' || c == ' ' || c == ',')
        .filter(|p| !p.is_empty())
        .collect();
    let numeric = |s: &str| (s.chars().all(|c| c.is_ascii_digit()) && !s.is_empty()).then(|| s.parse::<i32>().ok()).flatten();
    let (year, month, day) = match parts.as_slice() {
        [a, b, c] => match (numeric(a), numeric(b), numeric(c)) {
            (Some(y), Some(m), Some(d)) if a.len() == 4 => (y, m as u32, d),
            (Some(m), Some(d), Some(y)) => (full_year(y, c.len())?, m as u32, d),
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
    if date.year() < 1900 {
        return None;
    }
    let days = (date - NaiveDate::from_ymd_opt(1899, 12, 30)?).num_days();
    // Excel counts the 29 February 1900 that never was, so dates before March 1900 are one less.
    Some(if days < 61 { days - 1 } else { days } as f64)
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
}
