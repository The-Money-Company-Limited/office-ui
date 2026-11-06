//! Text read as a number, the way Excel reads it wherever a number is needed (`="12%"*2`,
//! `VALUE("1,234")`, `MONTH("2024-03-15")`, `="6:30 PM"+0`): numbers with thousands separators,
//! a currency sign, a percent sign or parentheses for negatives, dates, times, and dates with
//! times. Dates are 1900-system serials. The workbook's regional settings (crate::regional) decide
//! the order of day and month, the month names besides English ones, and the decimal separator.

use std::borrow::Cow;

use chrono::{Datelike, NaiveDate};
use duke_sheets_core::{DateOrder, FormatKind, NumberFormat, WorkbookSettings};

use crate::regional::month_named;

/// How text that reads as a number was written, which decides the format a typed entry gets.
#[derive(Debug, Clone, Copy, PartialEq)]
enum Written {
    Number { percent: bool, currency: bool, decimals: bool },
    Date(DateWritten),
    Time { seconds: bool, twelve_hour: bool },
    DateTime,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum DateWritten {
    /// `3/15/2024`, `2024-03-15`
    Numbers,
    /// `15-Mar-2024`, `Mar 15, 2024`
    NamedMonth,
    /// `Mar-2024`, `May 2026`
    MonthAndYear,
    /// `Feb 22`, `22 Feb`: a day of the current year
    DayAndMonth,
}

/// The number `text` stands for, or `None` when Excel would give #VALUE!.
pub fn text_to_number(text: &str) -> Option<f64> {
    read(text).map(|(number, _)| number)
}

fn read(text: &str) -> Option<(f64, Written)> {
    let text = text.trim();
    if text.is_empty() {
        return None;
    }
    plain_number(text).or_else(|| date_time(text)).filter(|(n, _)| n.is_finite())
}

/// What text typed into a cell with the format `current` stands for, as Excel reads it in the
/// workbook's settings: the number, and the built-in format the cell gets when it should change. Each
/// entry has a kind and a format Excel gives it: a percent `0%` or `0.00%` (9, 10) as decimals
/// were typed; a date the short date (14), or d-mmm-yy (15) or mmm-yy (17) when it names the
/// month; a time h:mm or h:mm:ss (20, 21; 18, 19 with AM/PM); a duration of 24 hours or more
/// [h]:mm:ss (46); a date with a time 22. The cell keeps its format for an entry of the same kind,
/// counting dates, times and dates with times as one kind (as LibreOffice does) and a time in a
/// duration format as a duration, and for a plain number, which in a percent format is a percent
/// from 1 up (`15` is 15%). A duration typed into a date or time cell gets 46, which doesn't wrap. `None` when the text stays text: the cell's format is Text,
/// the text isn't a number, it has a currency sign (no currency format yet), or it is a date
/// before the workbook's first day.
pub fn typed_number(text: &str, settings: &WorkbookSettings, current: &NumberFormat) -> Option<(f64, Option<u32>)> {
    let current = current.kind();
    if current == FormatKind::Text {
        return None;
    }
    let (mut value, written) = {
        let _restore = crate::regional::make_current(settings.regional);
        read(text)?
    };
    let (kind, format) = match written {
        Written::Number { currency: true, .. } => return None,
        Written::Number { percent: true, decimals, .. } => (FormatKind::Percent, if decimals { NumberFormat::ID_PERCENT_DEC2 } else { NumberFormat::ID_PERCENT_INT }),
        Written::Number { .. } => {
            if current == FormatKind::Percent && value.abs() >= 1.0 {
                value /= 100.0;
            }
            return Some((value, None));
        }
        Written::Date(DateWritten::Numbers) => (FormatKind::Date, NumberFormat::ID_DATE_SHORT),
        Written::Date(DateWritten::NamedMonth) => (FormatKind::Date, NumberFormat::ID_DATE_MEDIUM),
        Written::Date(DateWritten::MonthAndYear) => (FormatKind::Date, NumberFormat::ID_DATE_MONTH_YEAR),
        Written::Date(DateWritten::DayAndMonth) => (FormatKind::Date, NumberFormat::ID_DATE_DAY_MONTH),
        Written::Time { .. } if value >= 1.0 => (FormatKind::Duration, NumberFormat::ID_ELAPSED_HOURS),
        Written::Time { seconds, twelve_hour } => (FormatKind::Time, match (twelve_hour, seconds) {
            (true, false) => NumberFormat::ID_TIME_AMPM,
            (true, true) => NumberFormat::ID_TIME_AMPM_SEC,
            (false, false) => NumberFormat::ID_TIME_24H,
            (false, true) => NumberFormat::ID_TIME_24H_SEC,
        }),
        Written::DateTime => (FormatKind::DateTime, NumberFormat::ID_DATETIME),
    };
    if matches!(kind, FormatKind::Date | FormatKind::DateTime) && settings.date_1904 {
        // The text reads as a 1900-system serial; a 1904-system workbook counts days from 1904.
        value -= 1462.0;
        if value < 0.0 {
            return None;
        }
    }
    let date_or_time = |kind| matches!(kind, FormatKind::Date | FormatKind::Time | FormatKind::DateTime);
    let keeps = current == kind
        || (current == FormatKind::Duration && kind == FormatKind::Time)
        || (date_or_time(current) && date_or_time(kind));
    Some((value, (!keeps).then_some(format)))
}

/// `1234`, `-1,234.5`, `1.2E+3`, `$1,234`, `(5)`, `12%`, `+3`.
fn plain_number(text: &str) -> Option<(f64, Written)> {
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
    let (currency, text) = match text.strip_prefix('$') {
        Some(rest) => (true, rest.trim_start()),
        None => (false, text),
    };
    // Text written with the regional separators (1.234,5 or 1'234.5), read as 1,234.5. The other
    // separator in it, where it isn't one of the regional ones, makes it not a number.
    let settings = crate::regional::current();
    let text: Cow<str> = if (settings.decimal_separator, settings.thousands_separator) == ('.', ',') {
        text.into()
    } else {
        let space_grouped = settings.thousands_separator == '\u{a0}';
        text.chars()
            .map(|c| match c {
                c if c == settings.decimal_separator => Some('.'),
                c if c == settings.thousands_separator || (space_grouped && matches!(c, ' ' | '\u{202f}')) => Some(','),
                '.' | ',' => None,
                c => Some(c),
            })
            .collect::<Option<String>>()?
            .into()
    };
    let text = text.as_ref();
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
    let decimals = fraction.is_some_and(|f| !f.is_empty());
    Some((if negated { -number } else { sign * number }, Written::Number { percent, currency, decimals }))
}

/// The date in text that is a date, or a date followed by a time (DATEVALUE).
pub fn date_text(text: &str) -> Option<f64> {
    let text = text.trim();
    date(text).map(|(day, _)| day).or_else(|| split_date_time(text).map(|(day, _)| day))
}

/// The time in text that is a time, or a date followed by a time (TIMEVALUE).
pub fn time_text(text: &str) -> Option<f64> {
    let text = text.trim();
    time_of_day(text).map(|(time, _)| time).or_else(|| split_date_time(text).map(|(_, time)| time))
}

/// "2024-03-15 18:00", "3/15/2024 6:30 PM": the time starts at the first space that leaves a
/// valid date before it.
fn split_date_time(text: &str) -> Option<(f64, f64)> {
    text.match_indices(' ').find_map(|(i, _)| Some((date(text[..i].trim())?.0, time_of_day(text[i..].trim())?.0)))
}

/// A date, a time, or a date followed by a time.
fn date_time(text: &str) -> Option<(f64, Written)> {
    if let Some((time, written)) = time_of_day(text) {
        return Some((time, written));
    }
    if let Some((day, written)) = date(text) {
        return Some((day, Written::Date(written)));
    }
    split_date_time(text).map(|(day, time)| (day + time, Written::DateTime))
}

/// `18:30`, `18:30:15`, `6:30 PM`, `6 PM`, `1:05:00.5 am`: a fraction of a day (or more, `25:30`).
fn time_of_day(text: &str) -> Option<(f64, Written)> {
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
            // Fractional seconds may use the regional decimal separator (00:01:30,5).
            let s = s.replace(crate::regional::current().decimal_separator, ".");
            let (whole, fraction) = s.split_once('.').unwrap_or((&s, ""));
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
    let written = Written::Time { seconds: parts.len() == 3, twelve_hour: meridiem.is_some() };
    Some(((hours as f64 * 3600.0 + minutes as f64 * 60.0 + seconds) / 86400.0, written))
}

/// `2024-03-15`, `2024/3/15`, `3/15/2024`, `3-15-24`, `15-Mar-2024`, `15 March 2024`,
/// `Mar 15, 2024`, `March 15 2024`, `Mar-2024` (the 1st), `1Jan2016` (a month name needs no separators).
/// Three numbers are month first, day first or year first, as the regional settings order them,
/// separated throughout by one of `/`, `-` or the regional date separator (`15.3.2024`), so that
/// `3 1/4` and `1,2,3` aren't dates; a four-digit first number is always the year.
fn date(text: &str) -> Option<(f64, DateWritten)> {
    let settings = crate::regional::current();
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
        .split(|c: char| matches!(c, '/' | '-' | ' ' | ',') || c == settings.date_separator)
        .filter(|p| !p.is_empty())
        .collect();
    let numeric = |s: &str| (s.chars().all(|c| c.is_ascii_digit()) && !s.is_empty()).then(|| s.parse::<i32>().ok()).flatten();
    // Between the numbers, the same separator each time, with spaces after it if any (cs-CZ
    // `5. 11. 2026`), and the regional one may also end the date (hu-HU `2026. 11. 05.`).
    let one_separator = || {
        let text = text.strip_suffix(settings.date_separator).unwrap_or(text);
        let gaps: Vec<&str> = text.split(|c: char| c.is_ascii_digit()).filter(|gap| !gap.is_empty()).collect();
        let Some(separator) = gaps.first().and_then(|gap| gap.chars().next()) else {
            return false;
        };
        gaps.len() == 2
            && (matches!(separator, '/' | '-') || separator == settings.date_separator)
            && gaps.iter().all(|gap| gap.strip_prefix(separator).is_some_and(|rest| rest.chars().all(|c| c == ' ')))
    };
    let (year, month, day, written) = match parts.as_slice() {
        [a, b, c] => match (numeric(a), numeric(b), numeric(c)) {
            (Some(_), Some(_), Some(_)) if !one_separator() => return None,
            (Some(y), Some(m), Some(d)) if a.len() == 4 => (y, m as u32, d, DateWritten::Numbers),
            (Some(x), Some(y), Some(z)) => match settings.date_order {
                DateOrder::MonthFirst => (full_year(z, c.len())?, x as u32, y, DateWritten::Numbers),
                DateOrder::DayFirst => (full_year(z, c.len())?, y as u32, x, DateWritten::Numbers),
                DateOrder::YearFirst => (full_year(x, a.len())?, y as u32, z, DateWritten::Numbers),
            },
            (Some(d), None, Some(y)) => (full_year(y, c.len())?, month_named(b)?, d, DateWritten::NamedMonth),
            (None, Some(d), Some(y)) => (full_year(y, c.len())?, month_named(a)?, d, DateWritten::NamedMonth),
            _ => return None,
        },
        [a, b] => match (numeric(a), numeric(b)) {
            (None, Some(y)) if b.len() == 4 => (y, month_named(a)?, 1, DateWritten::MonthAndYear),
            // `Feb 22` and `22 Feb` are a day of the current year; `Feb 30` and `30 Feb`, with no
            // such day, are February 1930.
            (None, Some(n)) if b.len() <= 2 => month_and_number(month_named(a)?, n, b.len())?,
            (Some(n), None) if a.len() <= 2 => month_and_number(month_named(b)?, n, a.len())?,
            _ => return None,
        },
        _ => return None,
    };
    Some((serial(year, month, day)?, written))
}

fn month_and_number(month: u32, n: i32, digits: usize) -> Option<(i32, u32, i32, DateWritten)> {
    let year = chrono::Local::now().year();
    Some(match NaiveDate::from_ymd_opt(year, month, u32::try_from(n).ok()?) {
        Some(_) => (year, month, n, DateWritten::DayAndMonth),
        None => (full_year(n, digits)?, month, 1, DateWritten::MonthAndYear),
    })
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
    // Excel counts 29 February 1900, which never was (Lotus 1-2-3 thought 1900 a leap year).
    if (year, month, day) == (1900, 2, 29) {
        return Some(60.0);
    }
    let date = NaiveDate::from_ymd_opt(year, month, u32::try_from(day).ok()?)?;
    (date.year() >= 1900).then(|| crate::functions::date::excel1900_serial_from_ymd(year, month, day) as f64)
}

/// A number as Excel writes it as text (`&`, CONCATENATE, a text argument), with the regional
/// decimal separator (duke_sheets_core::number_text).
pub fn number_text(n: f64) -> String {
    duke_sheets_core::number_text(n, crate::regional::current().decimal_separator)
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
        let this_year = chrono::Datelike::year(&chrono::Local::now());
        let feb_22 = crate::functions::date::excel1900_serial_from_ymd(this_year, 2, 22) as f64;
        assert_eq!(text_to_number("Feb 22"), Some(feb_22));
        assert_eq!(text_to_number("22 Feb"), Some(feb_22));
        assert_eq!(text_to_number("Feb 30"), Some(10990.0));
        assert_eq!(text_to_number("30 Feb"), Some(10990.0));
        for text in ["2024-02-30", "13/1/2024", "12:60", "13:00 PM", "Marchy 5, 2024", "Q1 22"] {
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
            assert_eq!(super::number_text(1.5), "1,5");
        });
        with_settings("sv-SE", || {
            assert_eq!(text_to_number("24-03-13"), Some(45364.0));
            assert_eq!(text_to_number("1 234,5"), Some(1234.5));
        });
        with_settings("es-ES", || assert_eq!(text_to_number("13 marzo 2024"), Some(45364.0)));
        with_settings("de-CH", || {
            assert_eq!(text_to_number("1'234.5"), Some(1234.5));
            assert_eq!(text_to_number("13.3.2024"), Some(45364.0));
            assert_eq!(text_to_number("1,5"), None);
        });
    }

    #[test]
    fn reads_typed_entries_in_the_regional_settings() {
        use super::typed_number;
        use duke_sheets_core::{NumberFormat, RegionalSettings, WorkbookSettings};
        let typed = |tag: &str, text: &str| {
            let settings = WorkbookSettings { regional: RegionalSettings::from_tag(tag), ..Default::default() };
            typed_number(text, &settings, &NumberFormat::General)
        };
        let with = |value: f64, id: u32| Some((value, Some(id)));
        // 46331 is 5 November 2026, 46153 is 11 May 2026.
        assert_eq!(typed("en-GB", "5/11/2026"), with(46331.0, 14));
        assert_eq!(typed("en-US", "5/11/2026"), with(46153.0, 14));
        assert_eq!(typed("de-DE", "5.11.2026"), with(46331.0, 14));
        assert_eq!(typed("de-DE", "5.11.2026."), with(46331.0, 14));
        assert_eq!(typed("cs-CZ", "5. 11. 2026"), with(46331.0, 14));
        assert_eq!(typed("hu-HU", "2026. 11. 05."), with(46331.0, 14));
        assert_eq!(typed("en-GB", "5/11/2026 13:30"), with(46331.5625, 22));
        assert_eq!(typed("en-US", "15-Mar-2024"), with(45366.0, 15));
        assert_eq!(typed("en-US", "15 March 2024"), with(45366.0, 15));
        assert_eq!(typed("en-US", "Mar-2024"), with(45352.0, 17));
        // A day and month is this year's, shown d-mmm (16).
        assert_eq!(typed("en-US", "Feb 22").and_then(|(_, format)| format), Some(16));
        assert_eq!(typed("en-GB", "22 Feb").and_then(|(_, format)| format), Some(16));
        assert_eq!(typed("en-US", "Feb 30").and_then(|(_, format)| format), Some(17));
        assert_eq!(typed("en-US", "13:30"), with(0.5625, 20));
        assert_eq!(typed("en-US", "13:30:00"), with(0.5625, 21));
        assert_eq!(typed("en-US", "1:30 PM"), with(0.5625, 18));
        assert_eq!(typed("en-US", "1:30:00 pm"), with(0.5625, 19));
        assert_eq!(typed("en-US", "25:30"), with(25.5 / 24.0, 46));
        assert_eq!(typed("en-US", "2/29/1900"), with(60.0, 14));
        assert_eq!(typed("en-US", "12%"), with(0.12, 9));
        assert_eq!(typed("en-US", "12.%"), with(0.12, 9));
        assert_eq!(typed("en-US", "(12.5%)"), with(-0.125, 10));
        assert_eq!(typed("de-DE", "12,5%"), with(0.125, 10));
        assert_eq!(typed("de-DE", "1,5"), Some((1.5, None)));
        assert_eq!(typed("de-DE", "1.234,5"), Some((1234.5, None)));
        assert_eq!(typed("en-US", " 1,234.5 "), Some((1234.5, None)));
        assert_eq!(typed("en-US", "-2e3"), Some((-2000.0, None)));
        assert_eq!(typed("en-US", "(5)"), Some((-5.0, None)));
        // Text: lists and fractions aren't dates, and there's no currency format to give $.
        for text in ["1,5", "text", "$1,234", "3 1/4", "5 3/8", "1,2,3", "10, 20, 30", "5 10 15", "5/11-2026"] {
            assert_eq!(typed("en-US", text), None, "{text}");
        }
        // The settings are current only while the text is read.
        assert_eq!(text_to_number("1,5"), None);
    }

    #[test]
    fn typed_entries_keep_a_format_of_the_same_kind() {
        use super::typed_number;
        use duke_sheets_core::{NumberFormat, WorkbookSettings};
        let settings = WorkbookSettings::default();
        let format = |text: &str, current: NumberFormat| typed_number(text, &settings, &current).map(|(_, format)| format);
        let custom = |code: &str| NumberFormat::Custom(code.into());
        // Same kind: the cell keeps its format, the toolbar's Date (14) and Time (21) included.
        assert_eq!(format("Mar 15, 2024", NumberFormat::BuiltIn(14)), Some(None));
        assert_eq!(format("9:30", NumberFormat::BuiltIn(21)), Some(None));
        assert_eq!(format("5/11/2026", custom("dd mmm yyyy")), Some(None));
        assert_eq!(format("9:30", NumberFormat::BuiltIn(46)), Some(None));
        assert_eq!(format("12.5%", NumberFormat::BuiltIn(9)), Some(None));
        assert_eq!(format("42", NumberFormat::BuiltIn(14)), Some(None));
        // Dates, times and dates with times are one kind, as edit text needs: an h:mm cell
        // holding a date edits as one.
        assert_eq!(format("13:30", custom("dd mmm yyyy")), Some(None));
        assert_eq!(format("5/11/2026 9:00", NumberFormat::BuiltIn(14)), Some(None));
        assert_eq!(format("11/5/2026 2:30:00 PM", NumberFormat::BuiltIn(20)), Some(None));
        assert_eq!(format("2026-11-05", custom("yyyy-mm-dd hh:mm")), Some(None));
        // A date cell below day 1 edits as a time, and keeps its format typed back.
        assert_eq!(format("12:00:00", NumberFormat::BuiltIn(14)), Some(None));
        // Another kind: the entry's format. A duration doesn't fit a time format, which wraps.
        assert_eq!(format("5/11/2026", custom("0.00")), Some(Some(14)));
        assert_eq!(format("25:30", NumberFormat::BuiltIn(20)), Some(Some(46)));
        assert_eq!(format("5/11/2026", NumberFormat::BuiltIn(46)), Some(Some(14)));
        // Text stays text; in a percent format a number from 1 up is a percent.
        assert_eq!(typed_number("5/11/2026", &settings, &NumberFormat::BuiltIn(49)), None);
        assert_eq!(typed_number("15", &settings, &NumberFormat::BuiltIn(9)), Some((0.15, None)));
        assert_eq!(typed_number("0.5", &settings, &NumberFormat::BuiltIn(9)), Some((0.5, None)));
    }

    #[test]
    fn typed_dates_count_from_the_workbooks_first_day() {
        use super::typed_number;
        use duke_sheets_core::{NumberFormat, WorkbookSettings};
        let settings = WorkbookSettings { date_1904: true, ..Default::default() };
        let value = |text: &str| typed_number(text, &settings, &NumberFormat::General).map(|(value, _)| value);
        assert_eq!(value("11/5/2026"), Some(46331.0 - 1462.0));
        assert_eq!(value("13:30"), Some(0.5625));
        assert_eq!(value("1/1/1902"), None);
    }
}
