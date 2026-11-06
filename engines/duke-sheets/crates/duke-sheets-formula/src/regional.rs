//! The workbook's regional settings while a formula evaluates. Text coercion runs deep inside
//! values that carry no context, so `evaluate` makes the workbook's settings current on its
//! thread (rayon calculates on several) and puts back the previous ones when it returns. Doing
//! it there covers every way into a calculation: cells, lambdas and single formulas.

use std::borrow::Cow;
use std::cell::Cell;

use duke_sheets_core::{Language, RegionalSettings};

use crate::evaluator::EvaluationContext;

thread_local! {
    static CURRENT: Cell<RegionalSettings> = Cell::new(RegionalSettings::default());
}

/// The regional settings of the workbook being calculated on this thread (en-US outside one).
pub(crate) fn current() -> RegionalSettings {
    CURRENT.with(Cell::get)
}

/// Makes the context's workbook settings current until the guard drops. A nested evaluation of
/// the same workbook finds them current already and changes nothing.
pub(crate) fn enter(ctx: &EvaluationContext) -> Option<Restore> {
    let settings = ctx.workbook?.settings().regional;
    (current() != settings).then(|| make_current(settings))
}

/// Makes `settings` current until the guard drops.
pub(crate) fn make_current(settings: RegionalSettings) -> Restore {
    Restore(CURRENT.with(|c| c.replace(settings)))
}

pub(crate) struct Restore(RegionalSettings);

impl Drop for Restore {
    fn drop(&mut self) {
        CURRENT.with(|c| c.set(self.0));
    }
}

/// The month a name stands for (1-12): the full English name or its first three letters, or the
/// full or short name in the current language. A trailing period is optional.
pub(crate) fn month_named(word: &str) -> Option<u32> {
    let word = word.trim_end_matches('.');
    if word.chars().count() < 3 {
        return None;
    }
    let english = ssfmt::Locale::en_us();
    let local = current().format_locale();
    let lower = word.to_lowercase();
    let same = |name: &str| name.trim_end_matches('.').to_lowercase() == lower;
    (0..12)
        .find(|&i| {
            let full = english.month_names_full[i];
            full.eq_ignore_ascii_case(word) || (word.len() == 3 && full[..3].eq_ignore_ascii_case(word))
        })
        .or_else(|| (0..12).find(|&i| same(local.month_names_full[i]) || same(local.month_names_short[i])))
        .map(|i| i as u32 + 1)
}

/// The English code letter for a letter of a TEXT format in `language`, given how many times it
/// repeats.
fn english_code(language: Language, c: char, run: usize) -> Option<char> {
    Some(match (language, c.to_lowercase().next()?) {
        (Language::Swedish, 'å') => 'y',
        (Language::Swedish, 't') => 'h',
        (Language::German | Language::Dutch, 'j') => 'y',
        (Language::German, 't') => 'd',
        (Language::Dutch, 'u') => 'h',
        (Language::French, 'j') => 'd',
        (Language::Italian, 'g') => 'd',
        // `a` is the year in these languages, written aa or aaaa; a lone one is A/P.
        (Language::Spanish | Language::French | Language::Italian | Language::Portuguese, 'a') if run >= 2 => 'y',
        _ => return None,
    })
}

/// A piece of a format section: codes to translate, or text the number-format engine reads as
/// it is.
enum Piece<'a> {
    Code(String),
    Kept(&'a str),
    Ready(String),
}

/// A `TEXT` format written for the current settings, in the codes the number-format engine reads.
/// Excel reads a TEXT format in the computer's language: a Swedish one writes the year as `ÅÅÅÅ`,
/// and a German one writes `#.##0,00` for a number but `TT.MM.JJJJ` for a date, where the periods
/// are just periods. Each `;` section is a number or a date on its own. English codes still work.
/// Quoted text, escaped characters, colours and conditions, and AM/PM markers stay as they are,
/// and the language's name for General (`Standard`) is General.
pub(crate) fn format_code(code: &str) -> Cow<'_, str> {
    let settings = current();
    if settings.language == Language::English && us_separators(settings) {
        return Cow::Borrowed(code);
    }
    Cow::Owned(sections(code).into_iter().map(|section| translate(section, settings)).collect::<Vec<_>>().join(";"))
}

/// The format's `;` sections, leaving alone a `;` inside quotes or brackets or after a backslash.
fn sections(code: &str) -> Vec<&str> {
    let (mut sections, mut start) = (Vec::new(), 0);
    let (mut quoted, mut bracketed, mut escaped) = (false, false, false);
    for (i, c) in code.char_indices() {
        match c {
            _ if escaped => escaped = false,
            '\\' if !quoted => escaped = true,
            '"' => quoted = !quoted,
            '[' if !quoted => bracketed = true,
            ']' if !quoted => bracketed = false,
            ';' if !quoted && !bracketed => {
                sections.push(&code[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    sections.push(&code[start..]);
    sections
}

fn us_separators(settings: RegionalSettings) -> bool {
    (settings.decimal_separator, settings.thousands_separator) == ('.', ',')
}

/// The names of General in the languages here, which mean General in a TEXT format.
const GENERAL: [&str; 6] = ["standaard", "standard", "estándar", "generale", "allmänt", "geral"];

/// The length of the text at the start of `rest` that has the same letter throughout, as `[tt]`.
fn bracketed_letters(rest: &str) -> Option<(char, usize)> {
    let inner = &rest[1..rest.find(']')?];
    let first = inner.chars().next()?;
    let lower = first.to_lowercase().next();
    inner.chars().all(|c| c.is_alphabetic() && c.to_lowercase().next() == lower).then(|| (first, inner.chars().count()))
}

/// One section, in English codes and with the separators the number-format engine reads.
fn translate(section: &str, settings: RegionalSettings) -> String {
    let mut pieces = Vec::new();
    let mut rest = section;
    while let Some(c) = rest.chars().next() {
        // Elapsed time, [tt]:mm in Swedish: the letters are codes.
        if let Some((letter, run)) = (c == '[').then(|| bracketed_letters(rest)).flatten() {
            let english = english_code(settings.language, letter, run).unwrap_or(letter);
            let end = rest.find(']').map_or(rest.len(), |p| p + 1);
            pieces.push(Piece::Code(format!("[{}]", std::iter::repeat_n(english, run).collect::<String>())));
            rest = &rest[end..];
            continue;
        }
        let general = GENERAL.into_iter().find(|w| rest.get(..w.len()).is_some_and(|start| start.to_lowercase() == *w));
        if let Some(word) = general {
            pieces.push(Piece::Kept("General"));
            rest = &rest[word.len()..];
            continue;
        }
        // A condition, [>=1,5], with its number written the way the engine reads it.
        if rest.starts_with(['[']) && rest[1..].starts_with(['<', '>', '=']) {
            let end = rest.find(']').map_or(rest.len(), |p| p + 1);
            let condition = rest[..end]
                .chars()
                .filter(|&c| c != settings.thousands_separator || us_separators(settings))
                .map(|c| if c == settings.decimal_separator { '.' } else { c })
                .collect();
            pieces.push(Piece::Ready(condition));
            rest = &rest[end..];
            continue;
        }
        let kept = match c {
            '"' => rest[1..].find('"').map_or(rest.len(), |p| p + 2),
            '[' => rest.find(']').map_or(rest.len(), |p| p + 1),
            '\\' | '_' | '*' => rest.char_indices().nth(2).map_or(rest.len(), |(i, _)| i),
            _ => ["general", "am/pm", "a/p"]
                .into_iter()
                .find(|w| rest.get(..w.len()).is_some_and(|start| start.eq_ignore_ascii_case(w)))
                .map_or(0, str::len),
        };
        if kept > 0 {
            pieces.push(Piece::Kept(&rest[..kept]));
            rest = &rest[kept..];
            continue;
        }
        let lower = c.to_lowercase().next();
        let run = rest.chars().take_while(|d| d.to_lowercase().next() == lower).count();
        let english = english_code(settings.language, c, run);
        let text: String = rest.chars().take(run).map(|d| english.unwrap_or(d)).collect();
        match pieces.last_mut() {
            Some(Piece::Code(last)) => last.push_str(&text),
            _ => pieces.push(Piece::Code(text)),
        }
        rest = &rest[rest.char_indices().nth(run).map_or(rest.len(), |(i, _)| i)..];
    }
    let is_date = pieces.iter().any(|p| matches!(p, Piece::Code(text) if text.contains(|c: char| "yYmMdDhHsS".contains(c))));
    let mut out = String::with_capacity(section.len());
    for piece in pieces {
        match piece {
            Piece::Kept(text) => out.push_str(text),
            Piece::Ready(text) => out.push_str(&text),
            Piece::Code(text) if us_separators(settings) => out.push_str(&text),
            Piece::Code(text) if is_date => {
                // A period or comma in a date format is the character itself, except the regional
                // decimal point of fractional seconds (ss,000), which the engine reads as `.`.
                let chars: Vec<char> = text.chars().collect();
                for (i, &c) in chars.iter().enumerate() {
                    let fraction = i > 0 && matches!(chars[i - 1], 's' | 'S') && chars.get(i + 1) == Some(&'0');
                    out.push(if fraction && c == settings.decimal_separator { '.' } else { c });
                }
            }
            Piece::Code(text) => {
                // Where numbers group with spaces (1 234,5), a space between digit placeholders
                // is the grouping separator.
                let chars: Vec<char> = text.chars().collect();
                let placeholder = |i: Option<usize>| i.and_then(|i| chars.get(i)).is_some_and(|c| matches!(c, '#' | '0' | '?'));
                for (i, &c) in chars.iter().enumerate() {
                    out.push(match c {
                        c if c == settings.decimal_separator => '.',
                        c if c == settings.thousands_separator && !c.is_whitespace() => ',',
                        ' ' | '\u{a0}' | '\u{202f}'
                            if settings.thousands_separator == '\u{a0}' && placeholder(i.checked_sub(1)) && placeholder(Some(i + 1)) =>
                        {
                            ','
                        }
                        c => c,
                    });
                }
            }
        }
    }
    out
}

#[cfg(test)]
pub(crate) fn with_settings<T>(tag: &str, f: impl FnOnce() -> T) -> T {
    let _restore = make_current(RegionalSettings::from_tag(tag));
    f()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn translates_text_format_codes() {
        with_settings("sv-SE", || assert_eq!(format_code("ÅÅÅÅ-MM-DD tt:mm"), "yyyy-MM-DD hh:mm"));
        with_settings("de-DE", || assert_eq!(format_code("TT.MM.JJJJ"), "dd.MM.yyyy"));
        with_settings("de-DE", || assert_eq!(format_code("#.##0,00"), "#,##0.00"));
        with_settings("de-DE", || assert_eq!(format_code("#.##0,00;TT.MM.JJJJ"), "#,##0.00;dd.MM.yyyy"));
        with_settings("de-DE", || assert_eq!(format_code(r#""a;b" 0,0"#), r#""a;b" 0.0"#));
        with_settings("es-ES", || assert_eq!(format_code("dd/mm/aaaa hh:mm AM/PM"), "dd/mm/yyyy hh:mm AM/PM"));
        with_settings("es-ES", || assert_eq!(format_code("d mmm, aaaa"), "d mmm, yyyy"));
        with_settings("es-ES", || assert_eq!(format_code("\"año\" aa"), "\"año\" yy"));
        with_settings("it-IT", || assert_eq!(format_code("General"), "General"));
        with_settings("fr-FR", || assert_eq!(format_code("jj mmmm aaaa"), "dd mmmm yyyy"));
        with_settings("fr-FR", || assert_eq!(format_code("0,0 \"h\""), "0.0 \"h\""));
        with_settings("en-GB", || assert_eq!(format_code("dd/mm/yyyy"), "dd/mm/yyyy"));
        with_settings("de-DE", || assert_eq!(format_code("hh:mm:ss,000"), "hh:mm:ss.000"));
        with_settings("sv-SE", || assert_eq!(format_code("[tt]:mm"), "[hh]:mm"));
        with_settings("sv-SE", || assert_eq!(format_code("[Röd]0"), "[Röd]0"));
        with_settings("de-CH", || assert_eq!(format_code("TTTT, TT.MM.JJJJ"), "dddd, dd.MM.yyyy"));
        with_settings("de-CH", || assert_eq!(format_code("#'##0.00"), "#,##0.00"));
        with_settings("de-DE", || assert_eq!(format_code("Standard"), "General"));
        with_settings("de-DE", || assert_eq!(format_code("[>=1,5]0,00;0"), "[>=1.5]0.00;0"));
        with_settings("fr-FR", || assert_eq!(format_code("# ##0,00"), "#,##0.00"));
        with_settings("fr-FR", || assert_eq!(format_code("0 \"h\""), "0 \"h\""));
    }

    #[test]
    fn reads_month_names() {
        assert_eq!(month_named("Sep"), Some(9));
        assert_eq!(month_named("september"), Some(9));
        assert_eq!(month_named("enero"), None);
        with_settings("es-ES", || {
            assert_eq!(month_named("enero"), Some(1));
            assert_eq!(month_named("Dic"), Some(12));
            assert_eq!(month_named("March"), Some(3));
        });
        with_settings("fr-FR", || assert_eq!(month_named("févr."), Some(2)));
    }
}
