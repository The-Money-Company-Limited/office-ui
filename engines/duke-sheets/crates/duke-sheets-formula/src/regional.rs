//! The workbook's regional settings while a formula evaluates, and the month and day names of
//! each language. Text coercion runs deep inside values that carry no context, so `evaluate`
//! makes the workbook's settings current on its thread (rayon calculates on several) and puts
//! back the previous ones when it returns.

use std::borrow::Cow;
use std::cell::Cell;

use duke_sheets_core::{Language, RegionalSettings};

use crate::evaluator::EvaluationContext;

thread_local! {
    static CURRENT: Cell<RegionalSettings> = Cell::new(RegionalSettings::default());
}

/// The regional settings of the workbook being calculated on this thread (en-US outside one).
pub fn current() -> RegionalSettings {
    CURRENT.with(Cell::get)
}

/// Makes the context's workbook settings current until the guard drops. Nested evaluations of
/// the same workbook find them current already and do nothing.
pub(crate) fn enter(ctx: &EvaluationContext) -> Option<Restore> {
    let settings = ctx.workbook?.settings().regional;
    let previous = CURRENT.with(|c| c.replace(settings));
    (previous != settings).then_some(Restore(previous))
}

pub(crate) struct Restore(RegionalSettings);

impl Drop for Restore {
    fn drop(&mut self) {
        CURRENT.with(|c| c.set(self.0));
    }
}

/// Month and day names in one language, as Excel writes them.
pub struct Names {
    pub months: [&'static str; 12],
    pub months_short: [&'static str; 12],
    pub days: [&'static str; 7],
    pub days_short: [&'static str; 7],
}

pub fn names(language: Language) -> &'static Names {
    match language {
        Language::English => &ENGLISH,
        Language::Spanish => &Names {
            months: ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
            months_short: ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"],
            days: ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"],
            days_short: ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"],
        },
        Language::French => &Names {
            months: ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"],
            months_short: ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."],
            days: ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"],
            days_short: ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."],
        },
        Language::German => &Names {
            months: ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
            months_short: ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"],
            days: ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"],
            days_short: ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"],
        },
        Language::Italian => &Names {
            months: ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"],
            months_short: ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"],
            days: ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"],
            days_short: ["dom", "lun", "mar", "mer", "gio", "ven", "sab"],
        },
        Language::Portuguese => &Names {
            months: ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"],
            months_short: ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"],
            days: ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"],
            days_short: ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"],
        },
        Language::Dutch => &Names {
            months: ["januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november", "december"],
            months_short: ["jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"],
            days: ["zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"],
            days_short: ["zo", "ma", "di", "wo", "do", "vr", "za"],
        },
        Language::Swedish => &Names {
            months: ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december"],
            months_short: ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec"],
            days: ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"],
            days_short: ["sön", "mån", "tis", "ons", "tor", "fre", "lör"],
        },
    }
}

static ENGLISH: Names = Names {
    months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    months_short: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    days: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
    days_short: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
};

/// The month a name stands for (1-12): the full name or its three-letter start in English, or the
/// full or short name in the current language. A trailing period is optional.
pub fn month_named(word: &str) -> Option<u32> {
    let word = word.trim_end_matches('.').to_lowercase();
    if word.chars().count() < 3 {
        return None;
    }
    let english = ENGLISH.months.iter().position(|m| {
        let m = m.to_ascii_lowercase();
        m == word || (word.len() == 3 && m.starts_with(&word))
    });
    let local = || {
        let names = names(current().language);
        (0..12).find(|&i| {
            names.months[i].to_lowercase() == word || names.months_short[i].trim_end_matches('.').to_lowercase() == word
        })
    };
    english.or_else(local).map(|i| i as u32 + 1)
}

/// The number-format engine's locale for `TEXT` in the current settings.
pub fn format_locale() -> ssfmt::Locale {
    let settings = current();
    let names = names(settings.language);
    ssfmt::Locale {
        decimal_separator: settings.decimal_separator,
        thousands_separator: settings.thousands_separator,
        month_names_short: names.months_short,
        month_names_full: names.months,
        day_names_short: names.days_short,
        day_names_full: names.days,
        ..ssfmt::Locale::en_us()
    }
}

/// A `TEXT` format written for the current settings, in the codes the number-format engine reads.
/// Excel reads a TEXT format in the computer's language: a Swedish one writes the year as `ÅÅÅÅ`,
/// and a German one writes `#.##0,00` for a number but `TT.MM.JJJJ` for a date, where the periods
/// are just periods. English codes still work. Quoted text, escaped characters, `[…]` sections,
/// `General`, and AM/PM markers stay as they are.
pub fn format_code(code: &str) -> Cow<'_, str> {
    let settings = current();
    let swap_separators = settings.decimal_separator == ',';
    if settings.language == Language::English && !swap_separators {
        return Cow::Borrowed(code);
    }
    let letter = |c: char, run: usize| -> Option<char> {
        let lower = c.to_lowercase().next().unwrap_or(c);
        Some(match (settings.language, lower) {
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
    };
    // The format as pieces that are codes (true) or kept as written (false).
    let chars: Vec<char> = code.chars().collect();
    let mut pieces: Vec<(bool, String)> = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        let ahead: String = chars[i..chars.len().min(i + 7)].iter().collect::<String>().to_lowercase();
        let kept = match c {
            '"' => chars[i + 1..].iter().position(|&c| c == '"').map_or(chars.len() - i, |p| p + 2),
            '[' => chars[i..].iter().position(|&c| c == ']').map_or(chars.len() - i, |p| p + 1),
            '\\' | '_' | '*' => 2.min(chars.len() - i),
            _ => ["general", "am/pm", "a/p"].into_iter().find(|w| ahead.starts_with(w)).map_or(0, str::len),
        };
        if kept > 0 {
            pieces.push((false, chars[i..i + kept].iter().collect()));
            i += kept;
            continue;
        }
        let run = chars[i..].iter().take_while(|&&d| d.to_lowercase().eq(c.to_lowercase())).count();
        let text: String = match letter(c, run) {
            Some(english) => std::iter::repeat_n(english, run).collect(),
            None => chars[i..i + run].iter().collect(),
        };
        match pieces.last_mut() {
            Some((true, last)) => last.push_str(&text),
            _ => pieces.push((true, text)),
        }
        i += run;
    }
    let is_date = pieces.iter().any(|(code, text)| *code && text.chars().any(|c| "yYmMdDhHsS".contains(c)));
    let mut out = String::with_capacity(code.len());
    for (is_code, text) in pieces {
        if !is_code || !swap_separators {
            out.push_str(&text);
        } else if is_date {
            // The number-format engine would write these as the regional separators.
            for c in text.chars() {
                if matches!(c, '.' | ',') {
                    out.push('\\');
                }
                out.push(c);
            }
        } else {
            out.extend(text.chars().map(|c| match c {
                ',' => '.',
                '.' if settings.thousands_separator == '.' => ',',
                c => c,
            }));
        }
    }
    Cow::Owned(out)
}

#[cfg(test)]
pub(crate) fn with_settings<T>(tag: &str, f: impl FnOnce() -> T) -> T {
    let previous = CURRENT.with(|c| c.replace(RegionalSettings::from_tag(tag)));
    let _restore = Restore(previous);
    f()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn translates_text_format_codes() {
        with_settings("sv-SE", || assert_eq!(format_code("ÅÅÅÅ-MM-DD tt:mm"), "yyyy-MM-DD hh:mm"));
        with_settings("de-DE", || assert_eq!(format_code("TT.MM.JJJJ"), r"dd\.MM\.yyyy"));
        with_settings("de-DE", || assert_eq!(format_code("#.##0,00"), "#,##0.00"));
        with_settings("es-ES", || assert_eq!(format_code("dd/mm/aaaa hh:mm AM/PM"), "dd/mm/yyyy hh:mm AM/PM"));
        with_settings("es-ES", || assert_eq!(format_code("d mmm, aaaa"), r"d mmm\, yyyy"));
        with_settings("es-ES", || assert_eq!(format_code("\"año\" aa"), "\"año\" yy"));
        with_settings("it-IT", || assert_eq!(format_code("General"), "General"));
        with_settings("fr-FR", || assert_eq!(format_code("jj mmmm aaaa"), "dd mmmm yyyy"));
        with_settings("en-GB", || assert_eq!(format_code("dd/mm/yyyy"), "dd/mm/yyyy"));
        with_settings("fr-FR", || assert_eq!(format_code("0,0 \"h\""), "0.0 \"h\""));
    }

    #[test]
    fn reads_month_names() {
        assert_eq!(month_named("Sep"), Some(9));
        assert_eq!(month_named("enero"), None);
        with_settings("es-ES", || {
            assert_eq!(month_named("enero"), Some(1));
            assert_eq!(month_named("Dic"), Some(12));
            assert_eq!(month_named("March"), Some(3));
        });
        with_settings("fr-FR", || assert_eq!(month_named("févr."), Some(2)));
    }
}
