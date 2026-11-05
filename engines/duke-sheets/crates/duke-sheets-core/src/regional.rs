//! The regional settings formulas compute with, as Excel takes them from the computer it runs
//! on: whether `5/11/2026` is 5 November or May 11, the language `TEXT` names months and days in,
//! and the decimal separator in text such as `"1,5"`. A file doesn't record them, so the host sets
//! them (the editor, from the browser's language).

/// The language of month and day names, and of the date codes in a `TEXT` format.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Language {
    #[default]
    English,
    Spanish,
    French,
    German,
    Italian,
    Portuguese,
    Dutch,
    Swedish,
}

/// The order of day, month and year in a date written as numbers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum DateOrder {
    /// 11/5/2026 (United States).
    #[default]
    MonthFirst,
    /// 5/11/2026 (most of the world).
    DayFirst,
    /// 2026-11-05 (Sweden, East Asia).
    YearFirst,
}

/// Regional settings for formulas. The default is en-US.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RegionalSettings {
    pub language: Language,
    pub date_order: DateOrder,
    pub decimal_separator: char,
    pub thousands_separator: char,
}

impl Default for RegionalSettings {
    fn default() -> Self {
        Self {
            language: Language::English,
            date_order: DateOrder::MonthFirst,
            decimal_separator: '.',
            thousands_separator: ',',
        }
    }
}

impl RegionalSettings {
    /// The settings for a language tag such as `en-GB`, `es-ES`, `sv` or `de_CH`. Languages without
    /// their own month names here use English ones; an empty or unreadable tag gives en-US.
    pub fn from_tag(tag: &str) -> Self {
        let mut parts = tag.trim().split(['-', '_']);
        let language_code = parts.next().unwrap_or("").to_ascii_lowercase();
        // The region is the first two-letter (or three-digit, as in es-419) subtag after the language.
        let region = parts
            .find(|p| p.len() == 2 || (p.len() == 3 && p.bytes().all(|b| b.is_ascii_digit())))
            .unwrap_or("")
            .to_ascii_uppercase();
        let language = match language_code.as_str() {
            "es" => Language::Spanish,
            "fr" => Language::French,
            "de" => Language::German,
            "it" => Language::Italian,
            "pt" => Language::Portuguese,
            "nl" => Language::Dutch,
            "sv" => Language::Swedish,
            _ => Language::English,
        };
        let date_order = match (language_code.as_str(), region.as_str()) {
            ("", _) | ("en", "" | "US" | "PH") | ("es", "US") => DateOrder::MonthFirst,
            ("sv" | "ja" | "zh" | "ko" | "hu" | "lt" | "mn", _) => DateOrder::YearFirst,
            _ => DateOrder::DayFirst,
        };
        let (decimal_separator, thousands_separator) = match (language_code.as_str(), region.as_str()) {
            // Spanish in the Americas north of the equator, and Swiss German, write 1,234.5.
            ("es", "MX" | "US" | "PR" | "GT" | "DO" | "HN" | "NI" | "SV" | "PA" | "419") => ('.', ','),
            ("de", "CH" | "LI") => ('.', '\''),
            ("fr" | "sv", _) => (',', '\u{a0}'),
            ("es" | "de" | "it" | "pt" | "nl" | "da" | "nb" | "no" | "fi" | "pl" | "tr" | "ru" | "id", _) => (',', '.'),
            _ => ('.', ','),
        };
        Self { language, date_order, decimal_separator, thousands_separator }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_language_tags() {
        assert_eq!(RegionalSettings::from_tag(""), RegionalSettings::default());
        assert_eq!(RegionalSettings::from_tag("en-US"), RegionalSettings::default());
        assert_eq!(RegionalSettings::from_tag("en"), RegionalSettings::default());
        let gb = RegionalSettings::from_tag("en-GB");
        assert_eq!((gb.language, gb.date_order, gb.decimal_separator), (Language::English, DateOrder::DayFirst, '.'));
        let es = RegionalSettings::from_tag("es-ES");
        assert_eq!((es.language, es.date_order, es.decimal_separator), (Language::Spanish, DateOrder::DayFirst, ','));
        assert_eq!(RegionalSettings::from_tag("es-MX").decimal_separator, '.');
        assert_eq!(RegionalSettings::from_tag("es-419").decimal_separator, '.');
        let sv = RegionalSettings::from_tag("sv_SE");
        assert_eq!((sv.language, sv.date_order, sv.decimal_separator), (Language::Swedish, DateOrder::YearFirst, ','));
        assert_eq!(RegionalSettings::from_tag("de-CH").thousands_separator, '\'');
        assert_eq!(RegionalSettings::from_tag("zh-Hans-CN").date_order, DateOrder::YearFirst);
    }
}
