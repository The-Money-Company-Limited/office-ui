//! The regional settings formulas compute with, as Excel takes them from the computer it runs
//! on: whether `5/11/2026` is 5 November or May 11, the language `TEXT` names months and days in,
//! and the decimal separator in text such as `"1,5"`. A file doesn't record them, so the host sets
//! them (the editor, from the browser's language). The grid's own number display doesn't follow
//! them yet (COD-1800).

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
    /// The separator a date written as numbers uses besides `/`, `-` and spaces (`.` in
    /// `5.11.2026`).
    pub date_separator: char,
    pub decimal_separator: char,
    pub thousands_separator: char,
}

impl Default for RegionalSettings {
    fn default() -> Self {
        Self {
            language: Language::default(),
            date_order: DateOrder::default(),
            date_separator: '/',
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
        // The region is the first two-letter (or three-digit, as in es-419) subtag after the
        // language, before any extension (en-u-hc-h12 has none).
        let region = parts
            .take_while(|p| p.len() > 1)
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
            // Spanish in the Americas north of the equator writes 1,234.5, and Switzerland 1'234.5.
            ("es", "MX" | "US" | "PR" | "GT" | "DO" | "HN" | "NI" | "SV" | "PA" | "419") => ('.', ','),
            ("de" | "it", "CH" | "LI") => ('.', '\''),
            // 1 234,5
            ("fr" | "sv" | "ru" | "pl" | "fi" | "nb" | "no" | "cs" | "sk" | "uk" | "bg" | "hu" | "lt" | "lv" | "et", _)
            | ("pt", "PT") => (',', '\u{a0}'),
            // 1.234,5
            ("es" | "de" | "it" | "pt" | "nl" | "da" | "tr" | "id" | "ro" | "hr" | "sl" | "ca" | "el", _) => (',', '.'),
            _ => ('.', ','),
        };
        let date_separator = match (language_code.as_str(), region.as_str()) {
            ("de" | "fi" | "nb" | "no" | "da" | "pl" | "ru" | "tr" | "cs" | "sk" | "uk" | "ro" | "hr" | "sl" | "lv" | "et" | "bg", _) => '.',
            ("fr", "CH") | ("it", "CH") => '.',
            ("nl" | "sv", _) => '-',
            _ => '/',
        };
        Self { language, date_order, date_separator, decimal_separator, thousands_separator }
    }

    /// The number-format engine's locale: these separators, and month and day names as Excel
    /// writes them in this language.
    pub fn format_locale(&self) -> ssfmt::Locale {
        let english = ssfmt::Locale::en_us();
        let names = |month_names_full, month_names_short, day_names_full, day_names_short| ssfmt::Locale {
            month_names_full,
            month_names_short,
            day_names_full,
            day_names_short,
            ..ssfmt::Locale::en_us()
        };
        let named = match self.language {
            Language::English => english,
            Language::Spanish => names(
                ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
                ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"],
                ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"],
                ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"],
            ),
            Language::French => names(
                ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"],
                ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."],
                ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"],
                ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."],
            ),
            Language::German => names(
                ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Dezember"],
                ["Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"],
                ["Sonntag", "Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag"],
                ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"],
            ),
            Language::Italian => names(
                ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"],
                ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"],
                ["domenica", "lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato"],
                ["dom", "lun", "mar", "mer", "gio", "ven", "sab"],
            ),
            Language::Portuguese => names(
                ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"],
                ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"],
                ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"],
                ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"],
            ),
            Language::Dutch => names(
                ["januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus", "september", "oktober", "november", "december"],
                ["jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"],
                ["zondag", "maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag"],
                ["zo", "ma", "di", "wo", "do", "vr", "za"],
            ),
            Language::Swedish => names(
                ["januari", "februari", "mars", "april", "maj", "juni", "juli", "augusti", "september", "oktober", "november", "december"],
                ["jan", "feb", "mar", "apr", "maj", "jun", "jul", "aug", "sep", "okt", "nov", "dec"],
                ["söndag", "måndag", "tisdag", "onsdag", "torsdag", "fredag", "lördag"],
                ["sön", "mån", "tis", "ons", "tor", "fre", "lör"],
            ),
        };
        ssfmt::Locale { decimal_separator: self.decimal_separator, thousands_separator: self.thousands_separator, ..named }
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
        let ch = RegionalSettings::from_tag("de-CH");
        assert_eq!((ch.decimal_separator, ch.thousands_separator, ch.date_separator), ('.', '\'', '.'));
        assert_eq!(RegionalSettings::from_tag("en-GB").date_separator, '/');
        assert_eq!(RegionalSettings::from_tag("en-u-hc-h12"), RegionalSettings::default());
        assert_eq!(RegionalSettings::from_tag("ru-RU").thousands_separator, '\u{a0}');
        assert_eq!(RegionalSettings::from_tag("cs-CZ").decimal_separator, ',');
        assert_eq!(RegionalSettings::from_tag("es-ES").format_locale().month_names_full[0], "enero");
        assert_eq!(RegionalSettings::from_tag("zh-Hans-CN").date_order, DateOrder::YearFirst);
    }
}
