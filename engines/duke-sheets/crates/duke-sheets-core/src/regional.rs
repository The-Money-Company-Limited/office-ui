//! The regional settings Excel takes from the computer it runs on: whether `5/11/2026` is
//! 5 November or May 11, the language months and days are named in, and the decimal separator
//! (`1,5`). Formulas compute with them, typed entries are read with them, and cells display with
//! them. A file doesn't record them, so the host sets them (the editor, from the browser's
//! language).

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

/// Where the currency symbol goes in an amount.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CurrencyPosition {
    /// $1,234 (United States, United Kingdom).
    #[default]
    Before,
    /// € 1.234 (Netherlands, Switzerland).
    BeforeWithSpace,
    /// 1.234 € (Germany, France, Spain).
    AfterWithSpace,
}

/// Regional settings. The default is en-US.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RegionalSettings {
    pub language: Language,
    pub date_order: DateOrder,
    /// The separator a date written as numbers uses besides `/`, `-` and spaces (`.` in
    /// `5.11.2026`).
    pub date_separator: char,
    pub decimal_separator: char,
    pub thousands_separator: char,
    /// The computer's short date, which Excel shows its built-in short date (14) with:
    /// `m/d/yyyy` in en-US, `dd/mm/yyyy` in en-GB, `yyyy/m/d` in zh-CN.
    pub short_date: &'static str,
    /// The hour code of the computer's short time, `h` (9:30) or `hh` (09:30).
    pub short_hour: &'static str,
    /// The computer's currency symbol, which the built-in currency formats (5-8, 42 and 44)
    /// show and typed amounts may carry: `$`, `£`, `€`.
    pub currency_symbol: &'static str,
    pub currency_position: CurrencyPosition,
}

impl Default for RegionalSettings {
    fn default() -> Self {
        Self {
            language: Language::default(),
            date_order: DateOrder::default(),
            date_separator: '/',
            decimal_separator: '.',
            thousands_separator: ',',
            short_date: "m/d/yyyy",
            short_hour: "h",
            currency_symbol: "$",
            currency_position: CurrencyPosition::Before,
        }
    }
}

impl RegionalSettings {
    /// The settings for a language tag such as `en-GB`, `es-ES`, `sv` or `de_CH`. Languages without
    /// their own month names here use English ones; a tag that doesn't start with a two- or
    /// three-letter language gives en-US.
    pub fn from_tag(tag: &str) -> Self {
        let mut parts = tag.trim().split(['-', '_']);
        let language_code = parts.next().unwrap_or("").to_ascii_lowercase();
        if !(2..=3).contains(&language_code.len()) || !language_code.bytes().all(|b| b.is_ascii_lowercase()) {
            return Self::default();
        }
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
            ("sv" | "ja" | "zh" | "ko" | "hu" | "lt" | "mn", _) | ("en" | "fr", "CA") | ("en", "ZA") => DateOrder::YearFirst,
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
            ("de" | "fi" | "nb" | "no" | "da" | "pl" | "ru" | "tr" | "cs" | "sk" | "uk" | "ro" | "hr" | "sl" | "lv" | "et" | "bg" | "hu", _) => '.',
            ("fr", "CH") | ("it", "CH") => '.',
            ("nl" | "sv", _) | ("en" | "fr", "CA") => '-',
            _ => '/',
        };
        // Windows' own short date where it isn't the date order with padded day and month.
        let (short_date, short_hour) = match (language_code.as_str(), region.as_str()) {
            ("ja", _) => ("yyyy/mm/dd", "h"),
            ("zh", _) => ("yyyy/m/d", "h"),
            ("ko", _) | ("en", "CA") => ("yyyy-mm-dd", "h"),
            _ => match (date_order, date_separator) {
                (DateOrder::MonthFirst, _) => ("m/d/yyyy", "h"),
                (DateOrder::DayFirst, '.') => ("dd.mm.yyyy", "hh"),
                (DateOrder::DayFirst, '-') => ("dd-mm-yyyy", "hh"),
                (DateOrder::DayFirst, _) => ("dd/mm/yyyy", "hh"),
                (DateOrder::YearFirst, '.') => ("yyyy.mm.dd", "hh"),
                (DateOrder::YearFirst, '/') => ("yyyy/mm/dd", "hh"),
                (DateOrder::YearFirst, _) => ("yyyy-mm-dd", "hh"),
            },
        };
        let (currency_symbol, currency_position) = currency(&language_code, &region);
        Self {
            language,
            date_order,
            date_separator,
            decimal_separator,
            thousands_separator,
            short_date,
            short_hour,
            currency_symbol,
            currency_position,
        }
    }

    /// The code Excel shows a built-in currency format with where the currency isn't en-US's `$`
    /// before the amount (whose codes are the built-in table's): 5-8 (currency, with or without
    /// decimals, negatives in red or not) and 42 and 44 (accounting), with the symbol where the
    /// region puts it and a minus sign for negatives. `None` for another format, or for `$`.
    pub fn currency_code(&self, id: u32) -> Option<String> {
        if !matches!(id, 5..=8 | 42 | 44) || (self.currency_symbol, self.currency_position) == ("$", CurrencyPosition::Before) {
            return None;
        }
        let number = if matches!(id, 7 | 8 | 44) { "#,##0.00" } else { "#,##0" };
        if matches!(id, 42 | 44) {
            // The fill (`* `) pads between the symbol and the number, as Excel's do.
            let zero = if id == 44 { "\"-\"??" } else { "\"-\"" };
            let symbol = format!("\"{}\"", self.currency_symbol);
            let filled = |number: &str| match self.currency_position {
                CurrencyPosition::AfterWithSpace => format!("* {number} {symbol}"),
                _ => format!("{symbol}* {number}"),
            };
            return Some(format!("_-{}_-;\\-{}_-;_-{}_-;_-@_-", filled(number), filled(number), filled(zero)));
        }
        let red = if matches!(id, 6 | 8) { "[Red]" } else { "" };
        let amount = self.amount_code(number);
        Some(format!("{amount};{red}\\-{amount}"))
    }

    /// The format a typed amount with the currency symbol gets (`$1,234` → `"$"#,##0`), with
    /// two decimals when they were typed, as Excel gives it.
    pub fn typed_currency_code(&self, decimals: bool) -> String {
        self.amount_code(if decimals { "#,##0.00" } else { "#,##0" })
    }

    /// An amount written out with the currency symbol where the region puts it (`£5.00`,
    /// `5,00 €`), as DOLLAR writes it.
    pub fn currency_text(&self, amount: &str) -> String {
        place_currency(self.currency_symbol, self.currency_position, amount)
    }

    /// A number code with the currency symbol, quoted, where the region puts it.
    fn amount_code(&self, number: &str) -> String {
        place_currency(&format!("\"{}\"", self.currency_symbol), self.currency_position, number)
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
        ssfmt::Locale {
            decimal_separator: self.decimal_separator,
            thousands_separator: self.thousands_separator,
            date_separator: self.date_separator,
            currency_symbol: self.currency_symbol,
            ..named
        }
    }

    /// The short date with the short time, which Excel shows its built-in date with time (22) with.
    pub fn date_time_code(&self) -> String {
        format!("{} {}:mm", self.short_date, self.short_hour)
    }
}

fn place_currency(symbol: &str, position: CurrencyPosition, amount: &str) -> String {
    match position {
        CurrencyPosition::Before => format!("{symbol}{amount}"),
        CurrencyPosition::BeforeWithSpace => format!("{symbol} {amount}"),
        CurrencyPosition::AfterWithSpace => format!("{amount} {symbol}"),
    }
}

/// The currency symbol and where it goes, for a language and region from a tag (`en`, `GB`).
/// Without a region, the language's main one; `$` before the amount where neither is known.
fn currency(language: &str, region: &str) -> (&'static str, CurrencyPosition) {
    use CurrencyPosition::*;
    match (language, region) {
        (_, "GB") => ("£", Before),
        (_, "CH" | "LI") => ("CHF", BeforeWithSpace),
        (_, "IN") => ("₹", Before),
        (_, "JP") | ("ja", "") => ("¥", Before),
        (_, "CN") | ("zh", "") => ("¥", Before),
        (_, "KR") | ("ko", "") => ("₩", Before),
        (_, "BR") => ("R$", BeforeWithSpace),
        (_, "ZA") => ("R", Before),
        (_, "SE") | ("sv", "") => ("kr", AfterWithSpace),
        (_, "NO") | ("nb" | "no", "") => ("kr", BeforeWithSpace),
        (_, "DK") | ("da", "") => ("kr.", AfterWithSpace),
        (_, "PL") | ("pl", "") => ("zł", AfterWithSpace),
        (_, "CZ") | ("cs", "") => ("Kč", AfterWithSpace),
        (_, "RU") | ("ru", "") => ("₽", AfterWithSpace),
        (_, "TR") | ("tr", "") => ("₺", Before),
        ("fr", "CA") => ("$", AfterWithSpace),
        // The euro: before the amount in Ireland, Malta, the Netherlands and Austria, after it in
        // the rest of the euro area.
        ("en", "IE" | "MT") => ("€", Before),
        ("nl", "NL" | "") | ("de", "AT") => ("€", BeforeWithSpace),
        (_, "DE" | "AT" | "FR" | "ES" | "IT" | "PT" | "BE" | "LU" | "FI" | "GR" | "SK" | "SI" | "EE" | "LV" | "LT" | "HR" | "CY")
        | ("de" | "fr" | "es" | "it" | "pt" | "fi" | "el" | "sk" | "sl" | "et" | "lv" | "lt" | "hr", "") => ("€", AfterWithSpace),
        _ => ("$", Before),
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
        assert_eq!(RegionalSettings::from_tag("not a tag"), RegionalSettings::default());
        assert_eq!(RegionalSettings::from_tag("*"), RegionalSettings::default());
        assert_eq!(RegionalSettings::from_tag("ru-RU").thousands_separator, '\u{a0}');
        assert_eq!(RegionalSettings::from_tag("cs-CZ").decimal_separator, ',');
        assert_eq!(RegionalSettings::from_tag("es-ES").format_locale().month_names_full[0], "enero");
        assert_eq!(RegionalSettings::from_tag("zh-Hans-CN").date_order, DateOrder::YearFirst);
    }

    #[test]
    fn knows_the_regional_currency() {
        let currency = |tag: &str| {
            let settings = RegionalSettings::from_tag(tag);
            (settings.currency_symbol, settings.currency_position)
        };
        use CurrencyPosition::*;
        assert_eq!(currency("en-US"), ("$", Before));
        assert_eq!(currency("en"), ("$", Before));
        assert_eq!(currency("en-GB"), ("£", Before));
        assert_eq!(currency("de-DE"), ("€", AfterWithSpace));
        assert_eq!(currency("de"), ("€", AfterWithSpace));
        assert_eq!(currency("fr-FR"), ("€", AfterWithSpace));
        assert_eq!(currency("nl-NL"), ("€", BeforeWithSpace));
        assert_eq!(currency("en-IE"), ("€", Before));
        assert_eq!(currency("de-CH"), ("CHF", BeforeWithSpace));
        assert_eq!(currency("ja-JP"), ("¥", Before));
        assert_eq!(currency("sv-SE"), ("kr", AfterWithSpace));
        assert_eq!(currency("fr-CA"), ("$", AfterWithSpace));
        let gb = RegionalSettings::from_tag("en-GB");
        assert_eq!(gb.currency_code(7).as_deref(), Some("\"£\"#,##0.00;\\-\"£\"#,##0.00"));
        assert_eq!(gb.currency_code(6).as_deref(), Some("\"£\"#,##0;[Red]\\-\"£\"#,##0"));
        assert_eq!(gb.currency_code(14), None);
        assert_eq!(RegionalSettings::default().currency_code(7), None);
        assert_eq!(gb.typed_currency_code(false), "\"£\"#,##0");
        assert_eq!(RegionalSettings::from_tag("de-DE").typed_currency_code(true), "#,##0.00 \"€\"");
    }

    #[test]
    fn writes_the_short_date_in_the_date_order() {
        let code = |tag: &str| (RegionalSettings::from_tag(tag).short_date.to_string(), RegionalSettings::from_tag(tag).date_time_code());
        assert_eq!(code("en-US"), ("m/d/yyyy".into(), "m/d/yyyy h:mm".into()));
        assert_eq!(code("en-GB"), ("dd/mm/yyyy".into(), "dd/mm/yyyy hh:mm".into()));
        assert_eq!(code("de-DE"), ("dd.mm.yyyy".into(), "dd.mm.yyyy hh:mm".into()));
        assert_eq!(code("sv-SE").0, "yyyy-mm-dd");
        for tag in ["en-CA", "fr-CA"] {
            assert_eq!(code(tag).0, "yyyy-mm-dd", "{tag}");
        }
        assert_eq!(code("en-ZA").0, "yyyy/mm/dd");
        assert_eq!(code("ja-JP"), ("yyyy/mm/dd".into(), "yyyy/mm/dd h:mm".into()));
        assert_eq!(code("zh-CN"), ("yyyy/m/d".into(), "yyyy/m/d h:mm".into()));
        assert_eq!(code("ko-KR").0, "yyyy-mm-dd");
        assert_eq!(code("en-CA").1, "yyyy-mm-dd h:mm");
        assert_eq!(code("sv-SE").1, "yyyy-mm-dd hh:mm");
    }
}
