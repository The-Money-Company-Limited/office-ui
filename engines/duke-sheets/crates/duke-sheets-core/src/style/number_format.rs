//! Number format types

use std::borrow::Cow;

use crate::RegionalSettings;

/// What a number format shows a number as. A value typed into a cell keeps the cell's format
/// when it is of the same kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormatKind {
    General,
    Number,
    Percent,
    Date,
    Time,
    DateTime,
    /// Elapsed time, `[h]:mm:ss`.
    Duration,
    /// `@`: what is typed stays text.
    Text,
}

/// Number format for cell display
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub enum NumberFormat {
    /// General format (default)
    #[default]
    General,

    /// Built-in format by ID
    BuiltIn(u32),

    /// Custom format string
    Custom(String),
}

impl NumberFormat {
    /// General format
    pub const GENERAL: Self = NumberFormat::General;

    // Built-in format IDs
    /// 0 - General
    pub const ID_GENERAL: u32 = 0;
    /// 1 - 0
    pub const ID_NUMBER_INT: u32 = 1;
    /// 2 - 0.00
    pub const ID_NUMBER_DEC2: u32 = 2;
    /// 3 - #,##0
    pub const ID_NUMBER_SEP: u32 = 3;
    /// 4 - #,##0.00
    pub const ID_NUMBER_SEP_DEC2: u32 = 4;
    /// 9 - 0%
    pub const ID_PERCENT_INT: u32 = 9;
    /// 10 - 0.00%
    pub const ID_PERCENT_DEC2: u32 = 10;
    /// 11 - 0.00E+00
    pub const ID_SCIENTIFIC: u32 = 11;
    /// 12 - # ?/?
    pub const ID_FRACTION: u32 = 12;
    /// 13 - # ??/??
    pub const ID_FRACTION2: u32 = 13;
    /// 14 - the short date (m/d/yyyy in en-US)
    pub const ID_DATE_SHORT: u32 = 14;
    /// 15 - d-mmm-yy
    pub const ID_DATE_MEDIUM: u32 = 15;
    /// 16 - d-mmm
    pub const ID_DATE_DAY_MONTH: u32 = 16;
    /// 17 - mmm-yy
    pub const ID_DATE_MONTH_YEAR: u32 = 17;
    /// 18 - h:mm AM/PM
    pub const ID_TIME_AMPM: u32 = 18;
    /// 19 - h:mm:ss AM/PM
    pub const ID_TIME_AMPM_SEC: u32 = 19;
    /// 20 - h:mm
    pub const ID_TIME_24H: u32 = 20;
    /// 21 - h:mm:ss
    pub const ID_TIME_24H_SEC: u32 = 21;
    /// 22 - the short date with the time (m/d/yyyy h:mm in en-US)
    pub const ID_DATETIME: u32 = 22;
    /// 37 - #,##0 ;(#,##0)
    pub const ID_ACCOUNTING_INT: u32 = 37;
    /// 38 - #,##0 ;[Red](#,##0)
    pub const ID_ACCOUNTING_INT_RED: u32 = 38;
    /// 39 - #,##0.00;(#,##0.00)
    pub const ID_ACCOUNTING_DEC2: u32 = 39;
    /// 40 - #,##0.00;[Red](#,##0.00)
    pub const ID_ACCOUNTING_DEC2_RED: u32 = 40;
    /// 46 - [h]:mm:ss
    pub const ID_ELAPSED_HOURS: u32 = 46;
    /// 49 - @
    pub const ID_TEXT: u32 = 49;

    /// Create a number format from a format string
    pub fn from_string<S: Into<String>>(format: S) -> Self {
        NumberFormat::Custom(format.into())
    }

    /// Create a built-in format by ID
    pub fn from_id(id: u32) -> Self {
        NumberFormat::BuiltIn(id)
    }

    /// Integer format (0)
    pub fn integer() -> Self {
        NumberFormat::BuiltIn(Self::ID_NUMBER_INT)
    }

    /// Decimal format (0.00)
    pub fn decimal() -> Self {
        NumberFormat::BuiltIn(Self::ID_NUMBER_DEC2)
    }

    /// Number with thousands separator (#,##0)
    pub fn thousands() -> Self {
        NumberFormat::BuiltIn(Self::ID_NUMBER_SEP)
    }

    /// Number with thousands separator and decimals (#,##0.00)
    pub fn thousands_decimal() -> Self {
        NumberFormat::BuiltIn(Self::ID_NUMBER_SEP_DEC2)
    }

    /// Percentage (0%)
    pub fn percent() -> Self {
        NumberFormat::BuiltIn(Self::ID_PERCENT_INT)
    }

    /// Percentage with decimals (0.00%)
    pub fn percent_decimal() -> Self {
        NumberFormat::BuiltIn(Self::ID_PERCENT_DEC2)
    }

    /// Scientific notation (0.00E+00)
    pub fn scientific() -> Self {
        NumberFormat::BuiltIn(Self::ID_SCIENTIFIC)
    }

    /// Short date (m/d/yyyy in en-US)
    pub fn date_short() -> Self {
        NumberFormat::BuiltIn(Self::ID_DATE_SHORT)
    }

    /// Time with AM/PM (h:mm AM/PM)
    pub fn time_ampm() -> Self {
        NumberFormat::BuiltIn(Self::ID_TIME_AMPM)
    }

    /// Date and time (m/d/yyyy h:mm in en-US)
    pub fn datetime() -> Self {
        NumberFormat::BuiltIn(Self::ID_DATETIME)
    }

    /// Text format (@)
    pub fn text() -> Self {
        NumberFormat::BuiltIn(Self::ID_TEXT)
    }

    /// Get the format string
    pub fn format_string(&self) -> &str {
        match self {
            NumberFormat::General => "General",
            NumberFormat::BuiltIn(id) => Self::builtin_format_string(*id),
            NumberFormat::Custom(s) => s,
        }
    }

    /// The code a cell with this format displays with in `regional`: Excel shows the built-in
    /// short date (14) and date with time (22) in the computer's date order.
    pub fn display_code(&self, regional: &RegionalSettings) -> Cow<'_, str> {
        match self {
            NumberFormat::BuiltIn(Self::ID_DATE_SHORT) => regional.short_date.into(),
            NumberFormat::BuiltIn(Self::ID_DATETIME) => regional.date_time_code().into(),
            _ => self.format_string().into(),
        }
    }

    /// What this format shows a number as.
    pub fn kind(&self) -> FormatKind {
        self.kind_of(self.parsed().as_ref())
    }

    /// The format's code, parsed; `None` for General or a code that doesn't parse.
    pub(crate) fn parsed(&self) -> Option<ssfmt::NumberFormat> {
        let code = self.format_string();
        (!code.eq_ignore_ascii_case("general")).then(|| ssfmt::NumberFormat::parse(code).ok()).flatten()
    }

    /// The kind of this format, given its parsed code.
    pub(crate) fn kind_of(&self, parsed: Option<&ssfmt::NumberFormat>) -> FormatKind {
        let Some(parsed) = parsed else {
            let general = self.format_string().eq_ignore_ascii_case("general");
            return if general { FormatKind::General } else { FormatKind::Number };
        };
        if parsed.is_text_format() {
            FormatKind::Text
        } else if parsed.has_elapsed_time() {
            FormatKind::Duration
        } else if parsed.is_date_format() {
            match (parsed.has_calendar_date(), parsed.has_time_of_day()) {
                (true, true) => FormatKind::DateTime,
                (true, false) => FormatKind::Date,
                _ => FormatKind::Time,
            }
        } else if parsed.is_percentage() {
            FormatKind::Percent
        } else {
            FormatKind::Number
        }
    }

    /// Get built-in format string by ID (14 and 22 as Excel writes them in en-US)
    ///
    /// Reference: ECMA-376 Section 18.8.30 (numFmt), plus the implicit
    /// built-in formats defined by the spec and Microsoft documentation.
    fn builtin_format_string(id: u32) -> &'static str {
        match id {
            0 => "General",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            5 => "$#,##0_);($#,##0)",
            6 => "$#,##0_);[Red]($#,##0)",
            7 => "$#,##0.00_);($#,##0.00)",
            8 => "$#,##0.00_);[Red]($#,##0.00)",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            12 => "# ?/?",
            13 => "# ??/??",
            14 => "m/d/yyyy",
            15 => "d-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "m/d/yyyy h:mm",
            37 => "#,##0_);(#,##0)",
            38 => "#,##0_);[Red](#,##0)",
            39 => "#,##0.00_);(#,##0.00)",
            40 => "#,##0.00_);[Red](#,##0.00)",
            45 => "mm:ss",
            46 => "[h]:mm:ss",
            47 => "mm:ss.0",
            48 => "##0.0E+0",
            49 => "@",
            _ => "General",
        }
    }

    /// Whether this format shows a number as a date, a time or a duration.
    pub fn is_date_format(&self) -> bool {
        matches!(self.kind(), FormatKind::Date | FormatKind::Time | FormatKind::DateTime | FormatKind::Duration)
    }
}
