use ssfmt::{FormatOptions, NumberFormat};

#[test]
fn test_format_date_ymd() {
    let fmt = NumberFormat::parse("yyyy-mm-dd").unwrap();
    let opts = FormatOptions::default();

    // January 9, 2026 = serial 46031
    assert_eq!(fmt.format(46031.0, &opts), "2026-01-09");
}

#[test]
fn test_format_date_mdy() {
    let fmt = NumberFormat::parse("m/d/yyyy").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(46031.0, &opts), "1/9/2026");
}

#[test]
fn test_format_time() {
    let fmt = NumberFormat::parse("h:mm:ss").unwrap();
    let opts = FormatOptions::default();

    // 0.5 = 12:00:00 (noon)
    assert_eq!(fmt.format(0.5, &opts), "12:00:00");
}

#[test]
fn test_format_time_ampm() {
    let fmt = NumberFormat::parse("h:mm AM/PM").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(0.5, &opts), "12:00 PM");
    assert_eq!(fmt.format(0.25, &opts), "6:00 AM");
}

#[test]
fn test_format_datetime() {
    let fmt = NumberFormat::parse("yyyy-mm-dd h:mm").unwrap();
    let opts = FormatOptions::default();

    // January 9, 2026 at noon
    assert_eq!(fmt.format(46031.5, &opts), "2026-01-09 12:00");
}

#[test]
fn test_format_month_name() {
    let fmt = NumberFormat::parse("mmmm d, yyyy").unwrap();
    let opts = FormatOptions::default();

    assert_eq!(fmt.format(46031.0, &opts), "January 9, 2026");
}

#[test]
fn test_date_skip_char_emits_space_not_literal() {
    let fmt = NumberFormat::parse("m/d/yyyy_)").unwrap();
    let opts = FormatOptions::default();

    // Serial 46022 = 12/31/2025
    // _) means "skip width of ')'", should produce a space, not a literal ')'
    assert_eq!(fmt.format(46022.0, &opts), "12/31/2025 ");
}

#[test]
fn a_slash_in_a_date_format_is_the_locale_date_separator() {
    let opts = FormatOptions { locale: ssfmt::Locale { date_separator: '.', ..ssfmt::Locale::en_us() }, ..Default::default() };
    let show = |code: &str, value: f64| NumberFormat::parse(code).unwrap().format(value, &opts);
    // 46331 is 5 November 2026.
    assert_eq!(show("dd/mm/yyyy", 46331.0), "05.11.2026");
    assert_eq!(show("d/m/yy h:mm", 46331.5), "5.11.26 12:00");
    // Written as the character itself, escaped or quoted, it stays a slash.
    assert_eq!(show(r"dd\/mm\/yyyy", 46331.0), "05/11/2026");
    assert_eq!(show(r#"dd"/"mm"/"yyyy"#, 46331.0), "05/11/2026");
    // AM/PM, and a fraction or a slash in a number format, aren't date separators.
    assert_eq!(show("h:mm AM/PM", 0.75), "6:00 PM");
    assert_eq!(show("# ?/?", 1.5), "1 1/2");
    assert_eq!(NumberFormat::parse("dd/mm/yyyy").unwrap().format(46331.0, &FormatOptions::default()), "05/11/2026");
}
