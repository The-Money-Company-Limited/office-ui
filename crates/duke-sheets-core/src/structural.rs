//! Shared coordinate and formula reference transforms for workbook edits.
use crate::{CellAddress, CellRange, Error, Result, MAX_COLS, MAX_ROWS};

#[derive(Clone, Copy, Debug)]
pub(crate) struct AxisEdit {
    pub rows: bool,
    pub index: u32,
    pub count: u32,
    pub delete: bool,
}

impl AxisEdit {
    pub fn limit(self) -> u32 {
        if self.rows {
            MAX_ROWS
        } else {
            u32::from(MAX_COLS)
        }
    }

    pub fn validate(self) -> Result<()> {
        if self.count == 0 || self.index >= self.limit() || self.count > self.limit() - self.index {
            return Err(Error::other("Invalid structural edit range"));
        }
        Ok(())
    }

    pub fn point(self, value: u32) -> Option<u32> {
        if value < self.index {
            return Some(value);
        }
        if self.delete {
            (value >= self.index + self.count).then(|| value - self.count)
        } else {
            value.checked_add(self.count).filter(|v| *v < self.limit())
        }
    }

    pub fn interval(self, start: u32, end: u32) -> Option<(u32, u32)> {
        if start > end {
            return self.interval(end, start).map(|(a, b)| (b, a));
        }
        if !self.delete {
            return Some((self.point(start)?, self.point(end)?));
        }
        let after = self.index + self.count;
        if end < self.index {
            return Some((start, end));
        }
        if start >= after {
            return Some((start - self.count, end - self.count));
        }
        let first = if start < self.index {
            start
        } else {
            self.index
        };
        let last = if end >= after {
            end - self.count
        } else {
            self.index.checked_sub(1)?
        };
        (first <= last).then_some((first, last))
    }

    pub fn cell(self, row: u32, col: u16) -> Option<(u32, u16)> {
        if self.rows {
            Some((self.point(row)?, col))
        } else {
            Some((row, self.point(u32::from(col))? as u16))
        }
    }

    pub fn range(self, mut range: CellRange) -> Option<CellRange> {
        let (start, end) = if self.rows {
            (range.start.row, range.end.row)
        } else {
            (range.start.col.into(), range.end.col.into())
        };
        // Full-column/row validation and formatting ranges retain their grid edge.
        let (start, end) = if !self.delete && end == self.limit() - 1 {
            (self.point(start)?, end)
        } else {
            self.interval(start, end)?
        };
        if self.rows {
            range.start.row = start;
            range.end.row = end;
        } else {
            range.start.col = start as u16;
            range.end.col = end as u16;
        }
        Some(range)
    }

    pub fn anchor(self, row: &mut u32, col: &mut u16) {
        if self.rows {
            *row = self.point(*row).unwrap_or(self.index.min(self.limit() - 1));
        } else {
            *col = self
                .point((*col).into())
                .unwrap_or(self.index.min(self.limit() - 1)) as u16;
        }
    }
}

pub(crate) enum ReferenceEdit<'a> {
    Axis { sheet: &'a str, edit: AxisEdit },
    Rename { old: &'a str, new: &'a str },
    DeleteSheet { sheet: &'a str },
}

#[derive(Clone)]
struct Endpoint {
    row: Option<u32>,
    col: Option<u16>,
    row_absolute: bool,
    col_absolute: bool,
}

impl Endpoint {
    fn render(&self) -> String {
        let mut out = String::new();
        if let Some(col) = self.col {
            if self.col_absolute {
                out.push('$');
            }
            out.push_str(&CellAddress::column_to_letters(col));
        }
        if let Some(row) = self.row {
            if self.row_absolute {
                out.push('$');
            }
            out.push_str(&(row + 1).to_string());
        }
        out
    }
}

fn endpoint(text: &str, start: usize) -> Option<(Endpoint, usize)> {
    let b = text.as_bytes();
    let mut i = start;
    let absolute = b.get(i) == Some(&b'$');
    if absolute {
        i += 1;
    }
    let col_start = i;
    let mut col = 0u32;
    while b.get(i).is_some_and(u8::is_ascii_alphabetic) {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(b[i].to_ascii_uppercase() - b'A') + 1)?;
        i += 1;
    }
    let has_col = i > col_start;
    if col > u32::from(MAX_COLS) {
        return None;
    }
    let row_absolute = if has_col {
        let yes = b.get(i) == Some(&b'$');
        if yes {
            i += 1;
        }
        yes
    } else {
        absolute
    };
    let row_start = i;
    while b.get(i).is_some_and(u8::is_ascii_digit) {
        i += 1;
    }
    let row = if i > row_start {
        let value = text[row_start..i].parse::<u32>().ok()?;
        if value == 0 || value > MAX_ROWS {
            return None;
        }
        Some(value - 1)
    } else {
        None
    };
    if !has_col && row.is_none() {
        return None;
    }
    Some((
        Endpoint {
            row,
            col: has_col.then(|| (col - 1) as u16),
            row_absolute,
            col_absolute: absolute,
        },
        i,
    ))
}

fn name_char(c: char) -> bool {
    c.is_alphanumeric() || matches!(c, '_' | '.' | '\\' | '$')
}

// Keep strings, structured references, and external workbook references opaque.
// Only reference tokens are replaced, preserving all other formula spelling.
pub(crate) fn rewrite_references(text: &str, context: &str, change: &ReferenceEdit<'_>) -> String {
    let mut out = String::with_capacity(text.len());
    let mut i = 0;
    while i < text.len() {
        let c = text[i..].chars().next().unwrap();
        if c == '"' || c == '[' {
            let start = i;
            i += c.len_utf8();
            let mut depth = 1;
            while i < text.len() {
                let ch = text[i..].chars().next().unwrap();
                i += ch.len_utf8();
                if c == '"' && ch == '"' {
                    if text[i..].starts_with('"') {
                        i += 1;
                    } else {
                        break;
                    }
                } else if c == '[' {
                    if ch == '[' {
                        depth += 1;
                    }
                    if ch == ']' {
                        depth -= 1;
                        if depth == 0 {
                            break;
                        }
                    }
                }
            }
            if c == '[' {
                // An external prefix may continue as Sheet!A1, without brackets.
                let tail = &text[i..];
                if let Some(bang) = tail
                    .find('!')
                    .filter(|n| !tail[..*n].contains(['+', '-', '*', '/', ',', '(', ')', ' ']))
                {
                    i += bang + 1;
                    if let Some((_, end)) = endpoint(text, i) {
                        i = end;
                        if text[i..].starts_with(':') {
                            if let Some((_, end)) = endpoint(text, i + 1) {
                                i = end;
                            }
                        }
                    }
                }
            }
            out.push_str(&text[start..i]);
            continue;
        }
        let boundary = i == 0
            || !text[..i]
                .chars()
                .next_back()
                .is_some_and(|ch| name_char(ch) || matches!(ch, '!' | ']' | '\''));
        if !boundary {
            out.push(c);
            i += c.len_utf8();
            continue;
        }
        let start = i;
        let mut pos = i;
        let mut sheet = None;
        let mut prefix_end = i;
        if c == '\'' {
            pos += 1;
            let mut name = String::new();
            while pos < text.len() {
                let ch = text[pos..].chars().next().unwrap();
                pos += ch.len_utf8();
                if ch == '\'' {
                    if text[pos..].starts_with('\'') {
                        name.push('\'');
                        pos += 1;
                    } else {
                        break;
                    }
                } else {
                    name.push(ch);
                }
            }
            if text[pos..].starts_with('!') {
                pos += 1;
                prefix_end = pos;
                sheet = Some(name);
            } else {
                out.push_str(&text[start..pos]);
                i = pos;
                continue;
            }
        } else {
            while pos < text.len()
                && text[pos..]
                    .chars()
                    .next()
                    .is_some_and(|ch| name_char(ch) || ch == ':')
            {
                pos += text[pos..].chars().next().unwrap().len_utf8();
            }
            if text[pos..].starts_with('!') {
                sheet = Some(text[start..pos].to_owned());
                pos += 1;
                prefix_end = pos;
            } else {
                pos = start;
            }
        }
        let external = sheet.as_ref().is_some_and(|s| s.contains('['));
        let mut prefix = text[start..prefix_end].to_owned();
        if !external {
            if let (Some(sheet), ReferenceEdit::Rename { old, new }) = (&sheet, change) {
                let names: Vec<_> = sheet
                    .split(':')
                    .map(|name| {
                        if name.eq_ignore_ascii_case(old) {
                            *new
                        } else {
                            name
                        }
                    })
                    .collect();
                if names.join(":") != *sheet {
                    prefix = format!("'{}'!", names.join(":").replace('\'', "''"));
                }
            }
        }
        let parsed = endpoint(text, pos).and_then(|(a, end)| {
            if text[end..].starts_with(':') {
                let (b, end) = endpoint(text, end + 1)?;
                (a.row.is_some() == b.row.is_some() && a.col.is_some() == b.col.is_some())
                    .then_some((a, Some(b), end))
            } else {
                (a.row.is_some() && a.col.is_some()).then_some((a, None, end))
            }
        });
        if let Some((mut a, mut b, end)) = parsed.filter(|(_, _, end)| {
            !text[*end..]
                .chars()
                .next()
                .is_some_and(|ch| name_char(ch) || matches!(ch, '(' | '[' | '!'))
        }) {
            let target = sheet.as_deref().unwrap_or(context);
            let mut valid = true;
            let mut changed = false;
            if !external {
                match change {
                    ReferenceEdit::Axis { sheet, edit } if target.eq_ignore_ascii_case(sheet) => {
                        let first = if edit.rows {
                            a.row
                        } else {
                            a.col.map(u32::from)
                        };
                        if let Some(first) = first {
                            let last = b.as_ref().and_then(|b| {
                                if edit.rows {
                                    b.row
                                } else {
                                    b.col.map(u32::from)
                                }
                            });
                            let mapped = if let Some(last) = last {
                                edit.interval(first, last)
                            } else {
                                edit.point(first).map(|n| (n, n))
                            };
                            if let Some((first, last)) = mapped {
                                if edit.rows {
                                    a.row = Some(first);
                                    if let Some(b) = &mut b {
                                        b.row = Some(last);
                                    }
                                } else {
                                    a.col = Some(first as u16);
                                    if let Some(b) = &mut b {
                                        b.col = Some(last as u16);
                                    }
                                }
                                changed = true;
                            } else {
                                valid = false;
                            }
                        }
                    }
                    ReferenceEdit::DeleteSheet { sheet } if target.eq_ignore_ascii_case(sheet) => {
                        valid = false
                    }
                    _ => {}
                }
            }
            if valid {
                out.push_str(&prefix);
                if changed {
                    out.push_str(&a.render());
                    if let Some(b) = b {
                        out.push(':');
                        out.push_str(&b.render());
                    }
                } else {
                    out.push_str(&text[pos..end]);
                }
            } else {
                out.push_str("#REF!");
            }
            i = end;
        } else if prefix_end > start {
            out.push_str(&prefix);
            i = prefix_end;
        } else {
            out.push(c);
            i += c.len_utf8();
        }
    }
    out
}
