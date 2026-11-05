//! Excel's wildcards, in criteria (COUNTIF, SUMIFS, the D-functions), exact lookups (MATCH,
//! VLOOKUP, XLOOKUP) and SEARCH: `?` is any one character, `*` any run of them, and `~` makes the
//! next character literal (`~*` is a star). Matching ignores case.

#[derive(Clone, Copy, PartialEq)]
enum Token {
    One,
    Run,
    Char(char),
}

fn tokens(pattern: &str) -> Vec<Token> {
    let mut out = Vec::with_capacity(pattern.len());
    let mut chars = pattern.chars().flat_map(char::to_lowercase);
    while let Some(c) = chars.next() {
        out.push(match c {
            '?' => Token::One,
            '*' => Token::Run,
            '~' => Token::Char(chars.next().unwrap_or('~')),
            other => Token::Char(other),
        });
    }
    out
}

/// Whether `pattern` uses any wildcard (or the `~` escape).
pub fn has_wildcards(pattern: &str) -> bool {
    pattern.contains(['*', '?', '~'])
}

/// The length of the text the pattern matches from the start of `text`, if any: the whole of
/// `text` when `whole`, else the shortest prefix.
fn match_len(pattern: &[Token], text: &[char], whole: bool) -> Option<usize> {
    let (mut pi, mut ti) = (0usize, 0usize);
    let mut backtrack: Option<(usize, usize)> = None;
    loop {
        if pi == pattern.len() && (!whole || ti == text.len()) {
            return Some(ti);
        }
        let step = match (pattern.get(pi), text.get(ti)) {
            (Some(Token::Run), _) => {
                backtrack = Some((pi, ti));
                pi += 1;
                continue;
            }
            (Some(Token::One), Some(_)) => true,
            (Some(Token::Char(p)), Some(t)) => *p == *t,
            _ => false,
        };
        if step {
            pi += 1;
            ti += 1;
        } else if let Some((star, at)) = backtrack.filter(|&(_, at)| at < text.len()) {
            // Let the last `*` take one more character and try again from there.
            backtrack = Some((star, at + 1));
            pi = star + 1;
            ti = at + 1;
        } else {
            return None;
        }
    }
}

fn lower_chars(text: &str) -> Vec<char> {
    text.chars().flat_map(char::to_lowercase).collect()
}

/// Whether the whole of `text` matches `pattern`.
pub fn matches(pattern: &str, text: &str) -> bool {
    if !has_wildcards(pattern) {
        return pattern.to_lowercase() == text.to_lowercase();
    }
    match_len(&tokens(pattern), &lower_chars(text), true).is_some()
}

/// Where in `text` (a character index) the pattern first matches, for SEARCH.
pub fn find(pattern: &str, text: &str) -> Option<usize> {
    let pattern = tokens(pattern);
    let text = lower_chars(text);
    (0..=text.len()).find(|&start| match_len(&pattern, &text[start..], false).is_some())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matches_like_excel() {
        assert!(matches("ap*", "Apple"));
        assert!(matches("a?c", "ABC"));
        assert!(!matches("a?c", "abbc"));
        assert!(matches("*x*", "box"));
        assert!(matches("a~*", "a*"));
        assert!(!matches("a~*", "ab"));
        assert!(matches("~?", "?"));
        assert!(matches("*", ""));
        assert!(!matches("?", ""));
        assert!(matches("plain", "PLAIN"));
    }

    #[test]
    fn finds_like_search() {
        assert_eq!(find("a?c", "xabc"), Some(1));
        assert_eq!(find("b*d", "abcd"), Some(1));
        assert_eq!(find("~*", "a*b"), Some(1));
        assert_eq!(find("z", "abc"), None);
        assert_eq!(find("", "abc"), Some(0));
    }
}
