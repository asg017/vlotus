//! Hyperlink support for vlotus.
//!
//! Three layers:
//!  1. Auto-styling: text cells whose value matches [`looks_like_url`]
//!     render with underline + cyan fg in `ui::draw_grid`.
//!  2. Open: [`open_url`] hands a URL to the OS browser; the `go` keybind
//!     and ctrl-click both route through this.
//!  3. Formula: a `HYPERLINK(url, label)` custom function and matching
//!     `"hyperlink"` custom type, registered locally on every
//!     `Sheet::new` site via [`register`]. The custom value packs both
//!     URL and label into [`CustomValue::data`] using a single
//!     [`SEP`] (ASCII unit-separator) byte; [`split_payload`] is the
//!     canonical decoder. `display()` returns the label so the grid
//!     and `CONCAT` see the user-visible string; the engine never
//!     needs to know about the encoding.
//!
//! All three layers are vlotus-only — `lotus-url` and the engine stay
//! untouched. Other consumers of the engine (lotus-wasm, lotus-pyo3)
//! don't get HYPERLINK unless they wire it up themselves.

use std::sync::Arc;

use lotus_core::{
    BinaryOp, CellValue, CompareOp, CustomFunction, CustomTypeHandler, CustomValue, RegistryError,
    Sheet,
};

/// Type tag for the hyperlink custom type. Stored on the resulting
/// [`CustomValue`] so [`split_payload`] consumers can distinguish a
/// real hyperlink from a string that happens to contain the
/// separator.
pub const TYPE_TAG: &str = "hyperlink";

/// ASCII unit separator (U+001F). Picked because it never appears in
/// real cell data — its only standardised use is as a field separator
/// in legacy data formats. Used to pack `(url, label)` into the single
/// [`CustomValue::data`] string.
pub const SEP: char = '\u{1F}';

/// URL schemes vlotus recognises for auto-styling. Order is irrelevant
/// (longest-match isn't needed since none are prefixes of each other).
const SCHEMES: &[&str] = &[
    "http://",
    "https://",
    "mailto:",
    "ftp://",
    "ftps://",
    "file://",
];

/// True if `s` starts with a recognised URL scheme and has at least one
/// non-whitespace character after the prefix. Whitespace anywhere in
/// `s` (including leading) disqualifies — spreadsheet cells with
/// surrounding whitespace are usually mistakes, not URLs.
pub fn looks_like_url(s: &str) -> bool {
    if s.chars().any(char::is_whitespace) {
        return false;
    }
    for scheme in SCHEMES {
        if let Some(rest) = s.strip_prefix(scheme) {
            return !rest.is_empty();
        }
    }
    false
}

/// Hand `url` to the OS default browser. No-op under `cfg(test)` so the
/// test suite never spawns a real browser; production builds delegate
/// to the `open` crate (`/usr/bin/open` on macOS, `xdg-open` on Linux,
/// `cmd /c start` on Windows).
pub fn open_url(url: &str) -> Result<(), String> {
    if cfg!(test) {
        return Ok(());
    }
    open::that(url).map_err(|e| e.to_string())
}

/// Pack `(url, label)` into the encoded form stored in
/// [`CustomValue::data`]. Both pieces are taken verbatim — callers
/// must reject inputs containing [`SEP`] before getting here.
pub fn encode(url: &str, label: &str) -> String {
    let mut out = String::with_capacity(url.len() + 1 + label.len());
    out.push_str(url);
    out.push(SEP);
    out.push_str(label);
    out
}

/// Split an encoded payload back into `(url, label)`. Returns `None`
/// when no [`SEP`] is present — used by the render and `go` paths to
/// detect "is this string a hyperlink payload?".
pub fn split_payload(s: &str) -> Option<(&str, &str)> {
    let idx = s.find(SEP)?;
    Some((&s[..idx], &s[idx + SEP.len_utf8()..]))
}

/// Custom type handler for `"hyperlink"` values. `data` holds the
/// `url + SEP + label` payload; `display()` returns the label so the
/// grid and engine-level `CONCAT` see the user-visible string.
struct HyperlinkType;

impl CustomTypeHandler for HyperlinkType {
    fn type_tag(&self) -> &str {
        TYPE_TAG
    }

    fn display(&self, v: &CustomValue) -> String {
        match split_payload(&v.data) {
            Some((_url, label)) => label.to_string(),
            // Defensive: an unencoded data field shouldn't happen
            // (HyperlinkFn always encodes), but if it does, fall
            // back to showing the raw string instead of panicking.
            None => v.data.clone(),
        }
    }

    fn edit_repr(&self, v: &CustomValue) -> String {
        // The edit buffer is normally driven by the cell's raw
        // formula, not the computed value. This is a defensive fallback
        // for any UI that asks the engine directly: show the label so
        // the user sees something readable.
        self.display(v)
    }

    // No binary_op / compare overrides — engine defaults compare on
    // `data` (i.e. URL + label together), giving "same hyperlink"
    // semantics for `=A1=B1`. as_number returns None (text-like).
    fn binary_op(
        &self,
        _op: BinaryOp,
        _lhs: &CellValue,
        _rhs: &CellValue,
    ) -> Option<Result<CellValue, String>> {
        None
    }

    fn compare(
        &self,
        _op: CompareOp,
        _lhs: &CellValue,
        _rhs: &CellValue,
    ) -> Option<Result<bool, String>> {
        None
    }

    fn as_number(&self, _v: &CustomValue) -> Option<f64> {
        None
    }
}

/// `HYPERLINK(url)` / `HYPERLINK(url, label)` — produces a hyperlink
/// custom value. Empty labels (and the 1-arg form) fall back to using
/// the URL as its own label, matching Excel and Sheets.
struct HyperlinkFn;

impl CustomFunction for HyperlinkFn {
    fn name(&self) -> &str {
        "HYPERLINK"
    }

    fn call(&self, args: &[CellValue]) -> Result<CellValue, String> {
        let (url_val, label_val) = match args {
            [u] => (u, None),
            [u, l] => (u, Some(l)),
            _ => return Err("HYPERLINK: expected 1 or 2 arguments".into()),
        };
        let url = match url_val {
            CellValue::String(s) => s.as_str(),
            _ => return Err("HYPERLINK: url must be text".into()),
        };
        if url.contains(SEP) {
            return Err("HYPERLINK: url contains an unsupported control character".into());
        }
        let label_owned;
        let label = match label_val {
            None => url,
            Some(CellValue::String(s)) if s.is_empty() => url,
            Some(CellValue::String(s)) => s.as_str(),
            // Numeric labels: stringify so users can write
            // HYPERLINK("https://x.com", 42) and see "42".
            Some(CellValue::Number(n)) => {
                label_owned = n.to_string();
                label_owned.as_str()
            }
            Some(CellValue::Boolean(b)) => {
                label_owned = if *b { "TRUE".into() } else { "FALSE".into() };
                label_owned.as_str()
            }
            _ => return Err("HYPERLINK: label must be text, number, or boolean".into()),
        };
        if label.contains(SEP) {
            return Err("HYPERLINK: label contains an unsupported control character".into());
        }
        Ok(CellValue::Custom(CustomValue {
            type_tag: TYPE_TAG.into(),
            data: encode(url, label),
        }))
    }
}

/// Register HYPERLINK on a freshly-built [`Sheet`]. Idempotent guard:
/// if the type tag is already registered (e.g. a future caller
/// double-registers), errors with [`RegistryError::DuplicateTypeTag`].
/// Call this exactly once per `Sheet::new` site.
pub fn register(sheet: &mut Sheet) -> Result<(), RegistryError> {
    sheet.register_type(Arc::new(HyperlinkType))?;
    sheet.register_function(Arc::new(HyperlinkFn))?;
    Ok(())
}


#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matches_common_schemes() {
        assert!(looks_like_url("https://example.com"));
        assert!(looks_like_url("http://example.com/path?q=1"));
        assert!(looks_like_url("mailto:foo@bar.com"));
        assert!(looks_like_url("file:///tmp/x"));
        assert!(looks_like_url("ftp://ftp.example.com"));
        assert!(looks_like_url("ftps://ftp.example.com"));
    }

    #[test]
    fn rejects_scheme_with_no_rest() {
        assert!(!looks_like_url("https://"));
        assert!(!looks_like_url("mailto:"));
    }

    #[test]
    fn rejects_unknown_schemes() {
        assert!(!looks_like_url("foo:bar"));
        assert!(!looks_like_url("javascript:alert(1)"));
        assert!(!looks_like_url("ssh://host"));
    }

    #[test]
    fn rejects_plain_text() {
        assert!(!looks_like_url(""));
        assert!(!looks_like_url("hello"));
        assert!(!looks_like_url("example.com"));
        assert!(!looks_like_url("user@host"));
    }

    #[test]
    fn rejects_whitespace_anywhere() {
        // Leading: probably a typo; never auto-link it.
        assert!(!looks_like_url(" https://x.com"));
        // Trailing: same.
        assert!(!looks_like_url("https://x.com "));
        // Embedded: definitely not a URL.
        assert!(!looks_like_url("https://x .com"));
    }

    // ── Encoding ────────────────────────────────────────────────────

    #[test]
    fn encode_round_trips_through_split() {
        let payload = encode("https://example.com", "Click here");
        assert_eq!(
            split_payload(&payload),
            Some(("https://example.com", "Click here"))
        );
    }

    #[test]
    fn split_payload_returns_none_when_no_separator() {
        assert_eq!(split_payload("plain string"), None);
        assert_eq!(split_payload(""), None);
        assert_eq!(split_payload("https://example.com"), None);
    }

    #[test]
    fn split_payload_handles_empty_label() {
        let payload = encode("https://x.com", "");
        assert_eq!(split_payload(&payload), Some(("https://x.com", "")));
    }

    // ── HyperlinkType ───────────────────────────────────────────────

    fn cv(data: &str) -> CustomValue {
        CustomValue {
            type_tag: TYPE_TAG.into(),
            data: data.into(),
        }
    }

    #[test]
    fn hyperlink_display_returns_label() {
        let h = HyperlinkType;
        assert_eq!(
            h.display(&cv(&encode("https://x.com", "click"))),
            "click"
        );
    }

    #[test]
    fn hyperlink_display_falls_back_to_data_without_separator() {
        let h = HyperlinkType;
        assert_eq!(h.display(&cv("https://x.com")), "https://x.com");
    }

    #[test]
    fn hyperlink_edit_repr_matches_display() {
        let h = HyperlinkType;
        let v = cv(&encode("https://x.com", "click"));
        assert_eq!(h.edit_repr(&v), h.display(&v));
    }

    // ── HyperlinkFn ─────────────────────────────────────────────────

    fn call(args: &[CellValue]) -> Result<CellValue, String> {
        HyperlinkFn.call(args)
    }

    fn s(v: &str) -> CellValue {
        CellValue::String(v.into())
    }

    #[test]
    fn one_arg_uses_url_as_label() {
        let v = call(&[s("https://x.com")]).unwrap();
        match v {
            CellValue::Custom(cv) => {
                assert_eq!(cv.type_tag, TYPE_TAG);
                assert_eq!(
                    split_payload(&cv.data),
                    Some(("https://x.com", "https://x.com"))
                );
            }
            other => panic!("expected Custom, got {other:?}"),
        }
    }

    #[test]
    fn two_arg_packs_url_and_label() {
        let v = call(&[s("https://x.com"), s("click")]).unwrap();
        match v {
            CellValue::Custom(cv) => {
                assert_eq!(split_payload(&cv.data), Some(("https://x.com", "click")));
            }
            other => panic!("expected Custom, got {other:?}"),
        }
    }

    #[test]
    fn empty_label_falls_back_to_url() {
        let v = call(&[s("https://x.com"), s("")]).unwrap();
        match v {
            CellValue::Custom(cv) => {
                assert_eq!(
                    split_payload(&cv.data),
                    Some(("https://x.com", "https://x.com"))
                );
            }
            other => panic!("expected Custom, got {other:?}"),
        }
    }

    #[test]
    fn numeric_label_stringifies() {
        let v = call(&[s("https://x.com"), CellValue::Number(42.0)]).unwrap();
        match v {
            CellValue::Custom(cv) => {
                assert_eq!(split_payload(&cv.data), Some(("https://x.com", "42")));
            }
            other => panic!("expected Custom, got {other:?}"),
        }
    }

    #[test]
    fn rejects_zero_or_three_args() {
        assert!(call(&[]).is_err());
        assert!(call(&[s("a"), s("b"), s("c")]).is_err());
    }

    #[test]
    fn rejects_non_text_url() {
        assert!(call(&[CellValue::Number(1.0)]).is_err());
        assert!(call(&[CellValue::Boolean(true), s("label")]).is_err());
    }

    #[test]
    fn rejects_separator_in_url_or_label() {
        assert!(call(&[s("url\u{1F}injected")]).is_err());
        assert!(call(&[s("https://x.com"), s("la\u{1F}bel")]).is_err());
    }

    // ── Engine integration ──────────────────────────────────────────

    #[test]
    fn register_makes_hyperlink_callable_in_a_formula() {
        let mut sheet = Sheet::new();
        register(&mut sheet).unwrap();
        sheet
            .set_cells(&[(
                "A1".into(),
                "=HYPERLINK(\"https://x.com\", \"click\")".into(),
            )])
            .unwrap();
        match sheet.get("A1") {
            CellValue::Custom(cv) => {
                assert_eq!(cv.type_tag, TYPE_TAG);
                assert_eq!(split_payload(&cv.data), Some(("https://x.com", "click")));
            }
            other => panic!("expected Custom, got {other:?}"),
        }
    }

    #[test]
    fn concat_of_hyperlink_uses_label() {
        // CONCAT routes Custom values through display() per the trait
        // contract; HYPERLINK's display returns the label, so the user
        // sees "click here" not the encoded form.
        let mut sheet = Sheet::new();
        register(&mut sheet).unwrap();
        sheet
            .set_cells(&[
                (
                    "A1".into(),
                    "=HYPERLINK(\"https://x.com\", \"click\")".into(),
                ),
                ("B1".into(), "=A1 & \" here\"".into()),
            ])
            .unwrap();
        match sheet.get("B1") {
            CellValue::String(s) => assert_eq!(s, "click here"),
            other => panic!("expected String, got {other:?}"),
        }
    }

    #[test]
    fn double_register_errors() {
        let mut sheet = Sheet::new();
        register(&mut sheet).unwrap();
        let err = register(&mut sheet).unwrap_err();
        assert!(matches!(err, RegistryError::DuplicateTypeTag(_)));
    }

}
