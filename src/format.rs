//! Cell display formatting — composable axes (number / style / align /
//! colors).
//!
//! `CellFormat` is a struct that bundles every format dimension a cell
//! can carry. Each axis is independent: a cell can be USD-formatted +
//! bold + center-aligned + red text, all at once. Hand-rolled JSON
//! parser/emitter persists the struct in the existing
//! `datasette_sheets_cell.format_json` column.
//!
//! Axes are added incrementally:
//! - Number format (Usd, eventually Percent) — render layer
//! - Style flags (bold/italic/underline/strike) — render modifiers (T2)
//! - Alignment (Left/Center/Right + auto) — `align_text` override (T3)
//! - Colors (fg/bg) — render style fg/bg (T5)

const MAX_DECIMALS: u8 = 10;

/// Composable cell format. Defaults to no formatting on any axis.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CellFormat {
    pub number: Option<NumberFormat>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub align: Option<Align>,
    pub fg: Option<Color>,
    pub bg: Option<Color>,
    /// Per-cell strftime override applied when the cell's computed
    /// value is a `lotus-datetime` type (`jdate`, `jtime`, `jdatetime`,
    /// `jzoned`). Ignored for any other value type.
    pub date: Option<String>,
}

/// The numeric-format axis. `None` on `CellFormat::number` means the
/// cell renders its raw computed value verbatim.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NumberFormat {
    Usd { decimals: u8 },
    /// `0.045` with `decimals=1` renders as `4.5%`. The raw value is
    /// in fractional form (gsheets convention): `1.0` ↔ `100%`.
    Percent { decimals: u8 },
}

/// Explicit alignment override. `None` on `CellFormat::align` keeps
/// the existing classify-by-display-kind default (number = right,
/// boolean = center, text/error = left).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Align {
    Left,
    Center,
    Right,
}

/// 24-bit RGB color, packed as three bytes. Maps directly to ratatui
/// `Color::Rgb(r, g, b)` at render time.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Color {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

// API surface lands incrementally — T2-T6 callers will materialize
// over the next tickets in this epic. Keep the allow until those land.
#[allow(dead_code)]
impl Color {
    pub const fn rgb(r: u8, g: u8, b: u8) -> Self {
        Color { r, g, b }
    }
}

/// Parse the JSON payload stored in `format_json`. Accepts the literal
/// string `"null"` as the clear-sentinel (matches the COALESCE-aware
/// upsert path in `set_cells_and_recalculate`).
pub fn parse_format_json(s: &str) -> Option<CellFormat> {
    let trimmed = s.trim();
    if trimmed == "null" || trimmed.is_empty() {
        return None;
    }
    let mut fmt = CellFormat::default();

    // Number format: `"n":{"k":"usd","d":2}`
    if let Some(inner) = nested_obj(trimmed, "n") {
        let kind = json_str_field(&inner, "k")?;
        let decimals = json_num_field(&inner, "d")
            .and_then(|n| u8::try_from(n).ok())
            .unwrap_or(2)
            .min(MAX_DECIMALS);
        fmt.number = match kind.as_str() {
            "usd" => Some(NumberFormat::Usd { decimals }),
            "percent" => Some(NumberFormat::Percent { decimals }),
            _ => None,
        };
    }

    fmt.bold = json_bool_field(trimmed, "b").unwrap_or(false);
    fmt.italic = json_bool_field(trimmed, "i").unwrap_or(false);
    fmt.underline = json_bool_field(trimmed, "u").unwrap_or(false);
    fmt.strike = json_bool_field(trimmed, "s").unwrap_or(false);

    if let Some(a) = json_str_field(trimmed, "a") {
        fmt.align = match a.as_str() {
            "left" => Some(Align::Left),
            "center" => Some(Align::Center),
            "right" => Some(Align::Right),
            _ => None,
        };
    }

    fmt.fg = json_str_field(trimmed, "fg").and_then(|h| parse_hex_color(&h));
    fmt.bg = json_str_field(trimmed, "bg").and_then(|h| parse_hex_color(&h));
    fmt.date = json_str_field(trimmed, "df").filter(|s| !s.is_empty());

    Some(fmt)
}

/// Emit the canonical JSON payload. Stable key order so snapshot tests
/// don't churn. Unset fields are omitted to keep payloads compact.
pub fn to_format_json(fmt: &CellFormat) -> String {
    let mut out = String::from("{");
    let mut first = true;
    let mut sep = |s: &mut String| {
        if first {
            first = false;
        } else {
            s.push(',');
        }
    };

    if let Some(num) = fmt.number {
        sep(&mut out);
        let (kind, decimals) = match num {
            NumberFormat::Usd { decimals } => ("usd", decimals),
            NumberFormat::Percent { decimals } => ("percent", decimals),
        };
        out.push_str(&format!(r#""n":{{"k":"{kind}","d":{decimals}}}"#));
    }
    if fmt.bold {
        sep(&mut out);
        out.push_str(r#""b":true"#);
    }
    if fmt.italic {
        sep(&mut out);
        out.push_str(r#""i":true"#);
    }
    if fmt.underline {
        sep(&mut out);
        out.push_str(r#""u":true"#);
    }
    if fmt.strike {
        sep(&mut out);
        out.push_str(r#""s":true"#);
    }
    if let Some(a) = fmt.align {
        sep(&mut out);
        out.push_str(&format!(r#""a":"{}""#, align_str(a)));
    }
    if let Some(c) = fmt.fg {
        sep(&mut out);
        out.push_str(&format!(r#""fg":"{}""#, hex_color(c)));
    }
    if let Some(c) = fmt.bg {
        sep(&mut out);
        out.push_str(&format!(r#""bg":"{}""#, hex_color(c)));
    }
    if let Some(pat) = &fmt.date {
        sep(&mut out);
        out.push_str(&format!(r#""df":{}"#, json_string_lit(pat)));
    }
    out.push('}');
    out
}

/// Quote `s` as a JSON string literal: wraps in `"…"`, escapes `\\`,
/// `"`, and control chars. The strftime patterns we accept are ASCII
/// `%`-codes and literal text, so the bulk of inputs need no escaping —
/// but `\n` / `\t` / embedded quotes are tolerated rather than
/// silently dropped.
fn json_string_lit(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for c in s.chars() {
        match c {
            '\\' => out.push_str(r"\\"),
            '"' => out.push_str(r#"\""#),
            '\n' => out.push_str(r"\n"),
            '\t' => out.push_str(r"\t"),
            '\r' => out.push_str(r"\r"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

fn align_str(a: Align) -> &'static str {
    match a {
        Align::Left => "left",
        Align::Center => "center",
        Align::Right => "right",
    }
}

fn hex_color(c: Color) -> String {
    format!("#{:02x}{:02x}{:02x}", c.r, c.g, c.b)
}

/// Auto-detect a format hint at edit-commit time. Returns
/// `(numeric_raw_string, format)` if the input matches a known shape;
/// `None` otherwise. Recognises:
/// - `$`-prefixed currency (`$1.25`, `-$1,234.56` → USD)
/// - `%`-suffixed percent (`4.5%`, `-100%` → Percent, raw stored as
///   the fractional value: `4.5%` → `0.045`)
///
/// Rejects: bad grouping, trailing `.`, bare `$` / `%`, `$abc`, etc.
pub fn try_parse_typed_input(input: &str) -> Option<(String, CellFormat)> {
    let s = input.trim();
    if let Some(rest) = s.strip_suffix('%') {
        return parse_percent_suffix(rest);
    }
    parse_currency_prefix(s)
}

fn parse_currency_prefix(s: &str) -> Option<(String, CellFormat)> {
    let (sign, after_sign) = if let Some(rest) = s.strip_prefix('-') {
        ("-", rest)
    } else {
        ("", s)
    };
    let after_dollar = after_sign.strip_prefix('$')?;
    if after_dollar.is_empty() {
        return None;
    }

    let (int_part, frac_part) = match after_dollar.split_once('.') {
        Some((i, f)) if !f.is_empty() => (i, Some(f)),
        Some(_) => return None, // trailing dot
        None => (after_dollar, None),
    };

    let int_digits = parse_int_with_optional_grouping(int_part)?;
    if let Some(f) = frac_part {
        if !f.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
    }

    let decimals = frac_part.map(|f| f.len().min(MAX_DECIMALS as usize) as u8).unwrap_or(2);
    let numeric = match frac_part {
        Some(f) => format!("{sign}{int_digits}.{f}"),
        None => format!("{sign}{int_digits}"),
    };
    let fmt = CellFormat {
        number: Some(NumberFormat::Usd { decimals }),
        ..CellFormat::default()
    };
    Some((numeric, fmt))
}

fn parse_percent_suffix(rest: &str) -> Option<(String, CellFormat)> {
    let inner = rest.trim_end();
    if inner.is_empty() {
        return None;
    }
    let (sign, abs_str) = if let Some(r) = inner.strip_prefix('-') {
        ("-", r)
    } else {
        ("", inner)
    };
    if abs_str.is_empty() {
        return None;
    }
    let (int_part, frac_part) = match abs_str.split_once('.') {
        Some((i, f)) if !f.is_empty() => (i, Some(f)),
        Some(_) => return None,
        None => (abs_str, None),
    };
    if !int_part.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    if let Some(f) = frac_part {
        if !f.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
    }
    // Parse the typed value as f64, divide by 100 to get the raw
    // fractional form. Decimals = digits-after-`.` + 2 (since the
    // stored value gains two decimals when divided by 100).
    let typed = match frac_part {
        Some(f) => format!("{int_part}.{f}"),
        None => int_part.to_string(),
    };
    let typed_f: f64 = typed.parse().ok()?;
    let raw_value = typed_f / 100.0;
    let display_decimals = frac_part
        .map(|f| f.len().min(MAX_DECIMALS as usize) as u8)
        .unwrap_or(0);
    // Persist enough precision to round-trip the user's input. `4.5%`
    // → raw 0.045 → string "0.045"; render at display_decimals=1 → "4.5%".
    let raw_string = format_finite(raw_value, display_decimals as usize + 2);
    let raw_with_sign = if sign == "-" && raw_value != 0.0 {
        if raw_string.starts_with('-') {
            raw_string
        } else {
            format!("-{raw_string}")
        }
    } else {
        raw_string
    };
    let fmt = CellFormat {
        number: Some(NumberFormat::Percent { decimals: display_decimals }),
        ..CellFormat::default()
    };
    Some((raw_with_sign, fmt))
}

fn format_finite(n: f64, decimals: usize) -> String {
    let s = format!("{n:.decimals$}");
    // Strip trailing zeros + dot so "0.045" doesn't persist as "0.0450"
    // for a decimals=4 emit. Keep at least one digit after the dot
    // unless we were emitting an integer.
    if !s.contains('.') {
        return s;
    }
    let trimmed = s.trim_end_matches('0').trim_end_matches('.');
    if trimmed.is_empty() || trimmed == "-" {
        "0".to_string()
    } else {
        trimmed.to_string()
    }
}

/// Format a numeric `value_text` for display under the cell's number
/// format. Non-numeric inputs (`#DIV/0!`, strings) and cells without a
/// number format pass through verbatim — style/align/color don't
/// transform the text content.
pub fn render(value_text: &str, fmt: &CellFormat) -> String {
    let Some(num) = fmt.number else {
        return value_text.to_string();
    };
    let Ok(n) = value_text.trim().parse::<f64>() else {
        return value_text.to_string();
    };
    if !n.is_finite() {
        return value_text.to_string();
    }
    match num {
        NumberFormat::Usd { decimals } => render_usd(n, decimals),
        NumberFormat::Percent { decimals } => render_percent(n, decimals),
    }
}

/// Adjust the decimal count of a number format. Clamps to `[0, 10]`.
/// Used by `:fmt+` / `:fmt-` and `g.` / `g,`.
pub fn bump_number_decimals(fmt: &NumberFormat, delta: i8) -> NumberFormat {
    let clamp = |d: u8| {
        ((d as i16) + (delta as i16)).clamp(0, MAX_DECIMALS as i16) as u8
    };
    match fmt {
        NumberFormat::Usd { decimals } => NumberFormat::Usd { decimals: clamp(*decimals) },
        NumberFormat::Percent { decimals } => NumberFormat::Percent { decimals: clamp(*decimals) },
    }
}

/// Returns true iff `s` looks like the rendered output of a number-like
/// format. Used by `classify_display` to keep formatted numerics
/// right-aligned: `$1.25`, `-$3.00`, `4.5%`, `-100%`.
pub fn looks_numeric(s: &str) -> bool {
    let bytes = s.as_bytes();
    if matches!(bytes, [b'$', d, ..] if d.is_ascii_digit())
        || matches!(bytes, [b'-', b'$', d, ..] if d.is_ascii_digit())
    {
        return true;
    }
    if let Some(rest) = s.strip_suffix('%') {
        let inner = rest.strip_prefix('-').unwrap_or(rest);
        if inner.is_empty() {
            return false;
        }
        return inner.parse::<f64>().is_ok();
    }
    false
}

fn render_percent(n: f64, decimals: u8) -> String {
    let scaled = n * 100.0;
    let abs = scaled.abs();
    let formatted = format!("{abs:.*}", decimals as usize);
    let sign = if scaled.is_sign_negative() && scaled != 0.0 {
        "-"
    } else {
        ""
    };
    format!("{sign}{formatted}%")
}

fn render_usd(n: f64, decimals: u8) -> String {
    let abs = n.abs();
    let formatted = format!("{abs:.*}", decimals as usize);
    let (int_part, frac_part) = match formatted.split_once('.') {
        Some((i, f)) => (i, Some(f)),
        None => (formatted.as_str(), None),
    };
    let mut grouped = String::with_capacity(int_part.len() + int_part.len() / 3);
    let bytes = int_part.as_bytes();
    for (i, b) in bytes.iter().enumerate() {
        let from_right = bytes.len() - i;
        if i > 0 && from_right % 3 == 0 {
            grouped.push(',');
        }
        grouped.push(*b as char);
    }
    let sign = if n.is_sign_negative() && n != 0.0 { "-" } else { "" };
    match frac_part {
        Some(f) => format!("{sign}${grouped}.{f}"),
        None => format!("{sign}${grouped}"),
    }
}

/// Parse a thousands-grouped integer like `1,234,567` or a plain `1234567`.
/// Rejects bad groupings (`1,23`, `,123`, `12,3456`).
fn parse_int_with_optional_grouping(s: &str) -> Option<String> {
    if s.is_empty() {
        return None;
    }
    if !s.contains(',') {
        if !s.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
        return Some(s.to_string());
    }
    let groups: Vec<&str> = s.split(',').collect();
    let first = groups.first()?;
    if first.is_empty() || first.len() > 3 || !first.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    for g in &groups[1..] {
        if g.len() != 3 || !g.bytes().all(|b| b.is_ascii_digit()) {
            return None;
        }
    }
    Some(groups.concat())
}

/// Catppuccin Mocha presets with natural-language aliases. Lookup is
/// case-insensitive. Mirror of `theme::mocha::*` so the format module
/// stays self-contained (theme.rs is a UI concern; the format axis
/// values are persisted JSON, so they live here).
pub const COLOR_PRESETS: &[(&str, Color)] = &[
    // Reds / pinks
    ("red", Color::rgb(0xf3, 0x8b, 0xa8)),
    ("maroon", Color::rgb(0xeb, 0xa0, 0xac)),
    ("pink", Color::rgb(0xf5, 0xc2, 0xe7)),
    ("flamingo", Color::rgb(0xf2, 0xcd, 0xcd)),
    ("rosewater", Color::rgb(0xf5, 0xe0, 0xdc)),
    // Warm
    ("orange", Color::rgb(0xfa, 0xb3, 0x87)),
    ("peach", Color::rgb(0xfa, 0xb3, 0x87)),
    ("yellow", Color::rgb(0xf9, 0xe2, 0xaf)),
    // Greens
    ("green", Color::rgb(0xa6, 0xe3, 0xa1)),
    ("teal", Color::rgb(0x94, 0xe2, 0xd5)),
    // Blues
    ("cyan", Color::rgb(0x89, 0xdc, 0xeb)),
    ("sky", Color::rgb(0x89, 0xdc, 0xeb)),
    ("sapphire", Color::rgb(0x74, 0xc7, 0xec)),
    ("blue", Color::rgb(0x89, 0xb4, 0xfa)),
    ("lavender", Color::rgb(0xb4, 0xbe, 0xfe)),
    // Purples
    ("purple", Color::rgb(0xcb, 0xa6, 0xf7)),
    ("mauve", Color::rgb(0xcb, 0xa6, 0xf7)),
    // Neutrals
    ("white", Color::rgb(0xcd, 0xd6, 0xf4)),
    ("text", Color::rgb(0xcd, 0xd6, 0xf4)),
    ("gray", Color::rgb(0xa6, 0xad, 0xc8)),
    ("subtext", Color::rgb(0xa6, 0xad, 0xc8)),
    ("black", Color::rgb(0x11, 0x11, 0x1b)),
    ("crust", Color::rgb(0x11, 0x11, 0x1b)),
];

/// Look up a color by preset name (case-insensitive) or hex (`#rgb` /
/// `#rrggbb`). Returns None on unknown name or malformed hex.
pub fn parse_color_arg(s: &str) -> Option<Color> {
    let s = s.trim();
    if s.starts_with('#') {
        return parse_hex_color(s);
    }
    let lower = s.to_ascii_lowercase();
    COLOR_PRESETS
        .iter()
        .find(|(name, _)| *name == lower)
        .map(|(_, c)| *c)
}

/// Parse a hex color: `#rgb` or `#rrggbb` (case-insensitive).
/// `#fa3` expands to `(0xff, 0xaa, 0x33)`.
pub fn parse_hex_color(s: &str) -> Option<Color> {
    let hex = s.trim().strip_prefix('#')?;
    let (r, g, b) = match hex.len() {
        3 => {
            let bytes = hex.as_bytes();
            let r = hex_byte(bytes[0])? * 17;
            let g = hex_byte(bytes[1])? * 17;
            let b = hex_byte(bytes[2])? * 17;
            (r, g, b)
        }
        6 => {
            let bytes = hex.as_bytes();
            let r = hex_byte(bytes[0])? * 16 + hex_byte(bytes[1])?;
            let g = hex_byte(bytes[2])? * 16 + hex_byte(bytes[3])?;
            let b = hex_byte(bytes[4])? * 16 + hex_byte(bytes[5])?;
            (r, g, b)
        }
        _ => return None,
    };
    Some(Color { r, g, b })
}

fn hex_byte(b: u8) -> Option<u8> {
    match b {
        b'0'..=b'9' => Some(b - b'0'),
        b'a'..=b'f' => Some(b - b'a' + 10),
        b'A'..=b'F' => Some(b - b'A' + 10),
        _ => None,
    }
}

// ── JSON helpers ────────────────────────────────────────────────────

/// Extract `"key":"value"` from a flat JSON object. Tolerates whitespace
/// around the colon. Sufficient for the tiny payloads this module emits.
fn json_str_field(json: &str, key: &str) -> Option<String> {
    let needle = format!("\"{key}\"");
    let start = find_top_level_key(json, &needle)? + needle.len();
    let after_colon = json[start..].trim_start().strip_prefix(':')?.trim_start();
    let inside = after_colon.strip_prefix('"')?;
    let end = inside.find('"')?;
    Some(inside[..end].to_string())
}

/// Extract a non-negative integer field.
fn json_num_field(json: &str, key: &str) -> Option<u32> {
    let needle = format!("\"{key}\"");
    let start = find_top_level_key(json, &needle)? + needle.len();
    let after_colon = json[start..].trim_start().strip_prefix(':')?.trim_start();
    let end = after_colon
        .find(|c: char| !c.is_ascii_digit())
        .unwrap_or(after_colon.len());
    after_colon[..end].parse().ok()
}

/// Extract a boolean field (`"key":true` or `"key":false`).
fn json_bool_field(json: &str, key: &str) -> Option<bool> {
    let needle = format!("\"{key}\"");
    let start = find_top_level_key(json, &needle)? + needle.len();
    let after_colon = json[start..].trim_start().strip_prefix(':')?.trim_start();
    if let Some(rest) = after_colon.strip_prefix("true") {
        if rest.is_empty() || matches!(rest.as_bytes()[0], b',' | b'}' | b' ') {
            return Some(true);
        }
    }
    if let Some(rest) = after_colon.strip_prefix("false") {
        if rest.is_empty() || matches!(rest.as_bytes()[0], b',' | b'}' | b' ') {
            return Some(false);
        }
    }
    None
}

/// Extract a nested object value as a flat JSON string. Used to lift
/// `"n":{"k":"usd","d":2}` so it can be re-parsed with the flat helpers.
fn nested_obj(json: &str, key: &str) -> Option<String> {
    let needle = format!("\"{key}\"");
    let start = find_top_level_key(json, &needle)? + needle.len();
    let after_colon = json[start..].trim_start().strip_prefix(':')?.trim_start();
    let bytes = after_colon.as_bytes();
    if bytes.first() != Some(&b'{') {
        return None;
    }
    let mut depth = 0;
    for (i, &b) in bytes.iter().enumerate() {
        match b {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    return Some(after_colon[..=i].to_string());
                }
            }
            _ => {}
        }
    }
    None
}

/// Find a key at the top level of a flat JSON object — i.e. not inside
/// a nested `{}`. Without this, `json_str_field("…","k")` would match
/// the `"k"` inside `"n":{"k":"usd",…}` when looking for a top-level
/// `"k"` field that doesn't exist.
fn find_top_level_key(json: &str, needle: &str) -> Option<usize> {
    let bytes = json.as_bytes();
    let mut depth = 0i32;
    let mut i = 0;
    while i < bytes.len() {
        match bytes[i] {
            b'{' => depth += 1,
            b'}' => depth -= 1,
            b'"' if depth == 1 && json[i..].starts_with(needle) => {
                return Some(i);
            }
            _ => {}
        }
        i += 1;
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    fn usd(decimals: u8) -> CellFormat {
        CellFormat {
            number: Some(NumberFormat::Usd { decimals }),
            ..CellFormat::default()
        }
    }

    #[test]
    fn json_round_trip_usd_only() {
        let fmt = usd(2);
        let json = to_format_json(&fmt);
        assert_eq!(json, r#"{"n":{"k":"usd","d":2}}"#);
        assert_eq!(parse_format_json(&json), Some(fmt));
    }

    #[test]
    fn json_round_trip_zero_decimals() {
        let fmt = usd(0);
        assert_eq!(parse_format_json(&to_format_json(&fmt)), Some(fmt));
    }

    #[test]
    fn json_round_trip_all_axes() {
        let fmt = CellFormat {
            number: Some(NumberFormat::Usd { decimals: 4 }),
            bold: true,
            italic: true,
            underline: true,
            strike: true,
            align: Some(Align::Center),
            fg: Some(Color::rgb(0xff, 0x00, 0x00)),
            bg: Some(Color::rgb(0x1e, 0x1e, 0x2e)),
            date: Some("%Y-%m-%d".into()),
        };
        let json = to_format_json(&fmt);
        assert_eq!(parse_format_json(&json), Some(fmt));
    }

    #[test]
    fn json_round_trip_date_only() {
        let fmt = CellFormat {
            date: Some("%a %b %d".into()),
            ..CellFormat::default()
        };
        let json = to_format_json(&fmt);
        assert_eq!(parse_format_json(&json), Some(fmt));
    }

    #[test]
    fn json_default_emits_empty_object() {
        let fmt = CellFormat::default();
        assert_eq!(to_format_json(&fmt), "{}");
    }

    #[test]
    fn parse_clears_on_null_sentinel() {
        assert_eq!(parse_format_json("null"), None);
        assert_eq!(parse_format_json(""), None);
        assert_eq!(parse_format_json("  null  "), None);
    }

    #[test]
    fn parse_unknown_kind_yields_no_number_format() {
        let fmt = parse_format_json(r#"{"n":{"k":"bogus","d":2}}"#).unwrap();
        assert!(fmt.number.is_none());
    }

    #[test]
    fn parse_decimals_clamped_to_max() {
        let fmt = parse_format_json(r#"{"n":{"k":"usd","d":99}}"#).unwrap();
        assert_eq!(fmt.number, Some(NumberFormat::Usd { decimals: 10 }));
    }

    #[test]
    fn parse_align_strings() {
        for (s, expected) in [
            ("left", Align::Left),
            ("center", Align::Center),
            ("right", Align::Right),
        ] {
            let json = format!(r#"{{"a":"{s}"}}"#);
            assert_eq!(parse_format_json(&json).unwrap().align, Some(expected));
        }
    }

    #[test]
    fn parse_bool_flags() {
        let fmt =
            parse_format_json(r#"{"b":true,"i":false,"u":true,"s":false}"#).unwrap();
        assert!(fmt.bold);
        assert!(!fmt.italic);
        assert!(fmt.underline);
        assert!(!fmt.strike);
    }

    #[test]
    fn parse_colors_three_and_six_digit_hex() {
        let fmt =
            parse_format_json(r##"{"fg":"#fa3","bg":"#1e1e2e"}"##).unwrap();
        assert_eq!(fmt.fg, Some(Color::rgb(0xff, 0xaa, 0x33)));
        assert_eq!(fmt.bg, Some(Color::rgb(0x1e, 0x1e, 0x2e)));
    }

    #[test]
    fn parse_does_not_confuse_nested_keys_with_top_level() {
        // The nested `"k"` inside the `"n"` block must not be picked up
        // as a top-level key. Same for `"d"`. Easy to get wrong with a
        // naive substring search.
        let json = r#"{"n":{"k":"usd","d":2},"b":true}"#;
        let fmt = parse_format_json(json).unwrap();
        assert_eq!(fmt.number, Some(NumberFormat::Usd { decimals: 2 }));
        assert!(fmt.bold);
        // Top-level `"d"` doesn't exist; if it did, the parser would
        // need to ignore the nested one. Sanity check via num_field:
        assert_eq!(json_num_field(json, "d"), None);
    }

    #[test]
    fn currency_simple() {
        let (raw, fmt) = try_parse_typed_input("$1.25").unwrap();
        assert_eq!(raw, "1.25");
        assert_eq!(fmt.number, Some(NumberFormat::Usd { decimals: 2 }));
    }

    #[test]
    fn currency_thousands() {
        let (raw, fmt) = try_parse_typed_input("$1,234.56").unwrap();
        assert_eq!(raw, "1234.56");
        assert_eq!(fmt.number, Some(NumberFormat::Usd { decimals: 2 }));
    }

    #[test]
    fn currency_negative_outside_dollar() {
        let (raw, fmt) = try_parse_typed_input("-$1.25").unwrap();
        assert_eq!(raw, "-1.25");
        assert_eq!(fmt.number, Some(NumberFormat::Usd { decimals: 2 }));
    }

    #[test]
    fn currency_no_decimals_defaults_two() {
        let (raw, fmt) = try_parse_typed_input("$5").unwrap();
        assert_eq!(raw, "5");
        assert_eq!(fmt.number, Some(NumberFormat::Usd { decimals: 2 }));
    }

    #[test]
    fn currency_high_precision() {
        let (raw, fmt) = try_parse_typed_input("$0.12345").unwrap();
        assert_eq!(raw, "0.12345");
        assert_eq!(fmt.number, Some(NumberFormat::Usd { decimals: 5 }));
    }

    #[test]
    fn currency_leading_whitespace_ok() {
        assert!(try_parse_typed_input("  $1.25").is_some());
    }

    #[test]
    fn currency_rejects_text() {
        assert_eq!(try_parse_typed_input("$abc"), None);
    }

    #[test]
    fn currency_rejects_bare_dollar() {
        assert_eq!(try_parse_typed_input("$"), None);
        assert_eq!(try_parse_typed_input("-$"), None);
    }

    #[test]
    fn currency_rejects_trailing_dot() {
        assert_eq!(try_parse_typed_input("$1."), None);
    }

    #[test]
    fn currency_rejects_bad_grouping() {
        assert_eq!(try_parse_typed_input("$1,23"), None);
        assert_eq!(try_parse_typed_input("$,123"), None);
        assert_eq!(try_parse_typed_input("$12,3456"), None);
        assert_eq!(try_parse_typed_input("$1234,567"), None);
    }

    #[test]
    fn currency_rejects_minus_inside_dollar() {
        assert_eq!(try_parse_typed_input("$-1.25"), None);
    }

    #[test]
    fn render_basic() {
        let fmt = usd(2);
        assert_eq!(render("1.25", &fmt), "$1.25");
        assert_eq!(render("1234.56", &fmt), "$1,234.56");
        assert_eq!(render("0", &fmt), "$0.00");
    }

    #[test]
    fn render_negative_uses_dash_dollar() {
        let fmt = usd(2);
        assert_eq!(render("-3", &fmt), "-$3.00");
        assert_eq!(render("-1234.5", &fmt), "-$1,234.50");
    }

    #[test]
    fn render_zero_decimals() {
        let fmt = usd(0);
        assert_eq!(render("1234.99", &fmt), "$1,235");
    }

    #[test]
    fn render_passthrough_on_error_sentinel() {
        let fmt = usd(2);
        assert_eq!(render("#DIV/0!", &fmt), "#DIV/0!");
        assert_eq!(render("hello", &fmt), "hello");
    }

    #[test]
    fn render_passthrough_on_non_finite() {
        let fmt = usd(2);
        assert_eq!(render("inf", &fmt), "inf");
        assert_eq!(render("NaN", &fmt), "NaN");
    }

    #[test]
    fn render_no_number_format_passes_through() {
        // Style-only formats (bold/etc) don't transform text; render
        // returns the input verbatim so style still layers on top.
        let fmt = CellFormat {
            bold: true,
            ..CellFormat::default()
        };
        assert_eq!(render("hello", &fmt), "hello");
        assert_eq!(render("1.25", &fmt), "1.25");
    }

    #[test]
    fn render_large_thousands() {
        let fmt = usd(2);
        assert_eq!(render("1234567.89", &fmt), "$1,234,567.89");
        assert_eq!(render("-1000000", &fmt), "-$1,000,000.00");
    }

    #[test]
    fn render_rounding_boundary() {
        // f64 1.005 is actually 1.00499999...; {:.2} rounds to "1.00".
        // Documented here so we notice if Rust's formatter ever shifts.
        let fmt = usd(2);
        assert_eq!(render("1.005", &fmt), "$1.00");
        assert_eq!(render("1.015", &fmt), "$1.01");
    }

    #[test]
    fn bump_number_decimals_clamped() {
        let fmt = NumberFormat::Usd { decimals: 2 };
        assert_eq!(bump_number_decimals(&fmt, 1), NumberFormat::Usd { decimals: 3 });
        assert_eq!(bump_number_decimals(&fmt, -1), NumberFormat::Usd { decimals: 1 });
        assert_eq!(bump_number_decimals(&fmt, -10), NumberFormat::Usd { decimals: 0 });
        assert_eq!(bump_number_decimals(&fmt, 100), NumberFormat::Usd { decimals: 10 });
    }

    #[test]
    fn looks_numeric_recognizes_currency() {
        assert!(looks_numeric("$1.25"));
        assert!(looks_numeric("-$1.25"));
        assert!(looks_numeric("$0"));
        assert!(!looks_numeric("$"));
        assert!(!looks_numeric("$abc"));
        assert!(!looks_numeric("hello"));
        assert!(!looks_numeric(""));
    }

    #[test]
    fn looks_numeric_recognizes_percent() {
        assert!(looks_numeric("4.5%"));
        assert!(looks_numeric("100%"));
        assert!(looks_numeric("-100%"));
        assert!(looks_numeric("0%"));
        assert!(!looks_numeric("%"));
        assert!(!looks_numeric("%5"));
        assert!(!looks_numeric("abc%"));
    }

    #[test]
    fn percent_simple() {
        let (raw, fmt) = try_parse_typed_input("4.5%").unwrap();
        assert_eq!(raw, "0.045");
        assert_eq!(fmt.number, Some(NumberFormat::Percent { decimals: 1 }));
    }

    #[test]
    fn percent_integer_zero_decimals() {
        let (raw, fmt) = try_parse_typed_input("100%").unwrap();
        // 1.0 — with decimals + 2 = 2 fmt, then trim trailing zeros → "1".
        assert_eq!(raw, "1");
        assert_eq!(fmt.number, Some(NumberFormat::Percent { decimals: 0 }));
    }

    #[test]
    fn percent_negative() {
        let (raw, fmt) = try_parse_typed_input("-3%").unwrap();
        assert_eq!(raw, "-0.03");
        assert_eq!(fmt.number, Some(NumberFormat::Percent { decimals: 0 }));
    }

    #[test]
    fn percent_high_precision() {
        let (raw, fmt) = try_parse_typed_input("12.345%").unwrap();
        assert_eq!(fmt.number, Some(NumberFormat::Percent { decimals: 3 }));
        // raw should round-trip cleanly: 0.12345
        assert_eq!(raw, "0.12345");
    }

    #[test]
    fn percent_rejects_bad_input() {
        assert_eq!(try_parse_typed_input("%"), None);
        assert_eq!(try_parse_typed_input("-%"), None);
        assert_eq!(try_parse_typed_input("4.%"), None); // trailing dot
        assert_eq!(try_parse_typed_input("4%abc"), None); // we only check %-suffixed
    }

    #[test]
    fn render_percent_basic() {
        let fmt = CellFormat {
            number: Some(NumberFormat::Percent { decimals: 1 }),
            ..CellFormat::default()
        };
        assert_eq!(render("0.045", &fmt), "4.5%");
        assert_eq!(render("1", &fmt), "100.0%");
        assert_eq!(render("0", &fmt), "0.0%");
        assert_eq!(render("-0.03", &fmt), "-3.0%");
    }

    #[test]
    fn render_percent_zero_decimals() {
        let fmt = CellFormat {
            number: Some(NumberFormat::Percent { decimals: 0 }),
            ..CellFormat::default()
        };
        assert_eq!(render("0.05", &fmt), "5%");
        assert_eq!(render("1", &fmt), "100%");
    }

    #[test]
    fn percent_round_trip_typed_then_rendered() {
        // The auto-detect output should render back to (close to) the
        // original input.
        let (raw, fmt) = try_parse_typed_input("4.5%").unwrap();
        assert_eq!(render(&raw, &fmt), "4.5%");

        let (raw, fmt) = try_parse_typed_input("100%").unwrap();
        assert_eq!(render(&raw, &fmt), "100%");

        let (raw, fmt) = try_parse_typed_input("-12.34%").unwrap();
        assert_eq!(render(&raw, &fmt), "-12.34%");
    }

    #[test]
    fn json_round_trip_percent() {
        let fmt = CellFormat {
            number: Some(NumberFormat::Percent { decimals: 3 }),
            ..CellFormat::default()
        };
        let json = to_format_json(&fmt);
        assert_eq!(json, r#"{"n":{"k":"percent","d":3}}"#);
        assert_eq!(parse_format_json(&json), Some(fmt));
    }

    #[test]
    fn bump_decimals_works_for_percent_too() {
        let fmt = NumberFormat::Percent { decimals: 1 };
        assert_eq!(
            bump_number_decimals(&fmt, 2),
            NumberFormat::Percent { decimals: 3 }
        );
        assert_eq!(
            bump_number_decimals(&fmt, -5),
            NumberFormat::Percent { decimals: 0 }
        );
    }

    #[test]
    fn parse_color_arg_resolves_natural_names() {
        assert_eq!(parse_color_arg("red"), Some(Color::rgb(0xf3, 0x8b, 0xa8)));
        assert_eq!(parse_color_arg("RED"), Some(Color::rgb(0xf3, 0x8b, 0xa8)));
        assert_eq!(parse_color_arg("blue"), Some(Color::rgb(0x89, 0xb4, 0xfa)));
        assert_eq!(parse_color_arg("orange"), Some(Color::rgb(0xfa, 0xb3, 0x87)));
    }

    #[test]
    fn parse_color_arg_resolves_catppuccin_names() {
        // Aliases route to the same RGB as the natural names.
        assert_eq!(parse_color_arg("peach"), parse_color_arg("orange"));
        assert_eq!(parse_color_arg("mauve"), parse_color_arg("purple"));
        assert_eq!(parse_color_arg("sky"), parse_color_arg("cyan"));
    }

    #[test]
    fn parse_color_arg_falls_through_to_hex() {
        assert_eq!(
            parse_color_arg("#1e1e2e"),
            Some(Color::rgb(0x1e, 0x1e, 0x2e))
        );
    }

    #[test]
    fn parse_color_arg_rejects_unknown() {
        assert_eq!(parse_color_arg("notacolor"), None);
        assert_eq!(parse_color_arg(""), None);
    }

    #[test]
    fn parse_hex_color_three_digit() {
        assert_eq!(parse_hex_color("#fa3"), Some(Color::rgb(0xff, 0xaa, 0x33)));
        assert_eq!(parse_hex_color("#000"), Some(Color::rgb(0, 0, 0)));
        assert_eq!(parse_hex_color("#FFF"), Some(Color::rgb(0xff, 0xff, 0xff)));
    }

    #[test]
    fn parse_hex_color_six_digit() {
        assert_eq!(parse_hex_color("#1e1e2e"), Some(Color::rgb(0x1e, 0x1e, 0x2e)));
        assert_eq!(parse_hex_color("#FF0080"), Some(Color::rgb(0xff, 0x00, 0x80)));
    }

    #[test]
    fn parse_hex_color_rejects_bad_input() {
        assert_eq!(parse_hex_color("ff0000"), None); // missing #
        assert_eq!(parse_hex_color("#ff00"), None); // 4 digits
        assert_eq!(parse_hex_color("#fffffff"), None); // 7 digits
        assert_eq!(parse_hex_color("#xyz"), None); // bad hex
        assert_eq!(parse_hex_color(""), None);
    }
}
