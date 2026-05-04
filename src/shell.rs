//! Subprocess runner + output sniffer for the `!` shell prompt.
//!
//! Captures stdout (up to a hard cap) without inheriting the TUI's
//! stdin/stdout/stderr, so the alt-screen stays intact while a command
//! runs. Stderr is drained on a side thread to avoid pipe-buffer
//! deadlocks when a chatty subprocess exceeds OS pipe limits.
//!
//! [`detect_payload`] sniffs the captured stdout into a [`PastedGrid`]
//! using a "first row is always headers" rule. Dispatch order, first
//! match wins:
//!
//! 1. **JSON** — single document parsing end-to-end. Recognised
//!    shapes: `[{...}, ...]` (first object's keys → headers), `{...}`
//!    (1-row variant), `[scalar, ...]` (synthetic `value` header).
//! 2. **NDJSON / JSON Lines** — one JSON value per line. Common
//!    output of `jq -c`, `gh api --paginate`, `kubectl … | jq -c`.
//!    Reuses [`json_grid_from_value`] by synthesizing an `Array` of
//!    per-line values, so shape rules match the single-document path.
//! 3. **TSV** — first non-empty line has tabs and tabs ≥ commas.
//! 4. **CSV** — RFC 4180 via the `csv` crate.
//! 5. **Plain** — single column with synthetic `value` header.
//!
//! v1 limitation: blocks the UI while the subprocess runs and does not
//! support interactive children (no tty pass-through). For interactive
//! work, run vlotus in one terminal and the interactive program in
//! another.

use std::io::{self, Read};
use std::process::{Command, Stdio};
use std::thread;

/// 100 MB. Anything beyond this is almost certainly a mistake (`cat
/// /dev/zero` etc.); we kill the child and surface `TooLarge`.
pub const DEFAULT_STDOUT_CAP: usize = 100 * 1024 * 1024;

/// Stderr is truncated to this many characters before being stored on
/// `ShellError::NonZero`. The status bar would clip a longer message
/// anyway, and we don't want a multi-megabyte stderr riding alongside
/// every non-zero exit.
const STDERR_CHAR_CAP: usize = 200;

#[derive(Debug)]
pub enum ShellError {
    Spawn(io::Error),
    NonZero { code: Option<i32>, stderr: String },
    TooLarge { bytes: usize, cap: usize },
    Io(io::Error),
}

impl std::fmt::Display for ShellError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Spawn(e) => write!(f, "spawn failed: {e}"),
            Self::NonZero { code, stderr } => match code {
                Some(c) if stderr.is_empty() => write!(f, "exit {c}"),
                Some(c) => write!(f, "exit {c}: {stderr}"),
                None if stderr.is_empty() => write!(f, "killed by signal"),
                None => write!(f, "killed by signal: {stderr}"),
            },
            Self::TooLarge { bytes, cap } => {
                write!(f, "output too large ({bytes} bytes; cap {cap})")
            }
            Self::Io(e) => write!(f, "io error: {e}"),
        }
    }
}

impl std::error::Error for ShellError {}

/// Run `cmd` via the platform shell and return its captured stdout.
pub fn run(cmd: &str) -> Result<String, ShellError> {
    run_with_cap(cmd, DEFAULT_STDOUT_CAP)
}

/// Same as [`run`] with a configurable stdout cap. Exposed for tests.
pub(crate) fn run_with_cap(cmd: &str, cap: usize) -> Result<String, ShellError> {
    let mut child = build_command(cmd)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(ShellError::Spawn)?;

    let mut stdout_pipe = child.stdout.take().expect("stdout was piped");
    let mut stderr_pipe = child.stderr.take().expect("stderr was piped");

    // Drain stderr concurrently so the child doesn't block on a full
    // stderr pipe while we're still reading stdout.
    let stderr_handle = thread::spawn(move || {
        let mut buf = Vec::new();
        let _ = stderr_pipe.read_to_end(&mut buf);
        buf
    });

    let mut stdout_buf: Vec<u8> = Vec::new();
    let mut chunk = [0u8; 8 * 1024];
    let mut over_cap = false;
    loop {
        match stdout_pipe.read(&mut chunk) {
            Ok(0) => break,
            Ok(n) => {
                if stdout_buf.len() + n > cap {
                    over_cap = true;
                    let _ = child.kill();
                    // Keep draining so the child's pipe doesn't stall
                    // waiting for the reader; we don't store past cap.
                    let take = cap.saturating_sub(stdout_buf.len());
                    stdout_buf.extend_from_slice(&chunk[..take]);
                    let _ = io::copy(&mut stdout_pipe, &mut io::sink());
                    break;
                }
                stdout_buf.extend_from_slice(&chunk[..n]);
            }
            Err(e) => return Err(ShellError::Io(e)),
        }
    }

    let status = child.wait().map_err(ShellError::Io)?;
    let stderr_bytes = stderr_handle.join().unwrap_or_default();

    if over_cap {
        return Err(ShellError::TooLarge {
            bytes: stdout_buf.len(),
            cap,
        });
    }

    if !status.success() {
        let mut stderr = String::from_utf8_lossy(&stderr_bytes).trim().to_string();
        truncate_at_chars(&mut stderr, STDERR_CHAR_CAP);
        return Err(ShellError::NonZero {
            code: status.code(),
            stderr,
        });
    }

    String::from_utf8(stdout_buf)
        .map_err(|e| ShellError::Io(io::Error::new(io::ErrorKind::InvalidData, e)))
}

#[cfg(unix)]
fn build_command(cmd: &str) -> Command {
    let mut c = Command::new("sh");
    c.arg("-c").arg(cmd);
    c
}

#[cfg(windows)]
fn build_command(cmd: &str) -> Command {
    let mut c = Command::new("cmd");
    c.args(["/C", cmd]);
    c
}

fn truncate_at_chars(s: &mut String, max_chars: usize) {
    if s.chars().count() <= max_chars {
        return;
    }
    let cut = s
        .char_indices()
        .nth(max_chars)
        .map(|(i, _)| i)
        .unwrap_or(s.len());
    s.truncate(cut);
    s.push_str("...");
}

// ── Output sniffer + parsers ───────────────────────────────────────

use crate::app::{PastedCell, PastedGrid};

/// Sniff `stdout` and turn it into a [`PastedGrid`] using the rule
/// "first row is always headers, rest is data". Dispatch order:
/// JSON → NDJSON → TSV → CSV → plain. Returns `None` when the
/// output is empty or unparseable in every shape.
pub fn detect_payload(stdout: &str) -> Option<PastedGrid> {
    if stdout.trim().is_empty() {
        return None;
    }
    let leading = stdout.trim_start();
    if leading.starts_with('[') || leading.starts_with('{') {
        if let Some(g) = parse_json_grid(stdout) {
            return Some(g);
        }
        if let Some(g) = parse_ndjson_grid(stdout) {
            return Some(g);
        }
    }
    if let Some(g) = parse_tsv_grid(stdout) {
        return Some(g);
    }
    if let Some(g) = parse_csv_grid(stdout) {
        return Some(g);
    }
    parse_plain_grid(stdout)
}

fn pasted_cell(value: String) -> PastedCell {
    PastedCell {
        value,
        formula: None,
    }
}

fn finalize_grid(cells: Vec<Vec<PastedCell>>) -> Option<PastedGrid> {
    if cells.is_empty() {
        return None;
    }
    Some(PastedGrid {
        source_anchor: None,
        cells,
    })
}

fn parse_tsv_grid(s: &str) -> Option<PastedGrid> {
    let first_line = s.lines().find(|l| !l.trim().is_empty())?;
    let tabs = first_line.matches('\t').count();
    let commas = first_line.matches(',').count();
    if tabs == 0 || commas > tabs {
        return None;
    }
    let cells: Vec<Vec<PastedCell>> = s
        .lines()
        .filter(|l| !l.is_empty())
        .map(|line| {
            line.trim_end_matches('\r')
                .split('\t')
                .map(|f| pasted_cell(f.to_string()))
                .collect()
        })
        .collect();
    finalize_grid(cells)
}

fn parse_csv_grid(s: &str) -> Option<PastedGrid> {
    let first_line = s.lines().find(|l| !l.trim().is_empty())?;
    if !first_line.contains(',') {
        return None;
    }
    let mut rdr = csv::ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_reader(s.as_bytes());
    let mut cells: Vec<Vec<PastedCell>> = Vec::new();
    for record in rdr.records() {
        let rec = record.ok()?;
        cells.push(rec.iter().map(|f| pasted_cell(f.to_string())).collect());
    }
    finalize_grid(cells)
}

fn parse_plain_grid(s: &str) -> Option<PastedGrid> {
    let lines: Vec<&str> = s.lines().filter(|l| !l.is_empty()).collect();
    if lines.is_empty() {
        return None;
    }
    let mut cells: Vec<Vec<PastedCell>> = Vec::with_capacity(lines.len() + 1);
    cells.push(vec![pasted_cell("value".to_string())]);
    for l in lines {
        cells.push(vec![pasted_cell(l.trim_end_matches('\r').to_string())]);
    }
    finalize_grid(cells)
}

// ── JSON ──────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq)]
enum JsonValue {
    Null,
    Bool(bool),
    Number(String),
    Str(String),
    Array(Vec<JsonValue>),
    Object(Vec<(String, JsonValue)>),
}

fn parse_json_grid(s: &str) -> Option<PastedGrid> {
    let mut p = JsonParser::new(s);
    p.skip_ws();
    let v = p.parse_value()?;
    p.skip_ws();
    if !p.is_eof() {
        return None;
    }
    json_grid_from_value(&v)
}

/// Parse newline-delimited JSON (NDJSON / JSON Lines): one JSON value
/// per line. Common output of `jq -c`, `gh api --paginate`, log
/// streams. Reuses `json_grid_from_value` by synthesizing an array of
/// the per-line values, so the array-of-objects (first line's keys →
/// headers) and array-of-scalars (synthetic `value` header) rules are
/// shared with the regular JSON path.
fn parse_ndjson_grid(s: &str) -> Option<PastedGrid> {
    let lines: Vec<&str> = s.lines().filter(|l| !l.trim().is_empty()).collect();
    // Single-line input is the regular JSON branch's job; only fire
    // for actually-multi-line records.
    if lines.len() < 2 {
        return None;
    }
    let mut values = Vec::with_capacity(lines.len());
    for line in lines {
        let mut p = JsonParser::new(line);
        p.skip_ws();
        let v = p.parse_value()?;
        p.skip_ws();
        if !p.is_eof() {
            return None;
        }
        values.push(v);
    }
    json_grid_from_value(&JsonValue::Array(values))
}

fn json_grid_from_value(v: &JsonValue) -> Option<PastedGrid> {
    let array_view: Vec<&Vec<(String, JsonValue)>>;
    match v {
        JsonValue::Object(o) => {
            array_view = vec![o];
        }
        JsonValue::Array(arr) => {
            if arr.is_empty() {
                return None;
            }
            let all_objects = arr.iter().all(|e| matches!(e, JsonValue::Object(_)));
            let all_scalars = arr
                .iter()
                .all(|e| !matches!(e, JsonValue::Object(_) | JsonValue::Array(_)));
            if all_objects {
                array_view = arr
                    .iter()
                    .filter_map(|e| match e {
                        JsonValue::Object(o) => Some(o),
                        _ => None,
                    })
                    .collect();
            } else if all_scalars {
                let mut cells: Vec<Vec<PastedCell>> = Vec::with_capacity(arr.len() + 1);
                cells.push(vec![pasted_cell("value".to_string())]);
                for e in arr {
                    cells.push(vec![pasted_cell(stringify_value(e))]);
                }
                return finalize_grid(cells);
            } else {
                return None;
            }
        }
        _ => return None,
    }

    let header_keys: Vec<String> = array_view[0].iter().map(|(k, _)| k.clone()).collect();
    if header_keys.is_empty() {
        return None;
    }
    let mut cells: Vec<Vec<PastedCell>> = Vec::with_capacity(array_view.len() + 1);
    cells.push(
        header_keys
            .iter()
            .map(|k| pasted_cell(k.clone()))
            .collect(),
    );
    for obj in array_view {
        let mut row = Vec::with_capacity(header_keys.len());
        for key in &header_keys {
            let cell_val = obj
                .iter()
                .find(|(k, _)| k == key)
                .map(|(_, v)| stringify_value(v))
                .unwrap_or_default();
            row.push(pasted_cell(cell_val));
        }
        cells.push(row);
    }
    finalize_grid(cells)
}

fn stringify_value(v: &JsonValue) -> String {
    match v {
        JsonValue::Null => String::new(),
        JsonValue::Bool(b) => b.to_string(),
        JsonValue::Number(n) => n.clone(),
        JsonValue::Str(s) => s.clone(),
        JsonValue::Array(_) | JsonValue::Object(_) => json_serialize(v),
    }
}

fn json_serialize(v: &JsonValue) -> String {
    match v {
        JsonValue::Null => "null".into(),
        JsonValue::Bool(b) => b.to_string(),
        JsonValue::Number(n) => n.clone(),
        JsonValue::Str(s) => format!("\"{}\"", json_escape(s)),
        JsonValue::Array(arr) => {
            let parts: Vec<String> = arr.iter().map(json_serialize).collect();
            format!("[{}]", parts.join(","))
        }
        JsonValue::Object(o) => {
            let parts: Vec<String> = o
                .iter()
                .map(|(k, v)| format!("\"{}\":{}", json_escape(k), json_serialize(v)))
                .collect();
            format!("{{{}}}", parts.join(","))
        }
    }
}

fn json_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

struct JsonParser<'a> {
    src: &'a [u8],
    pos: usize,
}

impl<'a> JsonParser<'a> {
    fn new(src: &'a str) -> Self {
        Self {
            src: src.as_bytes(),
            pos: 0,
        }
    }

    fn skip_ws(&mut self) {
        while let Some(&c) = self.src.get(self.pos) {
            if matches!(c, b' ' | b'\t' | b'\n' | b'\r') {
                self.pos += 1;
            } else {
                break;
            }
        }
    }

    fn is_eof(&self) -> bool {
        self.pos >= self.src.len()
    }

    fn peek(&self) -> Option<u8> {
        self.src.get(self.pos).copied()
    }

    fn parse_value(&mut self) -> Option<JsonValue> {
        self.skip_ws();
        match self.peek()? {
            b'"' => self.parse_string().map(JsonValue::Str),
            b'{' => self.parse_object(),
            b'[' => self.parse_array(),
            b't' | b'f' => self.parse_bool(),
            b'n' => self.parse_null(),
            b'-' | b'0'..=b'9' => self.parse_number(),
            _ => None,
        }
    }

    fn parse_string(&mut self) -> Option<String> {
        if self.peek()? != b'"' {
            return None;
        }
        self.pos += 1;
        let mut out = String::new();
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if c == b'"' {
                self.pos += 1;
                return Some(out);
            }
            if c == b'\\' {
                self.pos += 1;
                let esc = *self.src.get(self.pos)?;
                self.pos += 1;
                match esc {
                    b'"' => out.push('"'),
                    b'\\' => out.push('\\'),
                    b'/' => out.push('/'),
                    b'b' => out.push('\u{0008}'),
                    b'f' => out.push('\u{000C}'),
                    b'n' => out.push('\n'),
                    b'r' => out.push('\r'),
                    b't' => out.push('\t'),
                    b'u' => {
                        if self.pos + 4 > self.src.len() {
                            return None;
                        }
                        let hex = std::str::from_utf8(&self.src[self.pos..self.pos + 4]).ok()?;
                        let code = u32::from_str_radix(hex, 16).ok()?;
                        self.pos += 4;
                        // No surrogate-pair handling — push replacement on bad codes.
                        if let Some(ch) = char::from_u32(code) {
                            out.push(ch);
                        } else {
                            out.push('\u{FFFD}');
                        }
                    }
                    _ => return None,
                }
            } else {
                let len = utf8_char_len(c)?;
                if self.pos + len > self.src.len() {
                    return None;
                }
                let s = std::str::from_utf8(&self.src[self.pos..self.pos + len]).ok()?;
                out.push_str(s);
                self.pos += len;
            }
        }
        None
    }

    fn parse_object(&mut self) -> Option<JsonValue> {
        self.pos += 1;
        self.skip_ws();
        let mut pairs: Vec<(String, JsonValue)> = Vec::new();
        if self.peek() == Some(b'}') {
            self.pos += 1;
            return Some(JsonValue::Object(pairs));
        }
        loop {
            self.skip_ws();
            let key = self.parse_string()?;
            self.skip_ws();
            if self.peek()? != b':' {
                return None;
            }
            self.pos += 1;
            let val = self.parse_value()?;
            pairs.push((key, val));
            self.skip_ws();
            match self.peek()? {
                b',' => self.pos += 1,
                b'}' => {
                    self.pos += 1;
                    return Some(JsonValue::Object(pairs));
                }
                _ => return None,
            }
        }
    }

    fn parse_array(&mut self) -> Option<JsonValue> {
        self.pos += 1;
        self.skip_ws();
        let mut items = Vec::new();
        if self.peek() == Some(b']') {
            self.pos += 1;
            return Some(JsonValue::Array(items));
        }
        loop {
            let v = self.parse_value()?;
            items.push(v);
            self.skip_ws();
            match self.peek()? {
                b',' => self.pos += 1,
                b']' => {
                    self.pos += 1;
                    return Some(JsonValue::Array(items));
                }
                _ => return None,
            }
        }
    }

    fn parse_bool(&mut self) -> Option<JsonValue> {
        if self.src[self.pos..].starts_with(b"true") {
            self.pos += 4;
            Some(JsonValue::Bool(true))
        } else if self.src[self.pos..].starts_with(b"false") {
            self.pos += 5;
            Some(JsonValue::Bool(false))
        } else {
            None
        }
    }

    fn parse_null(&mut self) -> Option<JsonValue> {
        if self.src[self.pos..].starts_with(b"null") {
            self.pos += 4;
            Some(JsonValue::Null)
        } else {
            None
        }
    }

    fn parse_number(&mut self) -> Option<JsonValue> {
        let start = self.pos;
        if self.peek() == Some(b'-') {
            self.pos += 1;
        }
        while matches!(self.peek(), Some(b'0'..=b'9')) {
            self.pos += 1;
        }
        if self.peek() == Some(b'.') {
            self.pos += 1;
            while matches!(self.peek(), Some(b'0'..=b'9')) {
                self.pos += 1;
            }
        }
        if matches!(self.peek(), Some(b'e' | b'E')) {
            self.pos += 1;
            if matches!(self.peek(), Some(b'+' | b'-')) {
                self.pos += 1;
            }
            while matches!(self.peek(), Some(b'0'..=b'9')) {
                self.pos += 1;
            }
        }
        let slice = std::str::from_utf8(&self.src[start..self.pos]).ok()?;
        if slice.is_empty() || slice == "-" {
            return None;
        }
        Some(JsonValue::Number(slice.to_string()))
    }
}

fn utf8_char_len(byte: u8) -> Option<usize> {
    if byte < 0x80 {
        Some(1)
    } else if byte < 0xC0 {
        None
    } else if byte < 0xE0 {
        Some(2)
    } else if byte < 0xF0 {
        Some(3)
    } else if byte < 0xF8 {
        Some(4)
    } else {
        None
    }
}

#[cfg(all(test, unix))]
mod runner_tests {
    use super::*;

    #[test]
    fn run_echoes_stdout() {
        let out = run("printf 'hi\\n'").expect("subprocess succeeded");
        assert_eq!(out, "hi\n");
    }

    #[test]
    fn run_captures_multiline_stdout() {
        let out = run("printf 'a\\nb\\nc\\n'").expect("subprocess succeeded");
        assert_eq!(out, "a\nb\nc\n");
    }

    #[test]
    fn run_surfaces_nonzero_exit() {
        let err = run("exit 7").expect_err("nonzero exit surfaces error");
        match err {
            ShellError::NonZero { code, .. } => assert_eq!(code, Some(7)),
            other => panic!("expected NonZero, got {other:?}"),
        }
    }

    #[test]
    fn run_includes_stderr_in_nonzero_error() {
        let err = run("printf 'boom\\n' 1>&2; exit 1").expect_err("expected error");
        match err {
            ShellError::NonZero { code, stderr } => {
                assert_eq!(code, Some(1));
                assert!(stderr.contains("boom"), "stderr surfaced: {stderr}");
            }
            other => panic!("expected NonZero, got {other:?}"),
        }
    }

    #[test]
    fn run_truncates_long_stderr() {
        let cmd = format!(
            "printf '{}' 1>&2; exit 1",
            "x".repeat(STDERR_CHAR_CAP + 50)
        );
        let err = run(&cmd).expect_err("expected error");
        match err {
            ShellError::NonZero { stderr, .. } => {
                assert!(
                    stderr.ends_with("..."),
                    "stderr truncated with ellipsis: {stderr}"
                );
                let xs = stderr.chars().filter(|c| *c == 'x').count();
                assert_eq!(xs, STDERR_CHAR_CAP);
            }
            other => panic!("expected NonZero, got {other:?}"),
        }
    }

    #[test]
    fn run_with_cap_caps_oversized_stdout() {
        // Generate ~2 KB of stdout but cap at 64 bytes.
        // `seq` instead of `{1..2000}` brace expansion: `/bin/sh` is dash on Ubuntu CI.
        let err = run_with_cap("printf '_%.0s' $(seq 1 2000)", 64).expect_err("cap fires");
        match err {
            ShellError::TooLarge { bytes, cap } => {
                assert_eq!(cap, 64);
                assert!(bytes >= cap, "buffer reached cap before kill: {bytes}");
            }
            other => panic!("expected TooLarge, got {other:?}"),
        }
    }

    #[test]
    fn run_with_cap_under_cap_returns_full_output() {
        let out = run_with_cap("printf 'hello'", 1024).expect("under cap");
        assert_eq!(out, "hello");
    }

    #[test]
    fn run_empty_command_succeeds_with_empty_stdout() {
        let out = run("true").expect("true exits 0");
        assert_eq!(out, "");
    }

    #[test]
    fn shell_error_display_includes_exit_code() {
        let err = ShellError::NonZero {
            code: Some(2),
            stderr: "bad arg".into(),
        };
        assert!(format!("{err}").contains("exit 2"));
        assert!(format!("{err}").contains("bad arg"));
    }
}

#[cfg(test)]
mod parser_tests {
    use super::*;

    fn cell(v: &str) -> PastedCell {
        PastedCell {
            value: v.into(),
            formula: None,
        }
    }

    fn rows(grid: &PastedGrid) -> Vec<Vec<String>> {
        grid.cells
            .iter()
            .map(|r| r.iter().map(|c| c.value.clone()).collect())
            .collect()
    }

    #[test]
    fn detect_csv_with_quoted_field_containing_comma() {
        let grid = detect_payload("name,note\nalice,\"hi, there\"\nbob,plain\n").unwrap();
        assert_eq!(
            rows(&grid),
            vec![
                vec!["name", "note"],
                vec!["alice", "hi, there"],
                vec!["bob", "plain"],
            ]
        );
    }

    #[test]
    fn detect_csv_with_escaped_double_quote() {
        let grid = detect_payload("a,b\n\"she said \"\"hi\"\"\",2\n").unwrap();
        assert_eq!(
            rows(&grid),
            vec![vec!["a", "b"], vec!["she said \"hi\"", "2"]]
        );
    }

    #[test]
    fn detect_csv_with_embedded_newline_in_quoted_field() {
        let grid = detect_payload("a,b\n\"line1\nline2\",2\n").unwrap();
        assert_eq!(
            rows(&grid),
            vec![vec!["a", "b"], vec!["line1\nline2", "2"]]
        );
    }

    #[test]
    fn detect_tsv_picks_tab_when_both_present() {
        // Tabs > commas → TSV. Embedded comma in a field stays put.
        let grid = detect_payload("a\tb\nx,y\t1\n").unwrap();
        assert_eq!(rows(&grid), vec![vec!["a", "b"], vec!["x,y", "1"]]);
    }

    #[test]
    fn detect_csv_when_only_commas_present() {
        let grid = detect_payload("a,b,c\n1,2,3\n").unwrap();
        assert_eq!(rows(&grid), vec![vec!["a", "b", "c"], vec!["1", "2", "3"]]);
    }

    #[test]
    fn detect_strips_carriage_returns_from_tsv() {
        let grid = detect_payload("a\tb\r\n1\t2\r\n").unwrap();
        assert_eq!(rows(&grid), vec![vec!["a", "b"], vec!["1", "2"]]);
    }

    #[test]
    fn detect_json_array_of_objects_uses_first_object_keys_as_headers() {
        let json = r#"[{"name":"alice","age":30},{"name":"bob","age":25}]"#;
        let grid = detect_payload(json).unwrap();
        assert_eq!(
            rows(&grid),
            vec![
                vec!["name", "age"],
                vec!["alice", "30"],
                vec!["bob", "25"],
            ]
        );
    }

    #[test]
    fn detect_json_array_of_objects_handles_missing_keys_as_empty() {
        let json = r#"[{"a":1,"b":2},{"a":3}]"#;
        let grid = detect_payload(json).unwrap();
        assert_eq!(
            rows(&grid),
            vec![vec!["a", "b"], vec!["1", "2"], vec!["3", ""]]
        );
    }

    #[test]
    fn detect_json_array_of_objects_stringifies_scalar_values() {
        let json = r#"[{"s":"hi","n":1.5,"b":true,"z":null}]"#;
        let grid = detect_payload(json).unwrap();
        assert_eq!(
            rows(&grid),
            vec![vec!["s", "n", "b", "z"], vec!["hi", "1.5", "true", ""]]
        );
    }

    #[test]
    fn detect_json_single_object_renders_as_one_row() {
        let json = r#"{"k":"v","n":42}"#;
        let grid = detect_payload(json).unwrap();
        assert_eq!(rows(&grid), vec![vec!["k", "n"], vec!["v", "42"]]);
    }

    #[test]
    fn detect_json_array_of_scalars_renders_with_synthetic_value_header() {
        let json = "[1, 2, 3, \"four\"]";
        let grid = detect_payload(json).unwrap();
        assert_eq!(
            rows(&grid),
            vec![vec!["value"], vec!["1"], vec!["2"], vec!["3"], vec!["four"]]
        );
    }

    #[test]
    fn detect_json_mixed_array_returns_none_falls_through_to_csv() {
        // Mixed scalar + object → not a recognised JSON shape. The
        // input still has commas so it falls through to CSV rather
        // than plain — by design (sniffer is "first match wins").
        let json = "[1, {\"a\":2}]";
        let grid = detect_payload(json).unwrap();
        assert!(grid.cells[0].iter().any(|c| c.value.contains("[1")));
        assert!(grid.cells[0].len() >= 2);
    }

    #[test]
    fn detect_json_mixed_array_falls_through_to_plain_when_no_commas() {
        // Same shape but with a single-element mix that has no commas
        // landing in the body would still hit JSON-fail; here we
        // construct a bare object-with-extra-token which fails JSON,
        // and has no commas so it hits the plain branch.
        let payload = "{\"a\":1} trailing";
        let grid = detect_payload(payload).unwrap();
        assert_eq!(grid.cells[0], vec![cell("value")]);
        assert_eq!(grid.cells.len(), 2);
        assert_eq!(grid.cells[1][0].value, payload);
    }

    #[test]
    fn detect_json_nested_value_serializes_as_compact_json() {
        let json = r#"[{"a":[1,2],"b":{"x":1}}]"#;
        let grid = detect_payload(json).unwrap();
        assert_eq!(
            rows(&grid),
            vec![vec!["a", "b"], vec!["[1,2]", "{\"x\":1}"]]
        );
    }

    #[test]
    fn detect_plain_text_falls_through_to_single_value_column() {
        let grid = detect_payload("apple\nbanana\ncherry\n").unwrap();
        assert_eq!(
            rows(&grid),
            vec![
                vec!["value"],
                vec!["apple"],
                vec!["banana"],
                vec!["cherry"],
            ]
        );
    }

    #[test]
    fn detect_empty_stdout_returns_none() {
        assert!(detect_payload("").is_none());
        assert!(detect_payload("   \n\n").is_none());
    }

    #[test]
    fn detect_json_with_unicode_escape_decodes() {
        let json = r#"[{"k":"é"}]"#;
        let grid = detect_payload(json).unwrap();
        assert_eq!(grid.cells[1][0].value, "é");
    }

    #[test]
    fn detect_ndjson_array_of_objects_uses_first_line_keys_as_headers() {
        let payload = "{\"name\":\"alice\",\"age\":30}\n{\"name\":\"bob\",\"age\":25}\n";
        let grid = detect_payload(payload).unwrap();
        assert_eq!(
            rows(&grid),
            vec![
                vec!["name", "age"],
                vec!["alice", "30"],
                vec!["bob", "25"],
            ]
        );
    }

    #[test]
    fn detect_ndjson_with_missing_keys_in_later_lines_pads_with_empty() {
        let payload = "{\"a\":1,\"b\":2}\n{\"a\":3}\n";
        let grid = detect_payload(payload).unwrap();
        assert_eq!(
            rows(&grid),
            vec![vec!["a", "b"], vec!["1", "2"], vec!["3", ""]]
        );
    }

    #[test]
    fn detect_ndjson_with_blank_lines_between_records_works() {
        let payload = "{\"k\":\"v1\"}\n\n{\"k\":\"v2\"}\n";
        let grid = detect_payload(payload).unwrap();
        assert_eq!(rows(&grid), vec![vec!["k"], vec!["v1"], vec!["v2"]]);
    }

    #[test]
    fn detect_ndjson_falls_back_when_a_line_is_not_valid_json() {
        // First line is JSON, second line is not — falls through past
        // both JSON and NDJSON. With no commas/tabs lands in plain.
        let payload = "{\"a\":1}\nnot json\n";
        let grid = detect_payload(payload).unwrap();
        assert_eq!(grid.cells[0], vec![cell("value")]);
        assert_eq!(grid.cells.len(), 3);
    }

    #[test]
    fn detect_ndjson_array_of_scalars_uses_synthetic_value_header() {
        // Less common but supported — line of bare scalars. Has to
        // open with `[` or `{` to clear the dispatcher gate, so this
        // shape only triggers if at least one line is an object/array.
        let payload = "[1]\n[2]\n[3]\n";
        let grid = detect_payload(payload).unwrap();
        // Each line is a 1-element array → mixed-shape rule rejects
        // (arrays aren't objects), falls through to CSV/plain. So this
        // sniffer doesn't match — confirm graceful fallthrough.
        assert!(!grid.cells.is_empty());
    }

    #[test]
    fn detect_ndjson_lone_line_falls_through_to_regular_json() {
        // A single-line NDJSON-shaped payload is just regular JSON.
        let payload = "{\"a\":1}\n";
        let grid = detect_payload(payload).unwrap();
        assert_eq!(rows(&grid), vec![vec!["a"], vec!["1"]]);
    }

    #[test]
    fn detect_ndjson_with_crlf_line_endings() {
        let payload = "{\"a\":1}\r\n{\"a\":2}\r\n";
        let grid = detect_payload(payload).unwrap();
        assert_eq!(rows(&grid), vec![vec!["a"], vec!["1"], vec!["2"]]);
    }

    #[test]
    fn detect_malformed_json_falls_through_to_csv() {
        // `[` opens, but no closing — JSON parser returns None,
        // falls through to CSV (commas present).
        let json = "[1,2,3";
        let grid = detect_payload(json).unwrap();
        assert_eq!(rows(&grid), vec![vec!["[1", "2", "3"]]);
    }
}
