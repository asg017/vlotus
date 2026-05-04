//! `vlotus patch …` subcommand. P2 of att `coisl46s`.
//!
//! Operates on `.lpatch` files (raw SQLite changeset bytes) without
//! launching the TUI — apply onto a workbook, print a human-readable
//! diff, write the inverse, or concatenate via SQLite's changegroup
//! API.

use std::fs::File;
use std::io::{BufWriter, Read, Write};
use std::path::Path;

use fallible_streaming_iterator::FallibleStreamingIterator;
use rusqlite::hooks::Action;
use rusqlite::session::{Changegroup, Changeset, ChangesetIter};
use rusqlite::types::ValueRef;

use crate::store::{
    coords::col_idx_to_letters, patch::read_patch_file, ConflictPolicy, Store, StoreError,
};
use crate::{ConflictPolicy as CliConflictPolicy, PatchOp};

pub fn dispatch(op: PatchOp) -> i32 {
    match op {
        PatchOp::Apply {
            db,
            patch,
            invert,
            on_conflict,
        } => run_apply(&db, &patch, invert, on_conflict.into()),
        PatchOp::Show { patch } => run_show(&patch),
        PatchOp::Invert { input, output } => run_invert(&input, &output),
        PatchOp::Combine { output, patches } => run_combine(&output, &patches),
        PatchOp::Diff { from, to, output } => run_diff(&from, &to, &output),
    }
}

impl From<CliConflictPolicy> for ConflictPolicy {
    fn from(p: CliConflictPolicy) -> Self {
        match p {
            CliConflictPolicy::Omit => ConflictPolicy::Omit,
            CliConflictPolicy::Replace => ConflictPolicy::Replace,
            CliConflictPolicy::Abort => ConflictPolicy::Abort,
        }
    }
}

pub fn run_apply(db: &str, patch: &str, invert: bool, policy: ConflictPolicy) -> i32 {
    let bytes = match read_patch_file(Path::new(patch)) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("read patch: {e}");
            return 1;
        }
    };
    let bytes = if invert {
        match invert_bytes(&bytes) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("invert: {e}");
                return 1;
            }
        }
    } else {
        bytes
    };

    let mut store = match Store::open(Path::new(db)) {
        Ok(s) => s,
        Err(e) => {
            eprintln!("open db: {e}");
            return 1;
        }
    };
    let conflicts = match store.apply_changeset(&bytes, policy) {
        Ok(n) => n,
        Err(e) => {
            eprintln!("apply: {e}");
            return 1;
        }
    };

    // Recalc every sheet that exists post-apply; the patch may have
    // touched any of them, and the cell.computed columns regenerate
    // from raw input. Recompute sheet list AFTER apply so newly-
    // patched-in sheets are included.
    let sheet_names: Vec<String> = match store.list_sheets() {
        Ok(s) => s.into_iter().map(|m| m.name).collect(),
        Err(e) => {
            eprintln!("list sheets: {e}");
            return 1;
        }
    };
    for name in sheet_names {
        if let Err(e) = store.recalculate(&name) {
            eprintln!("recalculate {name}: {e}");
            return 1;
        }
    }
    if let Err(e) = store.commit() {
        eprintln!("commit: {e}");
        return 1;
    }
    if conflicts == 0 {
        println!("applied {patch}");
    } else {
        println!("applied {patch} ({conflicts} conflict(s))");
    }
    0
}

pub fn run_show(patch: &str) -> i32 {
    let bytes = match read_patch_file(Path::new(patch)) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("read patch: {e}");
            return 1;
        }
    };
    match render_changeset(&bytes, &mut std::io::stdout().lock()) {
        Ok(()) => 0,
        Err(e) => {
            eprintln!("show: {e}");
            1
        }
    }
}

pub fn run_invert(input: &str, output: &str) -> i32 {
    let bytes = match read_patch_file(Path::new(input)) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("read input: {e}");
            return 1;
        }
    };
    let inverted = match invert_bytes(&bytes) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("invert: {e}");
            return 1;
        }
    };
    if let Err(e) = atomic_write(Path::new(output), &inverted) {
        eprintln!("write output: {e}");
        return 1;
    }
    println!("wrote {output}");
    0
}

pub fn run_combine(output: &str, patches: &[String]) -> i32 {
    let mut group = match Changegroup::new() {
        Ok(g) => g,
        Err(e) => {
            eprintln!("changegroup: {e}");
            return 1;
        }
    };
    for path in patches {
        let bytes = match read_patch_file(Path::new(path)) {
            Ok(b) => b,
            Err(e) => {
                eprintln!("read {path}: {e}");
                return 1;
            }
        };
        let mut input: &[u8] = &bytes;
        let mut input_dyn: &mut dyn Read = &mut input;
        if let Err(e) = group.add_stream(&mut input_dyn) {
            eprintln!("add {path} to changegroup: {e}");
            return 1;
        }
    }
    let combined = match group.output() {
        Ok(c) => c,
        Err(e) => {
            eprintln!("combine: {e}");
            return 1;
        }
    };
    if let Err(e) = write_changeset(Path::new(output), &combined) {
        eprintln!("write output: {e}");
        return 1;
    }
    println!("wrote {output}");
    0
}

pub fn run_diff(from: &str, to: &str, output: &str) -> i32 {
    let store = match Store::open(Path::new(to)) {
        Ok(s) => s,
        Err(e) => {
            eprintln!("open {to}: {e}");
            return 1;
        }
    };
    let conn = store.conn();
    if let Err(e) = conn.execute("ATTACH DATABASE ?1 AS base", rusqlite::params![from]) {
        eprintln!("attach {from} as base: {e}");
        return 1;
    }
    let mut session = match rusqlite::session::Session::new(conn) {
        Ok(s) => s,
        Err(e) => {
            eprintln!("session: {e}");
            return 1;
        }
    };
    session.table_filter(Some(|tbl: &str| {
        matches!(tbl, "sheet" | "column_meta" | "cell" | "cell_format")
    }));
    for table in ["sheet", "column_meta", "cell", "cell_format"] {
        if let Err(e) = session.diff(rusqlite::DatabaseName::Attached("base"), table) {
            // SQLITE_SCHEMA on a missing/incompatible table is
            // surfaced; skip those tables silently rather than
            // bail out — the destination simply lacks that table.
            eprintln!("diff {table}: {e}");
        }
    }
    let mut buf: Vec<u8> = Vec::new();
    let buf_dyn: &mut dyn Write = &mut buf;
    if let Err(e) = session.changeset_strm(buf_dyn) {
        eprintln!("changeset_strm: {e}");
        return 1;
    }
    if let Err(e) = atomic_write(Path::new(output), &buf) {
        eprintln!("write output: {e}");
        return 1;
    }
    println!("wrote {output}");
    0
}

// ── Helpers ──────────────────────────────────────────────────────────

fn invert_bytes(bytes: &[u8]) -> Result<Vec<u8>, StoreError> {
    use rusqlite::session::invert_strm;
    let mut input: &[u8] = bytes;
    let mut input_dyn: &mut dyn Read = &mut input;
    let mut output: Vec<u8> = Vec::new();
    let mut output_dyn: &mut dyn Write = &mut output;
    invert_strm(&mut input_dyn, &mut output_dyn)?;
    Ok(output)
}

fn write_changeset(path: &Path, cs: &Changeset) -> Result<(), StoreError> {
    let tmp = path.with_extension("lpatch.tmp");
    {
        let file = File::create(&tmp)?;
        let mut buf = BufWriter::new(file);
        let buf_dyn: &mut dyn Write = &mut buf;
        // Workaround: rusqlite doesn't expose Changeset bytes
        // directly, but Changegroup::output_strm is symmetric.
        // We rebuild a single-input changegroup just to stream-out.
        let mut group = Changegroup::new()?;
        group.add(cs)?;
        let buf_ref: &mut &mut dyn Write = &mut { buf_dyn };
        group.output_strm(*buf_ref)?;
        buf.flush()?;
    }
    std::fs::rename(&tmp, path)?;
    Ok(())
}

fn atomic_write(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    let tmp = path.with_extension("lpatch.tmp");
    {
        let mut f = File::create(&tmp)?;
        f.write_all(bytes)?;
        f.sync_all()?;
    }
    std::fs::rename(&tmp, path)
}

/// Public entry point for rendering a changeset to any `Write` —
/// used both by `vlotus patch show` (which writes to stdout) and by
/// `:patch show` in the TUI (which captures to a Vec for popup
/// display).
pub fn render_to(bytes: &[u8], out: &mut dyn Write) -> Result<(), StoreError> {
    render_changeset(bytes, out)
}

fn render_changeset(bytes: &[u8], out: &mut dyn Write) -> Result<(), StoreError> {
    let mut input: &[u8] = bytes;
    let input_dyn: &mut dyn Read = &mut input;
    let input_borrow: &&mut dyn Read = &input_dyn;
    let mut iter = ChangesetIter::start_strm(input_borrow)?;
    while let Some(item) = iter.next()? {
        let op = item.op()?;
        let table = op.table_name();
        let kind = match op.code() {
            Action::SQLITE_INSERT => "INSERT",
            Action::SQLITE_UPDATE => "UPDATE",
            Action::SQLITE_DELETE => "DELETE",
            _ => "?",
        };
        let n_cols = op.number_of_columns() as usize;
        let pk_mask = item.pk()?.to_vec();

        // Read every column. INSERT has new only; DELETE has old
        // only; UPDATE has both with PK columns mirrored on each side
        // (SQLite stores PK in old+new for updates).
        let new_vals: Vec<Option<ValueDescription>> = (0..n_cols)
            .map(|c| {
                if op.code() == Action::SQLITE_DELETE {
                    None
                } else {
                    item.new_value(c).ok().map(value_repr)
                }
            })
            .collect();
        let old_vals: Vec<Option<ValueDescription>> = (0..n_cols)
            .map(|c| {
                if op.code() == Action::SQLITE_INSERT {
                    None
                } else {
                    item.old_value(c).ok().map(value_repr)
                }
            })
            .collect();

        match table {
            "cell" | "cell_format" => {
                // PK = (sheet_name, row, col); cell has raw at idx 3,
                // computed at 4, owner_row at 5, owner_col at 6.
                // cell_format has format_json at idx 3.
                let sheet = pk_string(0, &new_vals, &old_vals);
                let row = pk_u64(1, &new_vals, &old_vals);
                let col = pk_u64(2, &new_vals, &old_vals);
                let coords = match (row, col) {
                    (Some(r), Some(c)) => format!(
                        "{sheet}!{}{}",
                        col_idx_to_letters(c as u32),
                        r + 1
                    ),
                    _ => format!("{sheet}!?"),
                };
                // Both `cell.raw` and `cell_format.format_json` sit
                // at column index 3 (after the composite PK columns).
                let value_col = 3;
                let new_v = new_vals
                    .get(value_col)
                    .and_then(|v| v.as_ref())
                    .map(|v| v.to_string());
                let old_v = old_vals
                    .get(value_col)
                    .and_then(|v| v.as_ref())
                    .map(|v| v.to_string());
                let suffix = if table == "cell_format" { " [format]" } else { "" };
                match (kind, old_v, new_v) {
                    ("INSERT", _, Some(n)) => writeln!(out, "{coords}{suffix}: + {n}")?,
                    ("DELETE", Some(o), _) => writeln!(out, "{coords}{suffix}: - {o}")?,
                    ("UPDATE", Some(o), Some(n)) => {
                        writeln!(out, "{coords}{suffix}: {o} → {n}")?
                    }
                    ("UPDATE", None, Some(n)) => {
                        writeln!(out, "{coords}{suffix}: → {n}")?
                    }
                    ("UPDATE", Some(o), None) => {
                        writeln!(out, "{coords}{suffix}: {o} →")?
                    }
                    _ => writeln!(out, "{coords}{suffix}: ({kind})")?,
                }
            }
            "sheet" => {
                let name = pk_string(0, &new_vals, &old_vals);
                writeln!(out, "{kind} sheet: {name}")?;
            }
            "column_meta" => {
                let sheet = pk_string(0, &new_vals, &old_vals);
                let col = pk_u64(1, &new_vals, &old_vals).unwrap_or(0);
                let letters = col_idx_to_letters(col as u32);
                let new_w = new_vals
                    .get(2)
                    .and_then(|v| v.as_ref())
                    .map(|v| v.to_string());
                let old_w = old_vals
                    .get(2)
                    .and_then(|v| v.as_ref())
                    .map(|v| v.to_string());
                match (kind, old_w, new_w) {
                    ("INSERT", _, Some(n)) => {
                        writeln!(out, "{sheet}!{letters} [width]: + {n}")?
                    }
                    ("DELETE", Some(o), _) => {
                        writeln!(out, "{sheet}!{letters} [width]: - {o}")?
                    }
                    ("UPDATE", Some(o), Some(n)) => {
                        writeln!(out, "{sheet}!{letters} [width]: {o} → {n}")?
                    }
                    _ => writeln!(out, "{sheet}!{letters} [width]: ({kind})")?,
                }
            }
            other => writeln!(out, "{kind} {other}")?,
        }
        let _ = pk_mask; // silence unused warning if PK introspection isn't used
    }
    Ok(())
}

#[derive(Debug, Clone)]
enum ValueDescription {
    Null,
    Int(i64),
    Real(f64),
    Text(String),
    Blob(usize),
}

impl std::fmt::Display for ValueDescription {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ValueDescription::Null => write!(f, "NULL"),
            ValueDescription::Int(n) => write!(f, "{n}"),
            ValueDescription::Real(x) => write!(f, "{x}"),
            ValueDescription::Text(s) => write!(f, "{s:?}"),
            ValueDescription::Blob(n) => write!(f, "<blob {n}B>"),
        }
    }
}

fn value_repr(v: ValueRef<'_>) -> ValueDescription {
    match v {
        ValueRef::Null => ValueDescription::Null,
        ValueRef::Integer(n) => ValueDescription::Int(n),
        ValueRef::Real(x) => ValueDescription::Real(x),
        ValueRef::Text(t) => ValueDescription::Text(
            String::from_utf8_lossy(t).into_owned(),
        ),
        ValueRef::Blob(b) => ValueDescription::Blob(b.len()),
    }
}

fn pk_string(
    idx: usize,
    new_vals: &[Option<ValueDescription>],
    old_vals: &[Option<ValueDescription>],
) -> String {
    let pick = new_vals
        .get(idx)
        .and_then(|v| v.as_ref())
        .or_else(|| old_vals.get(idx).and_then(|v| v.as_ref()));
    match pick {
        Some(ValueDescription::Text(s)) => s.clone(),
        Some(other) => other.to_string(),
        None => "?".into(),
    }
}

fn pk_u64(
    idx: usize,
    new_vals: &[Option<ValueDescription>],
    old_vals: &[Option<ValueDescription>],
) -> Option<u64> {
    let pick = new_vals
        .get(idx)
        .and_then(|v| v.as_ref())
        .or_else(|| old_vals.get(idx).and_then(|v| v.as_ref()));
    match pick {
        Some(ValueDescription::Int(n)) if *n >= 0 => Some(*n as u64),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::store::CellChange;

    fn fresh_store() -> Store {
        let mut s = Store::open_in_memory().unwrap();
        s.create_sheet("S").unwrap();
        s.commit().unwrap();
        s
    }

    fn ch(r: u32, c: u32, raw: &str) -> CellChange {
        CellChange {
            row_idx: r,
            col_idx: c,
            raw_value: raw.into(),
            format_json: None,
        }
    }

    fn unique_tmp_dir() -> std::path::PathBuf {
        // Tests can race on the nanosecond clock; salt with PID +
        // counter so parallel runs don't collide.
        use std::sync::atomic::{AtomicUsize, Ordering};
        static COUNTER: AtomicUsize = AtomicUsize::new(0);
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let n = COUNTER.fetch_add(1, Ordering::Relaxed);
        let pid = std::process::id();
        let dir = std::env::temp_dir().join(format!("vlotus-cli-{pid}-{nanos}-{n}"));
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    fn record_simple_patch(path: &Path) -> Vec<u8> {
        let mut store = fresh_store();
        store.patch_open(path.to_path_buf()).unwrap();
        store.apply("S", &[ch(0, 0, "5"), ch(0, 1, "=A1*2")]).unwrap();
        store.patch_close().unwrap();
        std::fs::read(path).unwrap()
    }

    #[test]
    fn show_renders_a1_style_lines() {
        let dir = unique_tmp_dir();
        let path = dir.join("p.lpatch");
        let bytes = record_simple_patch(&path);

        let mut buf: Vec<u8> = Vec::new();
        render_changeset(&bytes, &mut buf).unwrap();
        let out = String::from_utf8(buf).unwrap();
        assert!(out.contains("S!A1: + \"5\""), "got: {out}");
        assert!(out.contains("S!B1: + \"=A1*2\""), "got: {out}");
    }

    #[test]
    fn invert_round_trip_via_bytes() {
        let dir = unique_tmp_dir();
        let path = dir.join("p.lpatch");
        let original = record_simple_patch(&path);
        let inverse = invert_bytes(&original).unwrap();
        // Apply original then inverse on a fresh store; expect empty.
        let mut s = fresh_store();
        s.apply_changeset(&original, ConflictPolicy::Omit).unwrap();
        s.apply_changeset(&inverse, ConflictPolicy::Omit).unwrap();
        s.recalculate("S").unwrap();
        let snap = s.load_sheet("S").unwrap();
        assert!(
            snap.cells.is_empty(),
            "after apply+invert the sheet should be empty: {:?}",
            snap.cells
        );
    }

    #[test]
    fn diff_derives_patch_between_two_dbs() {
        let dir = unique_tmp_dir();
        let from_path = dir.join("from.db");
        let to_path = dir.join("to.db");
        let out_path = dir.join("d.lpatch");

        // Both DBs start with Sheet1 + A1=1.
        for path in [&from_path, &to_path] {
            let mut s = Store::open(path).unwrap();
            s.create_sheet("S").unwrap();
            s.apply("S", &[ch(0, 0, "1")]).unwrap();
            s.commit().unwrap();
            drop(s);
        }
        // `to` then mutates: A1 = 100.
        {
            let mut s = Store::open(&to_path).unwrap();
            s.apply("S", &[ch(0, 0, "100")]).unwrap();
            s.commit().unwrap();
        }

        let rc = run_diff(
            from_path.to_str().unwrap(),
            to_path.to_str().unwrap(),
            out_path.to_str().unwrap(),
        );
        assert_eq!(rc, 0);

        // Apply diff on a clone of `from` and assert it now matches
        // `to`.
        let clone_path = dir.join("clone.db");
        std::fs::copy(&from_path, &clone_path).unwrap();
        let bytes = std::fs::read(&out_path).unwrap();
        {
            let mut s = Store::open(&clone_path).unwrap();
            s.apply_changeset(&bytes, ConflictPolicy::Omit).unwrap();
            s.recalculate("S").unwrap();
            s.commit().unwrap();
            let snap = s.load_sheet("S").unwrap();
            let a1 = snap.cells.iter().find(|c| (c.row_idx, c.col_idx) == (0, 0));
            assert_eq!(a1.unwrap().raw_value, "100");
        }
    }

    #[test]
    fn combine_concatenates_in_order() {
        let dir = unique_tmp_dir();
        // Patch A: write A1=1.
        let a_path = dir.join("a.lpatch");
        {
            let mut s = fresh_store();
            s.patch_open(a_path.clone()).unwrap();
            s.apply("S", &[ch(0, 0, "1")]).unwrap();
            s.patch_close().unwrap();
        }
        // Patch B: against the post-A state, write A1=2.
        let b_path = dir.join("b.lpatch");
        {
            let mut s = fresh_store();
            s.apply("S", &[ch(0, 0, "1")]).unwrap();
            s.commit().unwrap();
            s.patch_open(b_path.clone()).unwrap();
            s.apply("S", &[ch(0, 0, "2")]).unwrap();
            s.patch_close().unwrap();
        }

        let combined = dir.join("ab.lpatch");
        let rc = run_combine(
            combined.to_str().unwrap(),
            &[
                a_path.to_str().unwrap().to_string(),
                b_path.to_str().unwrap().to_string(),
            ],
        );
        assert_eq!(rc, 0);

        // Apply the combined patch on a fresh store; A1 should land
        // at 2 (B's UPDATE composed onto A's INSERT).
        let mut s = fresh_store();
        let bytes = std::fs::read(&combined).unwrap();
        s.apply_changeset(&bytes, ConflictPolicy::Omit).unwrap();
        s.recalculate("S").unwrap();
        let snap = s.load_sheet("S").unwrap();
        let a1 = snap.cells.iter().find(|c| (c.row_idx, c.col_idx) == (0, 0));
        assert_eq!(a1.unwrap().raw_value, "2");
    }
}
