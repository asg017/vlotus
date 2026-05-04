//! Portable diff files via SQLite's session extension. P1 of the
//! patch epic (att `coisl46s`).
//!
//! A `Patch` records every authored mutation between `:patch new` and
//! `:patch close` as an SQLite changeset. The changeset is written to
//! a `.lpatch` file via `Session::changeset_strm`. Other workbooks
//! with the same schema can replay the patch with rusqlite's
//! `Connection::apply` — composite-PK identification means cells are
//! addressed by `(sheet_name, row, col)` so patches port across DBs
//! that share names rather than synthetic IDs.
//!
//! The session sits inside `Store` next to the `Connection`. Sessions
//! borrow the Connection in rusqlite's API surface, but at runtime
//! they hold a raw `*mut sqlite3_session` independent of the Rust
//! borrow — the C side maintains the relationship. Storing a
//! `Session<'static>` is sound iff we drop it before the Connection;
//! [`Store`]'s Drop guarantees this by ordering `patch` first and
//! reassigning `self.patch = None` ahead of any conn cleanup.

use std::fs::File;
use std::io::Write;
use std::path::PathBuf;

use rusqlite::session::{ConflictAction, Session};

use super::{Store, StoreError};

/// User-data tables that are recorded into a patch. Anything else
/// (`undo_entry`, `meta`) is filtered out so the patch carries only
/// authored state.
const PATCH_TABLES: &[&str] = &["sheet", "column_meta", "cell", "cell_format"];

#[derive(Debug, Clone, Copy)]
pub enum PatchSaveMode {
    /// Default: write the changeset to the patch file, keep the dirty
    /// buffer as-is so the user keeps editing.
    KeepDirty,
    /// Write the changeset and `:w` the workbook in the same step.
    Commit,
    /// Write the changeset and roll back the workbook ("shelve" the
    /// edits — useful for "save what I tried, revert the workbook").
    Rollback,
}

/// Snapshot of an active patch — what `:patch status` renders.
#[derive(Debug, Clone)]
pub struct PatchStatus {
    pub path: PathBuf,
    pub paused: bool,
}

pub(super) struct PatchState {
    pub path: PathBuf,
    /// `'static` is a lie — the session's runtime state is the raw C
    /// pointer, not the Rust borrow. `Store::drop` releases this
    /// before dropping `conn`, which is what makes the cast sound.
    pub session: Session<'static>,
    pub paused: bool,
}

impl Store {
    /// Open a new patch. Records every mutation to user-data tables
    /// (cell / cell_format / sheet / column_meta) until `patch_close`
    /// or `patch_detach`. `undo_entry` and `meta` are excluded so
    /// the patch carries only authored state.
    pub fn patch_open(&mut self, path: PathBuf) -> Result<(), StoreError> {
        if let Some(existing) = &self.patch {
            return Err(StoreError::PatchAlreadyOpen {
                path: existing.path.clone(),
            });
        }

        let mut session = Session::new(&self.conn)?;
        session.table_filter(Some(|tbl: &str| PATCH_TABLES.contains(&tbl)));
        session.attach(None)?;

        // Extend the session's lifetime to 'static for storage.
        // SAFETY: rusqlite's `Session<'conn>` is a wrapper around a
        // raw `*mut sqlite3_session` plus a phantom borrow of the
        // Connection. The runtime relationship between the session
        // and the connection is maintained on the C side — the Rust
        // lifetime is conservative bookkeeping. We extend to 'static
        // so the session can sit inside `Store` next to the
        // `Connection`. Soundness depends on:
        //   1. `Store` declares `patch: Option<PatchState>` ahead of
        //      `conn`, so compiler-generated drop order is patch
        //      first if our custom Drop body is bypassed.
        //   2. `Store::drop` explicitly takes `self.patch = None` as
        //      its first action, dropping the session (which calls
        //      sqlite3session_delete) before any conn cleanup.
        //   3. `PatchState` is `pub(super)` and never escapes this
        //      crate, so callers can't transmute or move the
        //      session out from under the Connection.
        #[allow(unsafe_code)]
        let session: Session<'static> = unsafe { std::mem::transmute(session) };

        self.patch = Some(PatchState {
            path,
            session,
            paused: false,
        });
        Ok(())
    }

    /// Write the accumulated changeset to disk and act on the buffer
    /// per `mode`. Idempotent w.r.t. the session — `KeepDirty` leaves
    /// the session in place so further edits keep accumulating.
    pub fn patch_save(&mut self, mode: PatchSaveMode) -> Result<PathBuf, StoreError> {
        let saved_path = {
            let patch = self.patch.as_mut().ok_or(StoreError::NoActivePatch)?;
            // tmp + rename so a crash mid-write doesn't truncate the
            // user's patch file.
            let tmp_path = patch.path.with_extension("lpatch.tmp");
            {
                let mut file = File::create(&tmp_path)?;
                let writer: &mut dyn Write = &mut file;
                patch.session.changeset_strm(writer)?;
                file.sync_all()?;
            }
            std::fs::rename(&tmp_path, &patch.path)?;
            patch.path.clone()
        };

        match mode {
            PatchSaveMode::KeepDirty => {}
            PatchSaveMode::Commit => self.commit()?,
            PatchSaveMode::Rollback => self.rollback()?,
        }
        Ok(saved_path)
    }

    /// Save (in `KeepDirty` mode) and stop recording. The session is
    /// dropped; subsequent edits do not land in any patch.
    pub fn patch_close(&mut self) -> Result<PathBuf, StoreError> {
        let path = self.patch_save(PatchSaveMode::KeepDirty)?;
        self.patch = None;
        Ok(path)
    }

    /// Stop recording WITHOUT saving. Discards the in-memory
    /// changeset; the workbook itself is unaffected.
    pub fn patch_detach(&mut self) -> Result<(), StoreError> {
        if self.patch.is_none() {
            return Err(StoreError::NoActivePatch);
        }
        self.patch = None;
        Ok(())
    }

    /// Apply the inverse of the recorded changeset to the live
    /// workbook — effectively rolls the workbook back to the state
    /// at `patch_open` time. The patch session itself is reset (the
    /// invert lands as new edits, but those are bracketed out so the
    /// changeset returns to empty).
    pub fn patch_invert(&mut self) -> Result<(), StoreError> {
        let inverse = {
            let patch = self.patch.as_mut().ok_or(StoreError::NoActivePatch)?;
            patch.session.changeset()?.invert()?
        };
        self.with_session_disabled(|store| {
            store.conn.apply(
                &inverse,
                None::<fn(&str) -> bool>,
                |_ctype, _item| ConflictAction::SQLITE_CHANGESET_OMIT,
            )
        })?;
        // Recalc every sheet so derived `computed` / `owner_*` columns
        // reflect the rolled-back authored state.
        let sheets: Vec<String> = self
            .list_sheets()?
            .into_iter()
            .map(|s| s.name)
            .collect();
        for name in sheets {
            self.recalculate(&name)?;
        }
        // Re-create the session so subsequent edits start fresh.
        let path = self
            .patch
            .as_ref()
            .map(|p| p.path.clone())
            .ok_or(StoreError::NoActivePatch)?;
        let paused = self.patch.as_ref().map(|p| p.paused).unwrap_or(false);
        self.patch = None;
        self.patch_open(path)?;
        if paused {
            self.patch_set_enabled(false);
        }
        self.mark_dirty();
        Ok(())
    }

    /// Pause / resume recording. While paused, edits land in the
    /// workbook but not in the changeset. Used by `:patch pause` /
    /// `:patch resume`.
    pub fn patch_set_enabled(&mut self, on: bool) {
        if let Some(patch) = self.patch.as_mut() {
            patch.session.set_enabled(on);
            patch.paused = !on;
        }
    }

    pub fn patch_status(&self) -> Option<PatchStatus> {
        self.patch.as_ref().map(|p| PatchStatus {
            path: p.path.clone(),
            paused: p.paused,
        })
    }

    /// Run `f` with the patch session temporarily disabled. Used to
    /// bracket recalc, undo-log writes, and patch-invert apply so
    /// derived/internal mutations don't bloat the patch.
    pub(super) fn with_session_disabled<F, R>(&mut self, f: F) -> R
    where
        F: FnOnce(&mut Self) -> R,
    {
        let was_enabled = self
            .patch
            .as_mut()
            .map(|p| {
                let prior = !p.paused;
                p.session.set_enabled(false);
                prior
            })
            .unwrap_or(false);
        let result = f(self);
        if was_enabled {
            if let Some(p) = self.patch.as_mut() {
                p.session.set_enabled(true);
            }
        }
        result
    }

    /// Drop the active patch session without saving. Called from
    /// `Store::rollback` because `:q!` discards the dirty buffer
    /// the session was recording against — the changeset would
    /// diverge from reality otherwise.
    pub(super) fn invalidate_patch_on_rollback(&mut self) -> bool {
        if self.patch.is_some() {
            self.patch = None;
            true
        } else {
            false
        }
    }
}

/// Conflict-handler policy for `Store::apply_changeset`. Mirrors
/// SQLite's `SQLITE_CHANGESET_OMIT` / `_REPLACE` / `_ABORT` (the
/// last halts the apply and rolls the partial work back).
#[derive(Debug, Clone, Copy)]
pub enum ConflictPolicy {
    Omit,
    Replace,
    Abort,
}

impl Store {
    /// Apply a `.lpatch` byte stream to the workbook. Recalc is the
    /// caller's responsibility — typically `Store::recalculate` per
    /// affected sheet, then `Store::commit`.
    ///
    /// Conflicts are routed through the configured policy and emitted
    /// to stderr (`<table>:<pk> CONFLICT (kind): <action>`) so the
    /// user isn't blind. Returns the number of conflicts encountered.
    pub fn apply_changeset(
        &mut self,
        bytes: &[u8],
        policy: ConflictPolicy,
    ) -> Result<usize, StoreError> {
        use rusqlite::session::{ConflictAction, ConflictType};
        use std::sync::atomic::{AtomicUsize, Ordering};

        let counter = std::sync::Arc::new(AtomicUsize::new(0));
        let counter_for_handler = std::sync::Arc::clone(&counter);

        let mut input: &[u8] = bytes;
        self.mark_dirty();
        self.with_session_disabled(|s| {
            let conn = &s.conn;
            let mut input_dyn: &mut dyn std::io::Read = &mut input;
            conn.apply_strm(
                &mut input_dyn,
                None::<fn(&str) -> bool>,
                move |ctype: ConflictType, item| {
                    counter_for_handler.fetch_add(1, Ordering::Relaxed);
                    let table = item.op().map(|op| op.table_name().to_string()).unwrap_or_default();
                    let kind_str = format!("{ctype:?}").replace("SQLITE_CHANGESET_", "");
                    let action = match policy {
                        ConflictPolicy::Omit => ConflictAction::SQLITE_CHANGESET_OMIT,
                        ConflictPolicy::Replace => match ctype {
                            // REPLACE is only legal for DATA / CONFLICT
                            // per the docs — fall back to OMIT for
                            // CONSTRAINT / NOTFOUND / FOREIGN_KEY.
                            ConflictType::SQLITE_CHANGESET_DATA
                            | ConflictType::SQLITE_CHANGESET_CONFLICT => {
                                ConflictAction::SQLITE_CHANGESET_REPLACE
                            }
                            _ => ConflictAction::SQLITE_CHANGESET_OMIT,
                        },
                        ConflictPolicy::Abort => ConflictAction::SQLITE_CHANGESET_ABORT,
                    };
                    let action_str = format!("{action:?}").replace("SQLITE_CHANGESET_", "");
                    eprintln!("{table}: CONFLICT ({kind_str}) → {action_str}");
                    action
                },
            )
        })?;
        Ok(counter.load(Ordering::Relaxed))
    }
}

/// Read a `.lpatch` file into bytes. Public for `patch_cli`.
pub fn read_patch_file(path: &std::path::Path) -> std::io::Result<Vec<u8>> {
    std::fs::read(path)
}

#[cfg(test)]
pub(super) fn apply_changeset_bytes(store: &mut Store, bytes: &[u8]) -> Result<usize, StoreError> {
    store.apply_changeset(bytes, ConflictPolicy::Omit)
}

#[cfg(test)]
mod tests {
    use super::super::CellChange;
    use super::*;

    fn fresh(name: &str) -> Store {
        let mut store = Store::open_in_memory().unwrap();
        store.create_sheet(name).unwrap();
        store.commit().unwrap();
        store
    }

    fn change(row: u32, col: u32, raw: &str) -> CellChange {
        CellChange {
            row_idx: row,
            col_idx: col,
            raw_value: raw.into(),
            format_json: None,
        }
    }

    #[test]
    fn patch_records_authored_changes_only() {
        let dir = unique_tmp_dir();
        let patch_path = dir.join("rec.lpatch");

        let mut a = fresh("S");
        a.patch_open(patch_path.clone()).unwrap();
        a.apply("S", &[change(0, 0, "5"), change(0, 1, "=A1*2")])
            .unwrap();
        a.patch_close().unwrap();
        let bytes = read_patch_file(&patch_path).unwrap();
        assert!(!bytes.is_empty(), "patch file populated");

        // Apply on a fresh store and assert the user-authored cells
        // round-trip. computed values regenerate via recalc.
        let mut b = fresh("S");
        apply_changeset_bytes(&mut b, &bytes).unwrap();
        b.recalculate("S").unwrap();
        let snap = b.load_sheet("S").unwrap();
        let a1 = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (0, 0))
            .unwrap();
        let b1 = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (0, 1))
            .unwrap();
        assert_eq!(a1.raw_value, "5");
        assert_eq!(b1.raw_value, "=A1*2");
        assert_eq!(b1.computed_value.as_deref(), Some("10"));
    }

    #[test]
    fn patch_excludes_undo_entry() {
        let dir = unique_tmp_dir();
        let patch_path = dir.join("p.lpatch");

        let mut a = fresh("S");
        a.patch_open(patch_path.clone()).unwrap();
        // Apply a change AND record an undo group — both kinds of
        // writes touch the conn, but only the cell change should
        // appear in the patch.
        let prior = a.snapshot_cell("S", 0, 0).unwrap();
        a.apply("S", &[change(0, 0, "x")]).unwrap();
        a.record_undo_group(&[prior]).unwrap();
        a.patch_close().unwrap();

        let bytes = read_patch_file(&patch_path).unwrap();
        // Apply on a fresh store and verify undo_entry is empty
        // (filter prevented those rows from being recorded).
        let mut b = fresh("S");
        apply_changeset_bytes(&mut b, &bytes).unwrap();
        let undo_count: i64 = b
            .conn()
            .query_row("SELECT COUNT(*) FROM undo_entry", [], |r| r.get(0))
            .unwrap();
        assert_eq!(undo_count, 0);
    }

    #[test]
    fn patch_invert_round_trips_authored_state() {
        let dir = unique_tmp_dir();
        let patch_path = dir.join("inv.lpatch");

        let mut store = fresh("S");
        // Establish baseline.
        store.apply("S", &[change(0, 0, "baseline")]).unwrap();
        store.commit().unwrap();

        store.patch_open(patch_path).unwrap();
        store.apply("S", &[change(0, 0, "edited"), change(1, 1, "new")])
            .unwrap();
        store.patch_invert().unwrap();
        let snap = store.load_sheet("S").unwrap();
        let a1 = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (0, 0));
        assert_eq!(a1.unwrap().raw_value, "baseline");
        // The (1,1) cell never existed in the baseline — invert
        // should remove it.
        assert!(!snap
            .cells
            .iter()
            .any(|c| (c.row_idx, c.col_idx) == (1, 1)));
    }

    #[test]
    fn patch_pause_resume_brackets_edits_out() {
        let dir = unique_tmp_dir();
        let patch_path = dir.join("p.lpatch");

        let mut a = fresh("S");
        a.patch_open(patch_path.clone()).unwrap();
        a.apply("S", &[change(0, 0, "kept")]).unwrap();
        a.patch_set_enabled(false);
        a.apply("S", &[change(1, 0, "skipped")]).unwrap();
        a.patch_set_enabled(true);
        a.apply("S", &[change(2, 0, "kept-again")]).unwrap();
        a.patch_close().unwrap();
        let bytes = read_patch_file(&patch_path).unwrap();

        let mut b = fresh("S");
        apply_changeset_bytes(&mut b, &bytes).unwrap();
        let snap = b.load_sheet("S").unwrap();
        // Only "kept" and "kept-again" should make it through.
        let raws: Vec<&str> = snap.cells.iter().map(|c| c.raw_value.as_str()).collect();
        assert!(raws.contains(&"kept"));
        assert!(raws.contains(&"kept-again"));
        assert!(
            !raws.contains(&"skipped"),
            "paused edits leaked into patch: {raws:?}"
        );
    }

    #[test]
    fn rollback_invalidates_active_patch() {
        let dir = unique_tmp_dir();
        let mut store = fresh("S");
        store.patch_open(dir.join("p.lpatch")).unwrap();
        assert!(store.patch_status().is_some());
        store.apply("S", &[change(0, 0, "x")]).unwrap();
        store.rollback().unwrap();
        assert!(
            store.patch_status().is_none(),
            "patch should be dropped on rollback"
        );
    }

    #[test]
    fn patch_open_twice_errors() {
        let dir = unique_tmp_dir();
        let mut store = fresh("S");
        store.patch_open(dir.join("a.lpatch")).unwrap();
        match store.patch_open(dir.join("b.lpatch")) {
            Err(StoreError::PatchAlreadyOpen { .. }) => {}
            other => panic!("expected PatchAlreadyOpen, got {other:?}"),
        }
    }

    fn unique_tmp_dir() -> PathBuf {
        use std::sync::atomic::{AtomicUsize, Ordering};
        static COUNTER: AtomicUsize = AtomicUsize::new(0);
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let n = COUNTER.fetch_add(1, Ordering::Relaxed);
        let pid = std::process::id();
        let dir = std::env::temp_dir().join(format!("vlotus-patch-{pid}-{nanos}-{n}"));
        std::fs::create_dir_all(&dir).unwrap();
        dir
    }

    #[test]
    fn patch_status_none_when_inactive() {
        let store = fresh("S");
        assert!(store.patch_status().is_none());
    }
}
