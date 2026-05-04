//! SQLite-backed workbook storage for vlotus.
//!
//! SQLite-backed workbook storage.
//!
//! # Lifecycle
//!
//! - [`Store::open`] runs schema setup, then opens the long-lived
//!   dirty-buffer txn with `BEGIN IMMEDIATE`. Every mutation lands
//!   inside that txn until [`Store::commit`] flushes it.
//! - [`Store::commit`] is `:w` — `COMMIT; BEGIN IMMEDIATE;` so the
//!   next edit is buffered again.
//! - [`Store::rollback`] is `:q!` — `ROLLBACK; BEGIN IMMEDIATE;`. The
//!   buffer is cleared but the store stays usable; the App's quit
//!   path drops the store after.
//! - The Drop guard rolls back any pending buffer best-effort. SQLite's
//!   WAL also discards an uncommitted txn on next open if the process
//!   crashes before Drop fires.
//! - [`Store::is_dirty`] mirrors "has anything changed since the last
//!   commit?" for the `:q` guard.

use std::path::{Path, PathBuf};

use rusqlite::Connection;

pub mod cells;
pub mod columns;
pub mod coords;
pub mod errors;
pub mod patch;
pub mod schema;
pub mod sheets;
pub mod structural;
pub mod undo;

pub use cells::{CellChange, StoredCell};
pub use columns::ColumnMeta;
pub use errors::StoreError;
pub use patch::{ConflictPolicy, PatchSaveMode};
pub use schema::{APPLICATION_ID, SCHEMA_VERSION};
pub use sheets::SheetMeta;
pub use undo::{UndoGroup, UndoOp};

pub struct Store {
    /// Active patch session. **Must** appear before `conn` in field
    /// order — the session holds a `'static` lifetime over the
    /// connection that is only sound because Drop releases the
    /// session before the connection. Compiler-generated drop order
    /// follows declaration order; our explicit `Drop` impl belt-
    /// and-suspenders this with `self.patch = None;` first.
    patch: Option<patch::PatchState>,
    conn: Connection,
    /// Path the connection was opened from, kept around for error
    /// reporting. `None` for in-memory stores.
    path: Option<PathBuf>,
    dirty: bool,
    /// In-memory cache of `(sheet_name, row, col) -> CustomCell` for
    /// cells whose computed value is a `CellValue::Custom`. Repopulated
    /// on every `recalculate`; ephemeral by design (no on-disk schema
    /// for it). Used by `App::datetime_tag_for_cell` to drive
    /// auto-styling and by `App::displayed_for` to apply a per-cell
    /// strftime pattern (`:fmt date`) without re-parsing.
    custom_cells: std::collections::HashMap<(String, u32, u32), CustomCell>,
}

/// In-memory snapshot of a `CellValue::Custom` keyed alongside the
/// stored `cell.computed`. Holds enough state for vlotus to drive
/// auto-style + per-cell strftime without reaching back into the
/// engine.
#[derive(Debug, Clone)]
#[allow(dead_code)] // fields read only under `feature = "datetime"`
pub struct CustomCell {
    pub type_tag: String,
    /// The handler's canonical data string (e.g. `2025-04-27` for
    /// jdate, `P1Y2M3D` for jspan). Distinct from `cell.computed`,
    /// which holds the friendly display form.
    pub data: String,
}

impl std::fmt::Debug for Store {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Store")
            .field("path", &self.path)
            .field("dirty", &self.dirty)
            .field("patch_active", &self.patch.is_some())
            .finish()
    }
}

impl Store {
    pub fn open(path: &Path) -> Result<Self, StoreError> {
        let conn = Connection::open(path)?;
        Self::initialise(conn, Some(path.to_path_buf()), /* in_memory */ false)
    }

    pub fn open_in_memory() -> Result<Self, StoreError> {
        let conn = Connection::open_in_memory()?;
        Self::initialise(conn, None, /* in_memory */ true)
    }

    fn initialise(
        conn: Connection,
        path: Option<PathBuf>,
        in_memory: bool,
    ) -> Result<Self, StoreError> {
        if !in_memory {
            // WAL is per-file; in-memory DBs silently downgrade to memory
            // mode but the call still succeeds. We skip it so `:memory:`
            // tests don't depend on platform shared-memory primitives.
            conn.execute_batch("PRAGMA journal_mode = WAL;")?;
        }
        conn.execute_batch("PRAGMA foreign_keys = ON;")?;

        let app_id = schema::read_application_id(&conn)?;
        let user_version = schema::read_user_version(&conn)?;

        match (app_id, user_version) {
            (APPLICATION_ID, SCHEMA_VERSION) => {
                // Existing vlotus DB at the current version. Run DDL
                // idempotently so a partially-written predecessor doesn't
                // strand us.
                schema::ensure_schema(&conn)?;
            }
            (APPLICATION_ID, v) if v > SCHEMA_VERSION => {
                return Err(StoreError::UnsupportedSchemaVersion {
                    path: path.unwrap_or_default(),
                    user_version: v,
                    supported: SCHEMA_VERSION,
                });
            }
            (0, 0) if !schema::has_any_user_table(&conn)? => {
                // Fresh / empty file. Stamp it and write our schema.
                schema::ensure_schema(&conn)?;
                schema::write_application_id(&conn, APPLICATION_ID)?;
                schema::write_user_version(&conn, SCHEMA_VERSION)?;
            }
            (other, _) if other != APPLICATION_ID && schema::has_legacy_datasette_table(&conn)? => {
                return Err(StoreError::LegacyDatabase {
                    path: path.unwrap_or_default(),
                });
            }
            (a, v) => {
                return Err(StoreError::UnknownDatabase {
                    path: path.unwrap_or_default(),
                    app_id: a,
                    user_version: v,
                });
            }
        }

        // Open the dirty-buffer txn. Every mutation lands inside this
        // txn until commit() / rollback() / Drop.
        conn.execute_batch("BEGIN IMMEDIATE;")?;

        Ok(Store {
            patch: None,
            conn,
            path,
            dirty: false,
            custom_cells: std::collections::HashMap::new(),
        })
    }

    /// Snapshot of the custom value at `(sheet, row, col)`, when the
    /// cell resolved to a `CellValue::Custom`. Populated during
    /// `recalculate`; returns `None` for built-in types and for cells
    /// whose sheet hasn't been recalculated since open.
    #[cfg(feature = "datetime")]
    pub fn cell_custom(&self, sheet: &str, row: u32, col: u32) -> Option<&CustomCell> {
        self.custom_cells.get(&(sheet.to_string(), row, col))
    }

    /// True when the dirty buffer has uncommitted writes. `:q` checks
    /// this to refuse a quit; `:w` resets it.
    pub fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Path the connection was opened from; `None` for in-memory.
    #[allow(dead_code)]
    pub fn path(&self) -> Option<&Path> {
        self.path.as_deref()
    }

    /// `:w` — flush the dirty-buffer txn to disk and open a fresh one
    /// so the next edit is buffered again. The undo log is cleared
    /// inside the same txn so a freshly-committed file has no
    /// dangling undo state from before the save.
    pub fn commit(&mut self) -> Result<(), StoreError> {
        self.clear_undo_log()?;
        self.conn.execute_batch("COMMIT; BEGIN IMMEDIATE;")?;
        self.dirty = false;
        Ok(())
    }

    /// `:q!` — discard everything since the last commit, then re-open
    /// the buffer so the store stays usable. Caller's responsibility
    /// to drop the store afterwards if the user is quitting.
    ///
    /// If a patch was being recorded, it's invalidated: the
    /// rolled-back edits would leave the changeset diverged from
    /// reality. Caller can detect this via `patch_status()` returning
    /// `None` after the call and surface a status message.
    pub fn rollback(&mut self) -> Result<(), StoreError> {
        // Drop the patch session BEFORE issuing ROLLBACK so the
        // session releases its conn-side state cleanly.
        let _patch_was_active = self.invalidate_patch_on_rollback();
        self.conn.execute_batch("ROLLBACK; BEGIN IMMEDIATE;")?;
        self.dirty = false;
        Ok(())
    }

    /// Read-only access to the underlying connection. T2-T5 use this to
    /// run cell / format / undo SQL until each one earns its own
    /// dedicated `Store` method.
    pub(crate) fn conn(&self) -> &Connection {
        &self.conn
    }

    /// Flip the dirty flag on. Submodules (`cells`, `sheets`, …) call
    /// this after every mutation so `is_dirty` reflects pending changes
    /// for T6's `:q` guard.
    pub(in crate::store) fn mark_dirty(&mut self) {
        self.dirty = true;
    }
}

impl Drop for Store {
    fn drop(&mut self) {
        // Drop the patch session FIRST. The 'static lifetime in
        // PatchState's Session is a fiction maintained by this
        // ordering: session must be released before the conn it
        // borrows from on the C side.
        self.patch = None;
        // Best-effort rollback of any uncommitted dirty buffer. Errors
        // can't be surfaced from Drop; SQLite's WAL also discards an
        // uncommitted txn on next open, so a failed ROLLBACK here is
        // not load-bearing.
        let _ = self.conn.execute_batch("ROLLBACK;");
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn open_in_memory_writes_schema_and_stamps_pragmas() {
        let store = Store::open_in_memory().unwrap();
        assert_eq!(
            schema::read_application_id(store.conn()).unwrap(),
            APPLICATION_ID
        );
        assert_eq!(
            schema::read_user_version(store.conn()).unwrap(),
            SCHEMA_VERSION
        );
        // All seven user tables should exist.
        let tables: Vec<String> = store
            .conn()
            .prepare(
                "SELECT name FROM sqlite_master \
                 WHERE type='table' AND name NOT LIKE 'sqlite_%' \
                 ORDER BY name",
            )
            .unwrap()
            .query_map([], |r| r.get::<_, String>(0))
            .unwrap()
            .map(Result::unwrap)
            .collect();
        assert_eq!(
            tables,
            vec![
                "cell",
                "cell_format",
                "column_meta",
                "meta",
                "sheet",
                "undo_entry",
            ]
        );
    }

    #[test]
    fn reopening_existing_vlotus_db_is_idempotent() {
        let dir = tempdir();
        let path = dir.join("wb.db");

        // First open: writes schema, commits, drops.
        {
            let mut store = Store::open(&path).unwrap();
            store.commit().unwrap();
        }

        // Second open: pragmas already match, ensure_schema is a no-op.
        {
            let store = Store::open(&path).unwrap();
            assert_eq!(
                schema::read_application_id(store.conn()).unwrap(),
                APPLICATION_ID
            );
            assert_eq!(
                schema::read_user_version(store.conn()).unwrap(),
                SCHEMA_VERSION
            );
        }
    }

    #[test]
    fn legacy_datasette_db_rejects_with_pointer_to_migrate() {
        let dir = tempdir();
        let path = dir.join("legacy.db");
        let conn = Connection::open(&path).unwrap();
        conn.execute_batch(
            "CREATE TABLE datasette_sheets_sheet (id TEXT PRIMARY KEY);",
        )
        .unwrap();
        drop(conn);

        match Store::open(&path) {
            Err(StoreError::LegacyDatabase { path: p }) => assert_eq!(p, path),
            other => panic!("expected LegacyDatabase, got {other:?}"),
        }
    }

    #[test]
    fn unknown_db_with_foreign_app_id_rejected() {
        // Pick a value that's a valid i32 *and* not VLOT. SQLite's
        // pragma parser rejects hex literals like 0xDEADBEEF and any
        // decimal that overflows i32, silently storing 0 — verified
        // experimentally during T1 spike.
        const FOREIGN_APP_ID: i32 = 0x1234_5678;
        let dir = tempdir();
        let path = dir.join("foreign.db");
        let conn = Connection::open(&path).unwrap();
        conn.execute_batch(&format!(
            "PRAGMA application_id = {FOREIGN_APP_ID}; \
             CREATE TABLE bespoke (x INT);",
        ))
        .unwrap();
        drop(conn);

        match Store::open(&path) {
            Err(StoreError::UnknownDatabase { app_id, .. }) => {
                assert_eq!(app_id, FOREIGN_APP_ID);
            }
            other => panic!("expected UnknownDatabase, got {other:?}"),
        }
    }

    #[test]
    fn newer_schema_version_rejected() {
        let dir = tempdir();
        let path = dir.join("future.db");
        let conn = Connection::open(&path).unwrap();
        conn.execute_batch(&format!(
            "PRAGMA application_id = {APPLICATION_ID}; \
             PRAGMA user_version  = {next};",
            next = SCHEMA_VERSION + 1
        ))
        .unwrap();
        drop(conn);

        match Store::open(&path) {
            Err(StoreError::UnsupportedSchemaVersion {
                user_version,
                supported,
                ..
            }) => {
                assert_eq!(user_version, SCHEMA_VERSION + 1);
                assert_eq!(supported, SCHEMA_VERSION);
            }
            other => panic!("expected UnsupportedSchemaVersion, got {other:?}"),
        }
    }

    #[test]
    fn pre_existing_tables_with_default_pragmas_rejected() {
        // application_id=0 + user_version=0 is the "fresh" case, but only
        // when the file is genuinely empty. A foreign sqlite file with
        // those defaults but real tables must NOT silently get a vlotus
        // schema written on top.
        let dir = tempdir();
        let path = dir.join("masquerade.db");
        let conn = Connection::open(&path).unwrap();
        conn.execute_batch("CREATE TABLE other (x INT);").unwrap();
        drop(conn);

        match Store::open(&path) {
            Err(StoreError::UnknownDatabase { .. }) => {}
            other => panic!("expected UnknownDatabase, got {other:?}"),
        }
    }

    #[test]
    fn commit_persists_and_rollback_discards() {
        let dir = tempdir();
        let path = dir.join("txn.db");

        // Insert a row, commit, drop.
        {
            let mut store = Store::open(&path).unwrap();
            store
                .conn()
                .execute(
                    "INSERT INTO sheet (name) VALUES (?)",
                    rusqlite::params!["committed"],
                )
                .unwrap();
            store.commit().unwrap();
        }

        // Reopen, verify committed row survived.
        {
            let mut store = Store::open(&path).unwrap();
            let count: i64 = store
                .conn()
                .query_row(
                    "SELECT COUNT(*) FROM sheet WHERE name = ?",
                    rusqlite::params!["committed"],
                    |r| r.get(0),
                )
                .unwrap();
            assert_eq!(count, 1);

            // Insert another, then rollback.
            store
                .conn()
                .execute(
                    "INSERT INTO sheet (name) VALUES (?)",
                    rusqlite::params!["rolled-back"],
                )
                .unwrap();
            store.rollback().unwrap();
            // Store is reusable after rollback — the buffer txn was
            // re-opened.
            assert!(!store.is_dirty());
        }

        // Reopen, verify rolled-back row never made it.
        {
            let store = Store::open(&path).unwrap();
            let count: i64 = store
                .conn()
                .query_row(
                    "SELECT COUNT(*) FROM sheet WHERE name = ?",
                    rusqlite::params!["rolled-back"],
                    |r| r.get(0),
                )
                .unwrap();
            assert_eq!(count, 0);
            let count: i64 = store
                .conn()
                .query_row(
                    "SELECT COUNT(*) FROM sheet WHERE name = ?",
                    rusqlite::params!["committed"],
                    |r| r.get(0),
                )
                .unwrap();
            assert_eq!(count, 1);
        }
    }

    #[test]
    fn drop_without_commit_rolls_back() {
        let dir = tempdir();
        let path = dir.join("drop.db");

        {
            let store = Store::open(&path).unwrap();
            store
                .conn()
                .execute(
                    "INSERT INTO sheet (name) VALUES (?)",
                    rusqlite::params!["dropped"],
                )
                .unwrap();
            // No commit; Drop guard rolls back.
        }

        let store = Store::open(&path).unwrap();
        let count: i64 = store
            .conn()
            .query_row(
                "SELECT COUNT(*) FROM sheet WHERE name = ?",
                rusqlite::params!["dropped"],
                |r| r.get(0),
            )
            .unwrap();
        assert_eq!(count, 0);
    }

    #[test]
    fn is_dirty_flips_with_commit_and_rollback() {
        let mut store = Store::open_in_memory().unwrap();
        assert!(!store.is_dirty());
        store.create_sheet("S").unwrap();
        assert!(store.is_dirty());
        store.commit().unwrap();
        assert!(!store.is_dirty());
        store.create_sheet("T").unwrap();
        assert!(store.is_dirty());
        store.rollback().unwrap();
        assert!(!store.is_dirty());
        // T was rolled back; S survived.
        let names: Vec<String> = store
            .list_sheets()
            .unwrap()
            .into_iter()
            .map(|s| s.name)
            .collect();
        assert_eq!(names, vec!["S"]);
    }

    #[test]
    fn cell_format_fk_cascade_on_cell_delete() {
        let store = Store::open_in_memory().unwrap();
        let conn = store.conn();
        conn.execute("INSERT INTO sheet (name) VALUES ('s')", []).unwrap();
        conn.execute(
            "INSERT INTO cell (sheet_name, row, col, raw) VALUES ('s', 0, 0, 'x')",
            [],
        )
        .unwrap();
        conn.execute(
            "INSERT INTO cell_format (sheet_name, row, col, format_json) \
             VALUES ('s', 0, 0, '{\"b\":true}')",
            [],
        )
        .unwrap();
        conn.execute(
            "DELETE FROM cell WHERE sheet_name='s' AND row=0 AND col=0",
            [],
        )
        .unwrap();
        let count: i64 = conn
            .query_row("SELECT COUNT(*) FROM cell_format", [], |r| r.get(0))
            .unwrap();
        assert_eq!(count, 0);
    }

    #[test]
    fn sheet_rename_cascades_to_cells_and_formats() {
        let store = Store::open_in_memory().unwrap();
        let conn = store.conn();
        conn.execute("INSERT INTO sheet (name) VALUES ('old')", []).unwrap();
        conn.execute(
            "INSERT INTO cell (sheet_name, row, col, raw) VALUES ('old', 0, 0, 'x')",
            [],
        )
        .unwrap();
        conn.execute(
            "INSERT INTO cell_format (sheet_name, row, col, format_json) \
             VALUES ('old', 0, 0, '{\"b\":true}')",
            [],
        )
        .unwrap();

        conn.execute("UPDATE sheet SET name = 'new' WHERE name = 'old'", [])
            .unwrap();

        let cell_sheet: String = conn
            .query_row("SELECT sheet_name FROM cell", [], |r| r.get(0))
            .unwrap();
        assert_eq!(cell_sheet, "new");
        let format_sheet: String = conn
            .query_row("SELECT sheet_name FROM cell_format", [], |r| r.get(0))
            .unwrap();
        assert_eq!(format_sheet, "new");
    }

    /// Tiny helper: a tempdir under `std::env::temp_dir()` cleaned up by
    /// the test process exiting. Bringing in `tempfile` for one path is
    /// not worth the dep cost.
    fn tempdir() -> PathBuf {
        use std::time::{SystemTime, UNIX_EPOCH};
        let nanos = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!("vlotus-store-test-{nanos}"));
        std::fs::create_dir_all(&path).unwrap();
        path
    }
}
