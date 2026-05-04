//! Sheet management — minimal slice carved off from T4 (`u9dgo2ug`)
//! because T2's `cell.sheet_name → sheet(name)` foreign key would
//! otherwise fail every cell write. Currently covers create / list /
//! delete / rename; `set_sheet_color`, `reorder_sheet` are still T4
//! follow-ups.

use rusqlite::params;

use super::{Store, StoreError};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetMeta {
    pub name: String,
    pub sort_order: i64,
    pub color: Option<String>,
}

impl Store {
    pub fn list_sheets(&self) -> Result<Vec<SheetMeta>, StoreError> {
        let mut stmt = self.conn().prepare(
            "SELECT name, sort_order, color FROM sheet \
             ORDER BY sort_order, created_at",
        )?;
        let rows = stmt
            .query_map([], |r| {
                Ok(SheetMeta {
                    name: r.get(0)?,
                    sort_order: r.get(1)?,
                    color: r.get(2)?,
                })
            })?
            .collect::<rusqlite::Result<Vec<_>>>()?;
        Ok(rows)
    }

    /// Used by T4's `add_sheet` uniqueness check; App today does the
    /// in-memory `iter().any(|s| s.name == name)` directly.
    #[allow(dead_code)]
    pub fn sheet_exists(&self, name: &str) -> Result<bool, StoreError> {
        let exists: bool = self.conn().query_row(
            "SELECT EXISTS(SELECT 1 FROM sheet WHERE name = ?)",
            params![name],
            |r| r.get(0),
        )?;
        Ok(exists)
    }

    /// Insert a new sheet. Returns `Ok(())` even if the name already
    /// exists — callers that need uniqueness should check
    /// `sheet_exists` first. (T4 will surface a `DuplicateSheetName`
    /// error variant; T2 keeps the API minimal.)
    pub fn create_sheet(&mut self, name: &str) -> Result<(), StoreError> {
        // `sort_order` defaults to 0 in DDL; if every sheet has 0 the
        // tiebreak falls to `created_at`, which is also defaulted —
        // good enough for T2.
        self.conn().execute(
            "INSERT OR IGNORE INTO sheet (name) VALUES (?)",
            params![name],
        )?;
        self.mark_dirty();
        Ok(())
    }

    pub fn delete_sheet(&mut self, name: &str) -> Result<(), StoreError> {
        // ON DELETE CASCADE on column_meta / cell / cell_format takes
        // care of dependent rows.
        self.conn()
            .execute("DELETE FROM sheet WHERE name = ?", params![name])?;
        self.mark_dirty();
        Ok(())
    }

    /// Rename `old` to `new`. FK `ON UPDATE CASCADE` rewrites
    /// `cell.sheet_name` / `cell_format.sheet_name` / `column_meta.sheet_name`
    /// in the same statement. Returns `DuplicateSheetName` if `new` is
    /// already taken; no-op when `old == new`. Caller is responsible for
    /// clearing undo log entries (their stored sheet_name strings would
    /// dangle after the rename).
    pub fn rename_sheet(&mut self, old: &str, new: &str) -> Result<(), StoreError> {
        if old == new {
            return Ok(());
        }
        if self.sheet_exists(new)? {
            return Err(StoreError::DuplicateSheetName(new.to_string()));
        }
        self.conn().execute(
            "UPDATE sheet SET name = ? WHERE name = ?",
            params![new, old],
        )?;
        self.mark_dirty();
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn create_then_list_returns_sheet() {
        let mut store = Store::open_in_memory().unwrap();
        store.create_sheet("Sheet1").unwrap();
        let sheets = store.list_sheets().unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].name, "Sheet1");
    }

    #[test]
    fn sheet_exists_round_trip() {
        let mut store = Store::open_in_memory().unwrap();
        assert!(!store.sheet_exists("missing").unwrap());
        store.create_sheet("present").unwrap();
        assert!(store.sheet_exists("present").unwrap());
    }

    #[test]
    fn create_is_idempotent() {
        // T2's INSERT OR IGNORE policy: re-creating an existing sheet
        // is a no-op. T4 will swap this for a `DuplicateSheetName`
        // error; until then App layers can call `create_sheet`
        // unconditionally on bootstrap.
        let mut store = Store::open_in_memory().unwrap();
        store.create_sheet("S").unwrap();
        store.create_sheet("S").unwrap();
        assert_eq!(store.list_sheets().unwrap().len(), 1);
    }

    #[test]
    fn rename_sheet_cascades_through_public_api() {
        let mut store = Store::open_in_memory().unwrap();
        store.create_sheet("old").unwrap();
        store
            .conn()
            .execute(
                "INSERT INTO cell (sheet_name, row, col, raw) VALUES ('old', 1, 2, 'x')",
                [],
            )
            .unwrap();

        store.rename_sheet("old", "new").unwrap();

        assert!(store.sheet_exists("new").unwrap());
        assert!(!store.sheet_exists("old").unwrap());
        let cell_sheet: String = store
            .conn()
            .query_row("SELECT sheet_name FROM cell", [], |r| r.get(0))
            .unwrap();
        assert_eq!(cell_sheet, "new");
    }

    #[test]
    fn rename_sheet_rejects_existing_target() {
        let mut store = Store::open_in_memory().unwrap();
        store.create_sheet("a").unwrap();
        store.create_sheet("b").unwrap();
        let err = store.rename_sheet("a", "b").unwrap_err();
        assert!(matches!(err, StoreError::DuplicateSheetName(ref n) if n == "b"));
        // Original is untouched.
        assert!(store.sheet_exists("a").unwrap());
        assert!(store.sheet_exists("b").unwrap());
    }

    #[test]
    fn rename_sheet_to_self_is_noop() {
        let mut store = Store::open_in_memory().unwrap();
        store.create_sheet("only").unwrap();
        store.rename_sheet("only", "only").unwrap();
        assert!(store.sheet_exists("only").unwrap());
    }

    #[test]
    fn delete_cascades_to_cells() {
        let mut store = Store::open_in_memory().unwrap();
        store.create_sheet("S").unwrap();
        store
            .conn()
            .execute(
                "INSERT INTO cell (sheet_name, row, col, raw) VALUES ('S', 0, 0, 'x')",
                [],
            )
            .unwrap();
        store.delete_sheet("S").unwrap();
        let count: i64 = store
            .conn()
            .query_row("SELECT COUNT(*) FROM cell", [], |r| r.get(0))
            .unwrap();
        assert_eq!(count, 0);
    }
}
