//! Persisted undo / redo log. T5 of the storage rewrite (att
//! `ngj1dfyc`). Each user action records its inverse ops as rows in
//! `undo_entry` inside the dirty-buffer txn; `Store::undo` pops the
//! latest group and replays the inverses; the App layer holds an
//! in-memory redo stack of popped groups so `Ctrl+r` can replay them
//! forward.
//!
//! `Store::commit` clears the undo log so each save window starts
//! fresh — matches the epic's "log dropped on every commit" decision.

use rusqlite::params;
use serde_json::{json, Value};

use super::{Store, StoreError};

/// Default cap on undo depth. Older groups are pruned on every
/// `record_undo_group` call.
pub const UNDO_DEPTH: usize = 50;

/// One inverse op: enough state to restore a single piece of
/// authoring metadata. Variants mirror the mutations vlotus records
/// today — cell raw + format together (so `:fmt` undo doesn't lose the
/// cell's value), column width.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum UndoOp {
    /// Restore a cell's raw value and format. `raw=None` means delete
    /// the cell. `format_json=None` means no format (no `cell_format`
    /// row); the special string `"null"` is the explicit clear sentinel
    /// preserved through `Store::apply` semantics.
    CellChange {
        sheet_name: String,
        row: u32,
        col: u32,
        raw: Option<String>,
        format_json: Option<String>,
    },
    /// Restore a column's width. `width=None` deletes the
    /// `column_meta` row.
    ColumnWidth {
        sheet_name: String,
        col: u32,
        width: Option<u32>,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UndoGroup {
    pub group_id: i64,
    pub ops: Vec<UndoOp>,
}

impl UndoOp {
    fn kind(&self) -> &'static str {
        match self {
            UndoOp::CellChange { .. } => "cell",
            UndoOp::ColumnWidth { .. } => "colwidth",
        }
    }

    fn to_payload(&self) -> String {
        let v = match self {
            UndoOp::CellChange {
                sheet_name,
                row,
                col,
                raw,
                format_json,
            } => json!({
                "s": sheet_name,
                "r": row,
                "c": col,
                "v": raw,
                "f": format_json,
            }),
            UndoOp::ColumnWidth {
                sheet_name,
                col,
                width,
            } => json!({
                "s": sheet_name,
                "c": col,
                "w": width,
            }),
        };
        v.to_string()
    }

    fn from_row(kind: &str, payload: &str) -> Result<Self, StoreError> {
        let v: Value = serde_json::from_str(payload)
            .map_err(|e| StoreError::Engine(format!("undo payload: {e}")))?;
        match kind {
            "cell" => Ok(UndoOp::CellChange {
                sheet_name: v
                    .get("s")
                    .and_then(Value::as_str)
                    .ok_or_else(|| StoreError::Engine("undo cell: missing s".into()))?
                    .to_string(),
                row: v
                    .get("r")
                    .and_then(Value::as_u64)
                    .ok_or_else(|| StoreError::Engine("undo cell: missing r".into()))?
                    as u32,
                col: v
                    .get("c")
                    .and_then(Value::as_u64)
                    .ok_or_else(|| StoreError::Engine("undo cell: missing c".into()))?
                    as u32,
                raw: v.get("v").and_then(Value::as_str).map(str::to_string),
                format_json: v.get("f").and_then(Value::as_str).map(str::to_string),
            }),
            "colwidth" => Ok(UndoOp::ColumnWidth {
                sheet_name: v
                    .get("s")
                    .and_then(Value::as_str)
                    .ok_or_else(|| StoreError::Engine("undo colwidth: missing s".into()))?
                    .to_string(),
                col: v
                    .get("c")
                    .and_then(Value::as_u64)
                    .ok_or_else(|| StoreError::Engine("undo colwidth: missing c".into()))?
                    as u32,
                width: v.get("w").and_then(Value::as_u64).map(|n| n as u32),
            }),
            other => Err(StoreError::Engine(format!("unknown undo kind: {other}"))),
        }
    }

    /// Apply this inverse op to the live store. Returns the *new*
    /// inverse — i.e., the op that would undo this restoration — so
    /// callers can populate a redo stack symmetrically.
    fn apply(&self, store: &mut Store) -> Result<UndoOp, StoreError> {
        match self {
            UndoOp::CellChange {
                sheet_name,
                row,
                col,
                raw,
                format_json,
            } => {
                // Capture the current state to roll forward into the
                // redo group later.
                let prior = store.snapshot_cell(sheet_name, *row, *col)?;
                // Now apply this inverse: write the captured raw +
                // format back. We bypass `Store::apply` to avoid
                // double-recording into the undo log.
                if let Some(raw_str) = raw {
                    store.conn().execute(
                        "INSERT INTO cell (sheet_name, row, col, raw) \
                         VALUES (?, ?, ?, ?) \
                         ON CONFLICT(sheet_name, row, col) DO UPDATE SET \
                             raw = excluded.raw, \
                             owner_row = NULL, \
                             owner_col = NULL",
                        params![sheet_name, row, col, raw_str],
                    )?;
                } else {
                    store.conn().execute(
                        "DELETE FROM cell \
                         WHERE sheet_name = ? AND row = ? AND col = ?",
                        params![sheet_name, row, col],
                    )?;
                }
                match format_json.as_deref() {
                    None => {
                        // Restoring "no format" means the cell_format
                        // row should be absent.
                        store.conn().execute(
                            "DELETE FROM cell_format \
                             WHERE sheet_name = ? AND row = ? AND col = ?",
                            params![sheet_name, row, col],
                        )?;
                    }
                    Some(json) => {
                        store.conn().execute(
                            "INSERT INTO cell_format \
                                 (sheet_name, row, col, format_json) \
                             VALUES (?, ?, ?, ?) \
                             ON CONFLICT(sheet_name, row, col) DO UPDATE SET \
                                 format_json = excluded.format_json",
                            params![sheet_name, row, col, json],
                        )?;
                    }
                }
                store.mark_dirty();
                store.recalculate(sheet_name)?;
                Ok(prior)
            }
            UndoOp::ColumnWidth {
                sheet_name,
                col,
                width,
            } => {
                let prior_width = store.column_width(sheet_name, *col)?;
                match width {
                    Some(w) => {
                        store.conn().execute(
                            "INSERT INTO column_meta (sheet_name, col, width) \
                             VALUES (?, ?, ?) \
                             ON CONFLICT(sheet_name, col) DO UPDATE SET \
                                 width = excluded.width",
                            params![sheet_name, col, w],
                        )?;
                    }
                    None => {
                        store.conn().execute(
                            "DELETE FROM column_meta \
                             WHERE sheet_name = ? AND col = ?",
                            params![sheet_name, col],
                        )?;
                    }
                }
                store.mark_dirty();
                Ok(UndoOp::ColumnWidth {
                    sheet_name: sheet_name.clone(),
                    col: *col,
                    width: prior_width,
                })
            }
        }
    }
}

impl Store {
    /// Read the current `(raw, format_json)` for a cell — used by
    /// undo bookkeeping to capture prior state.
    pub(super) fn snapshot_cell(
        &self,
        sheet_name: &str,
        row: u32,
        col: u32,
    ) -> Result<UndoOp, StoreError> {
        let row_format: (Option<String>, Option<String>) = self
            .conn
            .query_row(
                "SELECT c.raw, cf.format_json \
                 FROM cell c \
                 LEFT JOIN cell_format cf \
                     ON cf.sheet_name = c.sheet_name \
                    AND cf.row = c.row \
                    AND cf.col = c.col \
                 WHERE c.sheet_name = ? AND c.row = ? AND c.col = ?",
                params![sheet_name, row, col],
                |r| Ok((r.get(0)?, r.get(1)?)),
            )
            .unwrap_or((None, None));
        let (raw, format_json) = row_format;
        Ok(UndoOp::CellChange {
            sheet_name: sheet_name.to_string(),
            row,
            col,
            // A missing cell row → raw=None (delete). A present row
            // with empty raw means the cell exists but is spill-owned;
            // we still track the empty string so undo recreates it.
            raw,
            format_json,
        })
    }

    /// Read the current width for a column, `None` if no row.
    pub(super) fn column_width(
        &self,
        sheet_name: &str,
        col: u32,
    ) -> Result<Option<u32>, StoreError> {
        Ok(self
            .conn
            .query_row(
                "SELECT width FROM column_meta \
                 WHERE sheet_name = ? AND col = ?",
                params![sheet_name, col],
                |r| r.get::<_, Option<u32>>(0),
            )
            .ok()
            .flatten())
    }

    /// Persist a new undo group. Returns the assigned `group_id`.
    /// Prunes the oldest groups beyond [`UNDO_DEPTH`].
    pub fn record_undo_group(&mut self, ops: &[UndoOp]) -> Result<i64, StoreError> {
        if ops.is_empty() {
            return Ok(0);
        }
        let next_group_id: i64 = self
            .conn
            .query_row(
                "SELECT COALESCE(MAX(group_id), 0) + 1 FROM undo_entry",
                [],
                |r| r.get(0),
            )
            .unwrap_or(1);
        for op in ops {
            self.conn.execute(
                "INSERT INTO undo_entry (group_id, kind, payload) \
                 VALUES (?, ?, ?)",
                params![next_group_id, op.kind(), op.to_payload()],
            )?;
        }

        // Prune groups older than UNDO_DEPTH. Pick the oldest group_id
        // that should survive and delete everything below it.
        self.conn.execute(
            "DELETE FROM undo_entry \
             WHERE group_id <= ( \
                 SELECT group_id FROM ( \
                     SELECT DISTINCT group_id FROM undo_entry \
                     ORDER BY group_id DESC \
                     LIMIT 1 OFFSET ? \
                 ) \
             )",
            params![UNDO_DEPTH as i64],
        )?;
        Ok(next_group_id)
    }

    /// Pop the latest undo group, apply each op, and return a fresh
    /// group whose ops are the inverses-of-inverses (i.e., a redo).
    /// The App layer pushes this onto an in-memory redo stack.
    pub fn pop_undo(&mut self) -> Result<Option<UndoGroup>, StoreError> {
        let latest_group: Option<i64> = self
            .conn
            .query_row(
                "SELECT MAX(group_id) FROM undo_entry",
                [],
                |r| r.get::<_, Option<i64>>(0),
            )
            .ok()
            .flatten();
        let group_id = match latest_group {
            Some(id) => id,
            None => return Ok(None),
        };

        let ops = self.load_group_ops(group_id)?;
        // Delete the group from the log BEFORE applying — applying
        // does not write to undo_entry, but redo will need to.
        self.conn.execute(
            "DELETE FROM undo_entry WHERE group_id = ?",
            params![group_id],
        )?;

        let mut redo_ops = Vec::with_capacity(ops.len());
        for op in &ops {
            redo_ops.push(op.apply(self)?);
        }
        Ok(Some(UndoGroup {
            group_id,
            ops: redo_ops,
        }))
    }

    /// Re-apply a redo group. Returns a fresh undo group (the original
    /// inverses again) for the App to push back onto the undo log.
    pub fn apply_redo(&mut self, group: &UndoGroup) -> Result<UndoGroup, StoreError> {
        let mut new_inverses = Vec::with_capacity(group.ops.len());
        for op in &group.ops {
            new_inverses.push(op.apply(self)?);
        }
        // Persist the new undo group so further `u` keystrokes walk
        // the redo back the other way.
        let new_group_id = self.record_undo_group(&new_inverses)?;
        Ok(UndoGroup {
            group_id: new_group_id,
            ops: new_inverses,
        })
    }

    /// Drop every entry in `undo_entry`. Called from `Store::commit`
    /// so each save window starts with a clean log (per epic
    /// decision: undo doesn't survive `:w`).
    pub fn clear_undo_log(&mut self) -> Result<(), StoreError> {
        self.conn.execute("DELETE FROM undo_entry", [])?;
        Ok(())
    }

    fn load_group_ops(&self, group_id: i64) -> Result<Vec<UndoOp>, StoreError> {
        let mut stmt = self.conn.prepare(
            "SELECT kind, payload FROM undo_entry \
             WHERE group_id = ? ORDER BY id",
        )?;
        let rows = stmt
            .query_map(params![group_id], |r| {
                let kind: String = r.get(0)?;
                let payload: String = r.get(1)?;
                Ok((kind, payload))
            })?
            .collect::<rusqlite::Result<Vec<_>>>()?;
        let mut ops = Vec::with_capacity(rows.len());
        for (kind, payload) in rows {
            ops.push(UndoOp::from_row(&kind, &payload)?);
        }
        Ok(ops)
    }
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

    fn snapshot_for(store: &Store, sheet: &str, row: u32, col: u32) -> UndoOp {
        store.snapshot_cell(sheet, row, col).unwrap()
    }

    #[test]
    fn record_then_undo_restores_cell() {
        let mut store = fresh("S");
        // Capture prior state, write change, record inverse.
        let prior = snapshot_for(&store, "S", 0, 0);
        store.apply("S", &[change(0, 0, "5")]).unwrap();
        store.record_undo_group(&[prior]).unwrap();

        // Pop and apply the undo group.
        let popped = store.pop_undo().unwrap().unwrap();
        assert_eq!(popped.ops.len(), 1);
        let snap = store.load_sheet("S").unwrap();
        assert!(snap.cells.is_empty(), "undo restored to empty");
    }

    #[test]
    fn undo_then_redo_round_trips() {
        let mut store = fresh("S");
        let prior = snapshot_for(&store, "S", 0, 0);
        store.apply("S", &[change(0, 0, "hello")]).unwrap();
        store.record_undo_group(&[prior]).unwrap();

        let undone = store.pop_undo().unwrap().unwrap();
        assert!(store.load_sheet("S").unwrap().cells.is_empty());

        let _redo_inverse = store.apply_redo(&undone).unwrap();
        let snap = store.load_sheet("S").unwrap();
        assert_eq!(snap.cells.len(), 1);
        assert_eq!(snap.cells[0].raw_value, "hello");
    }

    #[test]
    fn undo_log_pruned_at_depth() {
        let mut store = fresh("S");
        // Record UNDO_DEPTH + 5 distinct groups.
        for i in 0..(UNDO_DEPTH as u32 + 5) {
            let prior = snapshot_for(&store, "S", i, 0);
            store.apply("S", &[change(i, 0, "x")]).unwrap();
            store.record_undo_group(&[prior]).unwrap();
        }
        let count: i64 = store
            .conn()
            .query_row(
                "SELECT COUNT(DISTINCT group_id) FROM undo_entry",
                [],
                |r| r.get(0),
            )
            .unwrap();
        assert_eq!(count as usize, UNDO_DEPTH);
    }

    #[test]
    fn commit_clears_undo_log() {
        let mut store = fresh("S");
        let prior = snapshot_for(&store, "S", 0, 0);
        store.apply("S", &[change(0, 0, "x")]).unwrap();
        store.record_undo_group(&[prior]).unwrap();
        store.commit().unwrap();
        let count: i64 = store
            .conn()
            .query_row("SELECT COUNT(*) FROM undo_entry", [], |r| r.get(0))
            .unwrap();
        assert_eq!(count, 0);
        // pop_undo on an empty log returns None.
        assert!(store.pop_undo().unwrap().is_none());
    }

    #[test]
    fn rollback_discards_undo_log() {
        let mut store = fresh("S");
        let prior = snapshot_for(&store, "S", 0, 0);
        store.apply("S", &[change(0, 0, "x")]).unwrap();
        store.record_undo_group(&[prior]).unwrap();
        store.rollback().unwrap();
        let count: i64 = store
            .conn()
            .query_row("SELECT COUNT(*) FROM undo_entry", [], |r| r.get(0))
            .unwrap();
        assert_eq!(count, 0);
    }

    #[test]
    fn group_with_multiple_ops_undoes_all_at_once() {
        let mut store = fresh("S");
        let p1 = snapshot_for(&store, "S", 0, 0);
        let p2 = snapshot_for(&store, "S", 0, 1);
        let p3 = snapshot_for(&store, "S", 0, 2);
        store
            .apply(
                "S",
                &[change(0, 0, "a"), change(0, 1, "b"), change(0, 2, "c")],
            )
            .unwrap();
        store.record_undo_group(&[p1, p2, p3]).unwrap();
        let popped = store.pop_undo().unwrap().unwrap();
        assert_eq!(popped.ops.len(), 3);
        assert!(store.load_sheet("S").unwrap().cells.is_empty());
    }

    #[test]
    fn column_width_undo_round_trip() {
        let mut store = fresh("S");
        let prior = UndoOp::ColumnWidth {
            sheet_name: "S".into(),
            col: 2,
            width: store.column_width("S", 2).unwrap(),
        };
        store.set_column_width("S", 2, 30).unwrap();
        store.record_undo_group(&[prior]).unwrap();
        store.pop_undo().unwrap();
        let cols = store.load_columns("S").unwrap();
        assert!(cols.is_empty(), "undo restored to no width row");
    }

    #[test]
    fn undo_payload_preserves_unicode_and_quotes() {
        // raw_value can contain anything user types — make sure
        // serde_json handles roundtrip cleanly.
        let mut store = fresh("S");
        let weird = "hello \"world\" 🦀\n\\tab".to_string();
        let prior = snapshot_for(&store, "S", 0, 0);
        store.apply("S", &[change(0, 0, &weird)]).unwrap();
        let snapshot_after = snapshot_for(&store, "S", 0, 0);
        store.record_undo_group(&[prior, snapshot_after.clone()]).unwrap();
        // Read back the second op as a sanity check.
        let popped = store.pop_undo().unwrap().unwrap();
        assert_eq!(popped.ops.len(), 2);
    }
}
