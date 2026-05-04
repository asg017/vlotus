//! Cell upsert / load / recalculate. T2 of the storage rewrite (att
//! `zhutjuyl`). Format reads/writes still go through `lotus-sqlite`
//! against the legacy schema during the T2-T3 window — `cell_format`
//! lands in T3 (`v115x15m`).

use std::collections::HashSet;

use lotus_core::types::CellId;
use lotus_core::{CellValue, Sheet};
use rusqlite::params;

use super::coords::{from_cell_id, to_cell_id};
use super::{Store, StoreError};

/// Field names mirror `lotus_sqlite::CellChange` so vlotus' existing
/// call sites in `app.rs` can flip imports without churning every
/// reference. T3 adds `format_json` back as a sibling lookup against
/// `cell_format`.
#[derive(Debug, Clone)]
pub struct CellChange {
    pub row_idx: u32,
    pub col_idx: u32,
    /// Empty string = delete the cell.
    pub raw_value: String,
    /// T3 will route this through the `cell_format` table. Until then
    /// the field exists for shape-compatibility but is ignored.
    pub format_json: Option<String>,
}

/// Same shape as `lotus_sqlite::StoredCell` minus `sheet_id` (the
/// snapshot is per-sheet, so the sheet name is implied) so consumers
/// in `app.rs` and `format.rs` don't need a structural rewrite.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredCell {
    pub row_idx: u32,
    pub col_idx: u32,
    pub raw_value: String,
    pub computed_value: Option<String>,
    /// T3 will populate this from the `cell_format` table. T2 always
    /// returns `None` so format display is inert during the migration.
    pub format_json: Option<String>,
    pub spill_anchor_row: Option<u32>,
    pub spill_anchor_col: Option<u32>,
}

#[derive(Debug, Clone)]
pub struct SheetSnapshot {
    pub cells: Vec<StoredCell>,
}

impl Store {
    /// Apply a batch of cell changes against `sheet_name`, then run a
    /// full sheet recalc. Writes land in the dirty-buffer txn; the
    /// caller flushes them with [`Store::commit`] (`:w`) or discards
    /// with [`Store::rollback`] (`:q!`).
    pub fn apply(
        &mut self,
        sheet_name: &str,
        changes: &[CellChange],
    ) -> Result<(), StoreError> {
        for change in changes {
            if change.raw_value.is_empty() {
                // ON DELETE CASCADE on cell_format takes care of the
                // format row.
                self.conn.execute(
                    "DELETE FROM cell \
                     WHERE sheet_name = ? AND row = ? AND col = ?",
                    params![sheet_name, change.row_idx, change.col_idx],
                )?;
                continue;
            }

            // Upsert raw. Clear any prior spill ownership so a user
            // typing over a spill descendant turns it into an authored
            // cell.
            self.conn.execute(
                "INSERT INTO cell (sheet_name, row, col, raw) \
                 VALUES (?, ?, ?, ?) \
                 ON CONFLICT(sheet_name, row, col) DO UPDATE SET \
                   raw = excluded.raw, \
                   owner_row = NULL, \
                   owner_col = NULL",
                params![
                    sheet_name,
                    change.row_idx,
                    change.col_idx,
                    change.raw_value
                ],
            )?;

            // Format axis. Three cases mirror the prior `format_json`
            // semantics:
            //   - None  → preserve any existing format.
            //   - "null" → clear (delete the cell_format row).
            //   - other  → upsert as the new format.
            match change.format_json.as_deref() {
                None => {}
                Some("null") => {
                    self.conn.execute(
                        "DELETE FROM cell_format \
                         WHERE sheet_name = ? AND row = ? AND col = ?",
                        params![sheet_name, change.row_idx, change.col_idx],
                    )?;
                }
                Some(json) => {
                    self.conn.execute(
                        "INSERT INTO cell_format \
                             (sheet_name, row, col, format_json) \
                         VALUES (?, ?, ?, ?) \
                         ON CONFLICT(sheet_name, row, col) DO UPDATE SET \
                             format_json = excluded.format_json",
                        params![sheet_name, change.row_idx, change.col_idx, json],
                    )?;
                }
            }
        }

        self.conn.execute(
            "UPDATE sheet \
             SET updated_at = strftime('%Y-%m-%dT%H:%M:%fZ','now') \
             WHERE name = ?",
            params![sheet_name],
        )?;

        self.mark_dirty();
        // Recalc writes to `cell.computed` and `cell.owner_*`; those
        // are derivable, so any active patch session must skip them
        // — otherwise the changeset bloats and applying it on a
        // destination conflicts with the destination's own recalc.
        self.with_session_disabled(|s| s.recalculate(sheet_name))
    }

    /// Load every cell on a sheet (authored + spill-owned), ordered by
    /// `(row, col)`.
    pub fn load_sheet(&self, sheet_name: &str) -> Result<SheetSnapshot, StoreError> {
        let mut stmt = self.conn.prepare(
            "SELECT c.row, c.col, c.raw, c.computed, \
                    c.owner_row, c.owner_col, cf.format_json \
             FROM cell c \
             LEFT JOIN cell_format cf \
                 ON cf.sheet_name = c.sheet_name \
                AND cf.row = c.row \
                AND cf.col = c.col \
             WHERE c.sheet_name = ? \
             ORDER BY c.row, c.col",
        )?;
        let rows = stmt
            .query_map(params![sheet_name], |r| {
                Ok(StoredCell {
                    row_idx: r.get(0)?,
                    col_idx: r.get(1)?,
                    raw_value: r.get(2)?,
                    computed_value: r.get(3)?,
                    spill_anchor_row: r.get(4)?,
                    spill_anchor_col: r.get(5)?,
                    format_json: r.get(6)?,
                })
            })?
            .collect::<rusqlite::Result<Vec<_>>>()?;
        Ok(SheetSnapshot { cells: rows })
    }

    /// Convenience for `vlotus eval`: read the computed value of a
    /// single cell. Returns `None` if the cell doesn't exist or has no
    /// computed value (e.g. authored but evaluates to empty).
    pub fn get_computed(
        &self,
        sheet_name: &str,
        row: u32,
        col: u32,
    ) -> Result<Option<String>, StoreError> {
        let computed = self
            .conn
            .query_row(
                "SELECT computed FROM cell \
                 WHERE sheet_name = ? AND row = ? AND col = ?",
                params![sheet_name, row, col],
                |r| r.get::<_, Option<String>>(0),
            )
            .ok()
            .flatten();
        Ok(computed)
    }

    /// Recalculate the entire sheet. Loads authored cells into a fresh
    /// `lotus_core::Sheet`, runs the engine, writes computed values
    /// back, materialises spill-owned rows, and prunes stale ones.
    pub fn recalculate(&mut self, sheet_name: &str) -> Result<(), StoreError> {
        let snapshot = self.load_sheet(sheet_name)?;

        let mut sheet = Sheet::new();
        // Wire vlotus-local custom types and functions onto the engine
        // before any formulas evaluate. Hyperlinks first, then the
        // optional datetime extension when its feature is enabled.
        crate::hyperlink::register(&mut sheet)
            .map_err(|e| StoreError::Engine(e.to_string()))?;
        #[cfg(feature = "datetime")]
        crate::datetime::register(&mut sheet)
            .map_err(|e| StoreError::Engine(e.to_string()))?;
        let changes: Vec<(CellId, String)> = snapshot
            .cells
            .iter()
            .filter(|c| !c.raw_value.is_empty())
            .map(|c| (to_cell_id(c.row_idx, c.col_idx), c.raw_value.clone()))
            .collect();
        sheet
            .set_cells(&changes)
            .map_err(|e| StoreError::Engine(e.to_string()))?;

        let prior_rows: HashSet<(u32, u32)> = snapshot
            .cells
            .iter()
            .map(|c| (c.row_idx, c.col_idx))
            .collect();
        let prior_spill_owned: HashSet<(u32, u32)> = snapshot
            .cells
            .iter()
            .filter(|c| c.raw_value.is_empty() && c.spill_anchor_row.is_some())
            .map(|c| (c.row_idx, c.col_idx))
            .collect();

        // Drop stale custom-value entries for this sheet; the loop
        // below repopulates them for cells whose new value is a Custom.
        self.custom_cells
            .retain(|(s, _, _), _| s != sheet_name);

        let mut kept: HashSet<(u32, u32)> = HashSet::new();
        for (cell_id, value) in sheet.get_all() {
            let Some((row, col)) = from_cell_id(cell_id) else {
                continue;
            };
            let computed_str = match value {
                CellValue::Number(n) => Some(n.to_string()),
                CellValue::String(s) => Some(s.clone()),
                CellValue::Boolean(b) => Some(if *b { "TRUE".into() } else { "FALSE".into() }),
                CellValue::Empty => continue,
                CellValue::Error(e) => Some(e.to_string()),
                // Hyperlinks carve out: App::displayed_for / url_for_cell
                // both decode the `url + SEP + label` payload from
                // cell.computed. Calling display() here would strip the
                // URL and break click-to-open. TODO follow-up: source
                // URLs from cell.raw and drop this branch.
                CellValue::Custom(cv) if cv.type_tag == crate::hyperlink::TYPE_TAG => {
                    self.custom_cells.insert(
                        (sheet_name.to_string(), row, col),
                        crate::store::CustomCell {
                            type_tag: cv.type_tag.clone(),
                            data: cv.data.clone(),
                        },
                    );
                    Some(cv.data.clone())
                }
                // Other custom values render via the handler's display():
                // jspan → "1y 2mo 3d", jdatetime → space-separated, etc.
                CellValue::Custom(cv) => {
                    self.custom_cells.insert(
                        (sheet_name.to_string(), row, col),
                        crate::store::CustomCell {
                            type_tag: cv.type_tag.clone(),
                            data: cv.data.clone(),
                        },
                    );
                    Some(sheet.registry().display(cv))
                }
            };

            let owner = sheet
                .owner_of(cell_id)
                .and_then(|anchor| from_cell_id(anchor));
            let (owner_row, owner_col) = match owner {
                Some((r, c)) => (Some(r), Some(c)),
                None => (None, None),
            };

            if prior_rows.contains(&(row, col)) {
                self.conn.execute(
                    "UPDATE cell \
                     SET computed = ?, owner_row = ?, owner_col = ? \
                     WHERE sheet_name = ? AND row = ? AND col = ?",
                    params![computed_str, owner_row, owner_col, sheet_name, row, col],
                )?;
            } else {
                self.conn.execute(
                    "INSERT INTO cell \
                         (sheet_name, row, col, raw, computed, owner_row, owner_col) \
                     VALUES (?, ?, ?, '', ?, ?, ?)",
                    params![sheet_name, row, col, computed_str, owner_row, owner_col],
                )?;
            }
            kept.insert((row, col));
        }

        for cell in &snapshot.cells {
            let coords = (cell.row_idx, cell.col_idx);
            if kept.contains(&coords) {
                continue;
            }
            if prior_spill_owned.contains(&coords) {
                self.conn.execute(
                    "DELETE FROM cell \
                     WHERE sheet_name = ? AND row = ? AND col = ?",
                    params![sheet_name, cell.row_idx, cell.col_idx],
                )?;
            } else {
                self.conn.execute(
                    "UPDATE cell \
                     SET computed = NULL, owner_row = NULL, owner_col = NULL \
                     WHERE sheet_name = ? AND row = ? AND col = ?",
                    params![sheet_name, cell.row_idx, cell.col_idx],
                )?;
            }
        }

        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fresh(name: &str) -> Store {
        let mut store = Store::open_in_memory().unwrap();
        store.create_sheet(name).unwrap();
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
    fn apply_inserts_authored_cell_and_recalcs() {
        let mut store = fresh("S");
        store.apply("S", &[change(0, 0, "=1+2")]).unwrap();
        let snap = store.load_sheet("S").unwrap();
        assert_eq!(snap.cells.len(), 1);
        let c = &snap.cells[0];
        assert_eq!((c.row_idx, c.col_idx), (0, 0));
        assert_eq!(c.raw_value, "=1+2");
        assert_eq!(c.computed_value.as_deref(), Some("3"));
        assert!(c.spill_anchor_row.is_none() && c.spill_anchor_col.is_none());
    }

    #[test]
    fn apply_deletes_with_empty_raw() {
        let mut store = fresh("S");
        store.apply("S", &[change(0, 0, "hi")]).unwrap();
        store.apply("S", &[change(0, 0, "")]).unwrap();
        let snap = store.load_sheet("S").unwrap();
        assert!(snap.cells.is_empty());
    }

    #[test]
    fn apply_recalcs_dependents() {
        let mut store = fresh("S");
        store
            .apply("S", &[change(0, 0, "5"), change(0, 1, "=A1*2")])
            .unwrap();
        let snap = store.load_sheet("S").unwrap();
        let b1 = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (0, 1))
            .unwrap();
        assert_eq!(b1.computed_value.as_deref(), Some("10"));

        store.apply("S", &[change(0, 0, "7")]).unwrap();
        let snap = store.load_sheet("S").unwrap();
        let b1 = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (0, 1))
            .unwrap();
        assert_eq!(b1.computed_value.as_deref(), Some("14"));
    }

    #[test]
    fn recalculate_writes_spill_descendants() {
        let mut store = fresh("S");
        store.apply("S", &[change(0, 0, "=SEQUENCE(5)")]).unwrap();
        let snap = store.load_sheet("S").unwrap();
        assert_eq!(snap.cells.len(), 5);
        for c in &snap.cells {
            assert_eq!(c.spill_anchor_row, Some(0));
            assert_eq!(c.spill_anchor_col, Some(0));
        }
        let anchor = snap.cells.iter().find(|c| c.row_idx == 0).unwrap();
        assert_eq!(anchor.raw_value, "=SEQUENCE(5)");
        for c in snap.cells.iter().filter(|c| c.row_idx != 0) {
            assert!(c.raw_value.is_empty());
        }
    }

    #[test]
    fn recalculate_cleans_stale_spill_rows() {
        let mut store = fresh("S");
        store.apply("S", &[change(0, 0, "=SEQUENCE(5)")]).unwrap();
        store.apply("S", &[change(0, 0, "=SEQUENCE(2)")]).unwrap();
        let snap = store.load_sheet("S").unwrap();
        assert_eq!(snap.cells.len(), 2);
        for c in &snap.cells {
            assert_eq!(c.spill_anchor_row, Some(0));
        }
    }

    #[test]
    fn apply_overwrites_spill_owned_cell() {
        let mut store = fresh("S");
        store.apply("S", &[change(0, 0, "=SEQUENCE(3)")]).unwrap();
        store.apply("S", &[change(1, 0, "99")]).unwrap();
        let snap = store.load_sheet("S").unwrap();
        let c = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (1, 0))
            .unwrap();
        assert_eq!(c.raw_value, "99");
        assert!(c.spill_anchor_row.is_none(), "spill cleared on overwrite");
        assert!(c.spill_anchor_col.is_none());
    }

    #[test]
    fn load_sheet_orders_by_row_col() {
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[change(2, 1, "c"), change(0, 0, "a"), change(1, 2, "b")],
            )
            .unwrap();
        let snap = store.load_sheet("S").unwrap();
        let coords: Vec<_> = snap.cells.iter().map(|c| (c.row_idx, c.col_idx)).collect();
        assert_eq!(coords, vec![(0, 0), (1, 2), (2, 1)]);
    }

    #[test]
    fn get_computed_returns_engine_result() {
        let mut store = fresh("S");
        store.apply("S", &[change(0, 0, "=10*4")]).unwrap();
        assert_eq!(store.get_computed("S", 0, 0).unwrap().as_deref(), Some("40"));
    }

    #[test]
    fn get_computed_missing_cell_is_none() {
        let store = fresh("S");
        assert!(store.get_computed("S", 5, 5).unwrap().is_none());
    }

    #[test]
    fn hyperlink_formula_round_trips_through_storage() {
        // Verifies T3's vlotus-local registration: a HYPERLINK formula
        // evaluates against the registered custom function, the
        // resulting CustomValue is stored verbatim in cell.computed
        // (with the unit-separator intact), and a second recalculate
        // produces the same payload.
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[change(
                    0,
                    0,
                    "=HYPERLINK(\"https://example.com\", \"click here\")",
                )],
            )
            .unwrap();
        let computed = store.get_computed("S", 0, 0).unwrap().unwrap();
        assert_eq!(
            crate::hyperlink::split_payload(&computed),
            Some(("https://example.com", "click here")),
        );
        // Idempotent: recalculate doesn't perturb the encoded form.
        store.recalculate("S").unwrap();
        let again = store.get_computed("S", 0, 0).unwrap().unwrap();
        assert_eq!(again, computed);
    }

    #[test]
    fn hyperlink_one_arg_uses_url_as_label_through_storage() {
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[change(0, 0, "=HYPERLINK(\"https://example.com\")")],
            )
            .unwrap();
        let computed = store.get_computed("S", 0, 0).unwrap().unwrap();
        assert_eq!(
            crate::hyperlink::split_payload(&computed),
            Some(("https://example.com", "https://example.com")),
        );
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn jdate_round_trips_through_storage() {
        let mut store = fresh("S");
        store
            .apply("S", &[change(0, 0, "2025-04-27")])
            .unwrap();
        assert_eq!(
            store.get_computed("S", 0, 0).unwrap().as_deref(),
            Some("2025-04-27"),
        );
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn jspan_renders_friendly_in_computed() {
        // 2025-02-01 minus 2025-01-01 = 31 days. The friendly form jiff
        // uses for that span — assert it isn't the canonical "P31D".
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[change(0, 0, "=DATE(2025, 2, 1) - DATE(2025, 1, 1)")],
            )
            .unwrap();
        let computed = store.get_computed("S", 0, 0).unwrap().unwrap();
        assert!(
            !computed.starts_with('P'),
            "expected friendly form, got canonical: {computed:?}"
        );
        assert!(
            computed.contains("31"),
            "expected 31-day span, got: {computed:?}"
        );
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn jdatetime_renders_with_space_separator() {
        let mut store = fresh("S");
        store
            .apply("S", &[change(0, 0, "2025-04-27T12:30:00")])
            .unwrap();
        let computed = store.get_computed("S", 0, 0).unwrap().unwrap();
        assert!(
            !computed.contains('T'),
            "expected space separator, got: {computed:?}"
        );
        assert!(computed.contains("2025-04-27"));
        assert!(computed.contains("12:30"));
    }
}
