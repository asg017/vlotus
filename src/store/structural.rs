//! Structural row/column ops — insert/delete with coordinate shift +
//! formula reference rewrite. Ports the SHIFT_PARK two-pass trick from
//! `lotus_sqlite::store` so an in-place `UPDATE row = row + 1` doesn't
//! collide with the composite PK.

use rusqlite::params;

use super::{Store, StoreError};

/// Intermediate offset used during coordinate shifts to keep PK
/// uniqueness while we move cells through their target positions. Must
/// be larger than any real cell coordinate (the engine grid is far
/// smaller than 10⁹).
const SHIFT_PARK: u32 = 1_000_000;

impl Store {
    pub fn insert_rows(
        &mut self,
        sheet_name: &str,
        at: u32,
        count: u32,
    ) -> Result<(), StoreError> {
        if count == 0 {
            return Ok(());
        }
        {
            // Park then unpark. `cell_format.row` rides along via
            // `ON UPDATE CASCADE` on its (sheet_name, row, col) FK to
            // `cell` — explicitly shifting it here would double-apply.
            self.conn.execute(
                "UPDATE cell SET row = row + ? \
                 WHERE sheet_name = ? AND row >= ?",
                params![SHIFT_PARK, sheet_name, at],
            )?;
            self.conn.execute(
                "UPDATE cell SET row = row - ? + ? \
                 WHERE sheet_name = ? AND row >= ?",
                params![SHIFT_PARK, count, sheet_name, SHIFT_PARK],
            )?;
            // owner_row pointers must shift in lockstep with the
            // coordinate they reference.
            self.conn.execute(
                "UPDATE cell SET owner_row = owner_row + ? \
                 WHERE sheet_name = ? AND owner_row IS NOT NULL AND owner_row >= ?",
                params![count, sheet_name, at],
            )?;
        }

        let insertion = lotus_core::Insertion {
            rows: vec![at; count as usize],
            cols: vec![],
        };
        self.rewrite_formulas(sheet_name, |raw| {
            lotus_core::adjust_refs_for_insertion(raw, &insertion).into_owned()
        })?;
        self.mark_dirty();
        // See `Store::apply` for why recalc is bracketed out of any
        // active patch session.
        self.with_session_disabled(|s| s.recalculate(sheet_name))
    }

    pub fn insert_cols(
        &mut self,
        sheet_name: &str,
        at: u32,
        count: u32,
    ) -> Result<(), StoreError> {
        if count == 0 {
            return Ok(());
        }
        {
            // `cell_format.col` rides along via ON UPDATE CASCADE.
            self.conn.execute(
                "UPDATE cell SET col = col + ? \
                 WHERE sheet_name = ? AND col >= ?",
                params![SHIFT_PARK, sheet_name, at],
            )?;
            self.conn.execute(
                "UPDATE cell SET col = col - ? + ? \
                 WHERE sheet_name = ? AND col >= ?",
                params![SHIFT_PARK, count, sheet_name, SHIFT_PARK],
            )?;
            self.conn.execute(
                "UPDATE cell SET owner_col = owner_col + ? \
                 WHERE sheet_name = ? AND owner_col IS NOT NULL AND owner_col >= ?",
                params![count, sheet_name, at],
            )?;
            // column_meta tracks per-column widths and must shift too.
            self.conn.execute(
                "UPDATE column_meta SET col = col + ? \
                 WHERE sheet_name = ? AND col >= ?",
                params![SHIFT_PARK, sheet_name, at],
            )?;
            self.conn.execute(
                "UPDATE column_meta SET col = col - ? + ? \
                 WHERE sheet_name = ? AND col >= ?",
                params![SHIFT_PARK, count, sheet_name, SHIFT_PARK],
            )?;
        }

        let insertion = lotus_core::Insertion {
            rows: vec![],
            cols: vec![at; count as usize],
        };
        self.rewrite_formulas(sheet_name, |raw| {
            lotus_core::adjust_refs_for_insertion(raw, &insertion).into_owned()
        })?;
        self.mark_dirty();
        // See `Store::apply` for why recalc is bracketed out of any
        // active patch session.
        self.with_session_disabled(|s| s.recalculate(sheet_name))
    }

    pub fn delete_rows(
        &mut self,
        sheet_name: &str,
        start: u32,
        count: u32,
    ) -> Result<(), StoreError> {
        if count == 0 {
            return Ok(());
        }
        let end_excl = start.saturating_add(count);
        {
            // Drop the deleted band — `cell_format` cascades via FK on
            // (sheet_name, row, col). Owner pointers into the deleted
            // band are blanked out (their anchor formulas are about to
            // recalc anyway).
            self.conn.execute(
                "DELETE FROM cell \
                 WHERE sheet_name = ? AND row >= ? AND row < ?",
                params![sheet_name, start, end_excl],
            )?;
            // Park-shift the rows below. `cell_format.row` rides along
            // via ON UPDATE CASCADE — explicit shift would double-apply
            // and break the FK.
            self.conn.execute(
                "UPDATE cell SET row = row + ? \
                 WHERE sheet_name = ? AND row >= ?",
                params![SHIFT_PARK, sheet_name, end_excl],
            )?;
            self.conn.execute(
                "UPDATE cell SET row = row - ? - ? \
                 WHERE sheet_name = ? AND row >= ?",
                params![SHIFT_PARK, count, sheet_name, SHIFT_PARK],
            )?;
            self.conn.execute(
                "UPDATE cell SET owner_row = owner_row - ? \
                 WHERE sheet_name = ? AND owner_row IS NOT NULL AND owner_row >= ?",
                params![count, sheet_name, end_excl],
            )?;
        }

        let deletion = lotus_core::Deletion {
            rows: (start..end_excl).collect(),
            cols: vec![],
        };
        self.rewrite_formulas(sheet_name, |raw| {
            lotus_core::adjust_refs_for_deletion(raw, &deletion).into_owned()
        })?;
        self.mark_dirty();
        // See `Store::apply` for why recalc is bracketed out of any
        // active patch session.
        self.with_session_disabled(|s| s.recalculate(sheet_name))
    }

    pub fn delete_cols(
        &mut self,
        sheet_name: &str,
        start: u32,
        count: u32,
    ) -> Result<(), StoreError> {
        if count == 0 {
            return Ok(());
        }
        let end_excl = start.saturating_add(count);
        {
            self.conn.execute(
                "DELETE FROM cell \
                 WHERE sheet_name = ? AND col >= ? AND col < ?",
                params![sheet_name, start, end_excl],
            )?;
            self.conn.execute(
                "DELETE FROM column_meta \
                 WHERE sheet_name = ? AND col >= ? AND col < ?",
                params![sheet_name, start, end_excl],
            )?;
            // `cell_format.col` rides along via ON UPDATE CASCADE.
            self.conn.execute(
                "UPDATE cell SET col = col + ? \
                 WHERE sheet_name = ? AND col >= ?",
                params![SHIFT_PARK, sheet_name, end_excl],
            )?;
            self.conn.execute(
                "UPDATE cell SET col = col - ? - ? \
                 WHERE sheet_name = ? AND col >= ?",
                params![SHIFT_PARK, count, sheet_name, SHIFT_PARK],
            )?;
            self.conn.execute(
                "UPDATE cell SET owner_col = owner_col - ? \
                 WHERE sheet_name = ? AND owner_col IS NOT NULL AND owner_col >= ?",
                params![count, sheet_name, end_excl],
            )?;
            self.conn.execute(
                "UPDATE column_meta SET col = col + ? \
                 WHERE sheet_name = ? AND col >= ?",
                params![SHIFT_PARK, sheet_name, end_excl],
            )?;
            self.conn.execute(
                "UPDATE column_meta SET col = col - ? - ? \
                 WHERE sheet_name = ? AND col >= ?",
                params![SHIFT_PARK, count, sheet_name, SHIFT_PARK],
            )?;
        }

        let deletion = lotus_core::Deletion {
            rows: vec![],
            cols: (start..end_excl).collect(),
        };
        self.rewrite_formulas(sheet_name, |raw| {
            lotus_core::adjust_refs_for_deletion(raw, &deletion).into_owned()
        })?;
        self.mark_dirty();
        // See `Store::apply` for why recalc is bracketed out of any
        // active patch session.
        self.with_session_disabled(|s| s.recalculate(sheet_name))
    }

    /// Apply `transform` to every formula's raw text on the sheet. Used
    /// by the structural ops above to retarget refs after a coordinate
    /// shift.
    fn rewrite_formulas<F: Fn(&str) -> String>(
        &self,
        sheet_name: &str,
        transform: F,
    ) -> Result<(), StoreError> {
        let snapshot: Vec<(u32, u32, String)> = {
            let mut stmt = self.conn.prepare(
                "SELECT row, col, raw FROM cell \
                 WHERE sheet_name = ? AND raw LIKE '=%'",
            )?;
            let rows = stmt
                .query_map(params![sheet_name], |r| {
                    Ok((
                        r.get::<_, u32>(0)?,
                        r.get::<_, u32>(1)?,
                        r.get::<_, String>(2)?,
                    ))
                })?
                .collect::<rusqlite::Result<Vec<_>>>()?;
            rows
        };
        for (r, c, raw) in snapshot {
            let new_raw = transform(&raw);
            if new_raw != raw {
                self.conn.execute(
                    "UPDATE cell SET raw = ? \
                     WHERE sheet_name = ? AND row = ? AND col = ?",
                    params![new_raw, sheet_name, r, c],
                )?;
            }
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::super::CellChange;
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
    fn insert_rows_shifts_cells_and_rewrites_refs() {
        let mut store = fresh("S");
        store
            .apply("S", &[change(0, 0, "5"), change(0, 1, "=A1*2")])
            .unwrap();
        store.insert_rows("S", 0, 1).unwrap();
        let snap = store.load_sheet("S").unwrap();
        // Row 0 → row 1.
        let a2 = snap.cells.iter().find(|c| (c.row_idx, c.col_idx) == (1, 0));
        let b2 = snap.cells.iter().find(|c| (c.row_idx, c.col_idx) == (1, 1));
        assert!(a2.is_some());
        assert_eq!(b2.unwrap().raw_value, "=A2*2"); // ref shifted
        assert_eq!(b2.unwrap().computed_value.as_deref(), Some("10"));
    }

    #[test]
    fn delete_rows_drops_band_and_shifts_below() {
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[change(0, 0, "a"), change(1, 0, "b"), change(2, 0, "c")],
            )
            .unwrap();
        store.delete_rows("S", 1, 1).unwrap();
        let snap = store.load_sheet("S").unwrap();
        let coords: Vec<_> = snap.cells.iter().map(|c| (c.row_idx, c.col_idx)).collect();
        assert_eq!(coords, vec![(0, 0), (1, 0)]);
        let row1 = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (1, 0))
            .unwrap();
        assert_eq!(row1.raw_value, "c");
    }

    #[test]
    fn insert_cols_shifts_column_widths() {
        let mut store = fresh("S");
        store.set_column_width("S", 0, 14).unwrap();
        store.set_column_width("S", 1, 22).unwrap();
        store.insert_cols("S", 0, 1).unwrap();
        let cols = store.load_columns("S").unwrap();
        assert_eq!(cols.len(), 2);
        // Both columns shift right by 1.
        assert_eq!(cols[0].col_idx, 1);
        assert_eq!(cols[0].width, 14);
        assert_eq!(cols[1].col_idx, 2);
        assert_eq!(cols[1].width, 22);
    }

    #[test]
    fn insert_rows_carries_cell_format_through_shift() {
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[CellChange {
                    row_idx: 0,
                    col_idx: 0,
                    raw_value: "x".into(),
                    format_json: Some("{\"b\":true}".into()),
                }],
            )
            .unwrap();
        store.insert_rows("S", 0, 1).unwrap();
        let snap = store.load_sheet("S").unwrap();
        let cell = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (1, 0))
            .expect("formatted cell shifted to row 1");
        assert_eq!(cell.format_json.as_deref(), Some("{\"b\":true}"));
    }

    #[test]
    fn insert_cols_carries_cell_format_through_shift() {
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[CellChange {
                    row_idx: 0,
                    col_idx: 0,
                    raw_value: "x".into(),
                    format_json: Some("{\"b\":true}".into()),
                }],
            )
            .unwrap();
        store.insert_cols("S", 0, 1).unwrap();
        let snap = store.load_sheet("S").unwrap();
        let cell = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (0, 1))
            .expect("formatted cell shifted to col 1");
        assert_eq!(cell.format_json.as_deref(), Some("{\"b\":true}"));
    }

    #[test]
    fn delete_rows_carries_cell_format_through_shift() {
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[
                    CellChange {
                        row_idx: 0,
                        col_idx: 0,
                        raw_value: "drop".into(),
                        format_json: None,
                    },
                    CellChange {
                        row_idx: 1,
                        col_idx: 0,
                        raw_value: "keep".into(),
                        format_json: Some("{\"b\":true}".into()),
                    },
                ],
            )
            .unwrap();
        store.delete_rows("S", 0, 1).unwrap();
        let snap = store.load_sheet("S").unwrap();
        let cell = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (0, 0))
            .expect("formatted cell shifted up to row 0");
        assert_eq!(cell.raw_value, "keep");
        assert_eq!(cell.format_json.as_deref(), Some("{\"b\":true}"));
    }

    #[test]
    fn delete_cols_carries_cell_format_through_shift() {
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[
                    CellChange {
                        row_idx: 0,
                        col_idx: 0,
                        raw_value: "drop".into(),
                        format_json: None,
                    },
                    CellChange {
                        row_idx: 0,
                        col_idx: 1,
                        raw_value: "keep".into(),
                        format_json: Some("{\"b\":true}".into()),
                    },
                ],
            )
            .unwrap();
        store.delete_cols("S", 0, 1).unwrap();
        let snap = store.load_sheet("S").unwrap();
        let cell = snap
            .cells
            .iter()
            .find(|c| (c.row_idx, c.col_idx) == (0, 0))
            .expect("formatted cell shifted left to col 0");
        assert_eq!(cell.raw_value, "keep");
        assert_eq!(cell.format_json.as_deref(), Some("{\"b\":true}"));
    }

    #[test]
    fn delete_cols_drops_band_and_column_widths() {
        let mut store = fresh("S");
        store
            .apply(
                "S",
                &[change(0, 0, "x"), change(0, 1, "y"), change(0, 2, "z")],
            )
            .unwrap();
        store.set_column_width("S", 1, 22).unwrap();
        store.delete_cols("S", 1, 1).unwrap();
        let snap = store.load_sheet("S").unwrap();
        let coords: Vec<_> = snap.cells.iter().map(|c| (c.row_idx, c.col_idx)).collect();
        assert_eq!(coords, vec![(0, 0), (0, 1)]);
        let cols = store.load_columns("S").unwrap();
        assert!(cols.is_empty(), "deleted column's width was dropped");
    }
}
