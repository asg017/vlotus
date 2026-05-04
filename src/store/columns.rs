//! Column metadata — `width` per `(sheet_name, col)`. T2 ports the
//! shape from `lotus_sqlite::ColumnMeta`/`load_columns`/`set_column_width`.
//! `name` is dropped — vlotus never read it (header rendering uses
//! `col_idx_to_letters`), and the new schema has no column to back it.

use rusqlite::params;

use super::{Store, StoreError};

#[derive(Debug, Clone)]
pub struct ColumnMeta {
    pub col_idx: u32,
    pub width: u32,
}

impl Store {
    pub fn load_columns(&self, sheet_name: &str) -> Result<Vec<ColumnMeta>, StoreError> {
        let mut stmt = self.conn.prepare(
            "SELECT col, COALESCE(width, 100) FROM column_meta \
             WHERE sheet_name = ? ORDER BY col",
        )?;
        let rows = stmt
            .query_map(params![sheet_name], |r| {
                Ok(ColumnMeta {
                    col_idx: r.get(0)?,
                    width: r.get(1)?,
                })
            })?
            .collect::<rusqlite::Result<Vec<_>>>()?;
        Ok(rows)
    }

    pub fn set_column_width(
        &mut self,
        sheet_name: &str,
        col: u32,
        width: u32,
    ) -> Result<(), StoreError> {
        self.conn.execute(
            "INSERT INTO column_meta (sheet_name, col, width) \
             VALUES (?, ?, ?) \
             ON CONFLICT(sheet_name, col) DO UPDATE SET width = excluded.width",
            params![sheet_name, col, width],
        )?;
        self.mark_dirty();
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

    #[test]
    fn set_then_load_round_trip() {
        let mut store = fresh("S");
        store.set_column_width("S", 0, 14).unwrap();
        store.set_column_width("S", 3, 22).unwrap();
        let cols = store.load_columns("S").unwrap();
        assert_eq!(cols.len(), 2);
        assert_eq!(cols[0].col_idx, 0);
        assert_eq!(cols[0].width, 14);
        assert_eq!(cols[1].col_idx, 3);
        assert_eq!(cols[1].width, 22);
    }

    #[test]
    fn set_overwrites_existing_width() {
        let mut store = fresh("S");
        store.set_column_width("S", 0, 14).unwrap();
        store.set_column_width("S", 0, 30).unwrap();
        let cols = store.load_columns("S").unwrap();
        assert_eq!(cols.len(), 1);
        assert_eq!(cols[0].width, 30);
    }

    #[test]
    fn delete_sheet_cascades_to_columns() {
        let mut store = fresh("S");
        store.set_column_width("S", 0, 14).unwrap();
        store.delete_sheet("S").unwrap();
        let count: i64 = store
            .conn()
            .query_row("SELECT COUNT(*) FROM column_meta", [], |r| r.get(0))
            .unwrap();
        assert_eq!(count, 0);
    }
}
