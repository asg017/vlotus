//! Coordinate helpers — copy of `lotus_sqlite::coords`. T8 deletes the
//! lotus-sqlite original and this becomes the canonical home.

use lotus_core::types::{col_to_index, index_to_col, CellId};

pub fn col_idx_to_letters(col_idx: u32) -> String {
    index_to_col(col_idx + 1)
}

pub fn letters_to_col_idx(letters: &str) -> u32 {
    col_to_index(letters) - 1
}

pub fn to_cell_id(row_idx: u32, col_idx: u32) -> CellId {
    format!("{}{}", col_idx_to_letters(col_idx), row_idx + 1)
}

pub fn from_cell_id(cell_id: &str) -> Option<(u32, u32)> {
    let col_end = cell_id.find(|c: char| c.is_ascii_digit())?;
    let letters = &cell_id[..col_end];
    let row_num: u32 = cell_id[col_end..].parse().ok()?;
    if letters.is_empty() || row_num == 0 {
        return None;
    }
    Some((row_num - 1, letters_to_col_idx(letters)))
}
