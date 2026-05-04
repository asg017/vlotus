//! Bundled vimtutor-style curriculum for `vlotus tutor`.
//!
//! Each lesson is one sheet in the tutorial workbook. Column A holds
//! the instructions (full text visible in the formula bar as the
//! cursor moves down it); columns B–H hold practice data the student
//! manipulates. Sheet names follow `LN: <topic>` so the tab list reads
//! as a curriculum, and `:reset` finds the matching lesson by sheet id.
//!
//! Implementation deliberately keeps lessons as a `const &[Lesson]` so
//! diffs are reviewable when content evolves — much friendlier than
//! checking in a binary `.db`.

use crate::app::App;
use crate::store::{CellChange, Store};

pub struct Lesson {
    /// Sheet name shown in the tab list. Also the `:reset` lookup key.
    pub name: &'static str,
    /// `(row_idx, col_idx, raw_value)` — written verbatim through the
    /// engine, so formulas in lesson content also work if needed.
    pub cells: &'static [(u32, u32, &'static str)],
}

impl Lesson {
    fn apply(&self, store: &mut Store) -> Result<(), String> {
        let changes: Vec<CellChange> = self
            .cells
            .iter()
            .map(|(r, c, v)| CellChange {
                row_idx: *r,
                col_idx: *c,
                raw_value: (*v).to_string(),
                format_json: None,
            })
            .collect();
        store.apply(self.name, &changes).map_err(|e| e.to_string())
    }
}

/// Create every lesson sheet and seed its cells.
///
/// Each lesson gets an explicit `sort_order` matching its array index so
/// the tabline always reads L1 → L<n> regardless of created_at collisions.
/// `Store::list_sheets` orders by `(sort_order, created_at)` and the
/// table is `WITHOUT ROWID` (primary key = name), so ties fall back to
/// alphabetical name order — which puts "L10" before "L8". Setting
/// sort_order explicitly avoids that.
pub fn seed_tutor_db(store: &mut Store) -> Result<(), String> {
    for (idx, lesson) in LESSONS.iter().enumerate() {
        store.create_sheet(lesson.name).map_err(|e| e.to_string())?;
        store
            .conn()
            .execute(
                "UPDATE sheet SET sort_order = ? WHERE name = ?",
                rusqlite::params![idx as i64, lesson.name],
            )
            .map_err(|e| e.to_string())?;
        lesson.apply(store)?;
    }
    Ok(())
}

/// `:reset` handler — wipe and reseed the active sheet if it matches a
/// lesson. Returns false when called outside the tutor workbook.
pub fn reset_active_sheet(app: &mut App) -> bool {
    let active_name = app.active_sheet_name().to_string();
    let lesson = match LESSONS.iter().find(|l| l.name == active_name) {
        Some(l) => l,
        None => return false,
    };
    // Clear every cell first so removed steps don't linger.
    let _ = app.store.conn().execute(
        "DELETE FROM cell WHERE sheet_name = ?",
        rusqlite::params![&active_name],
    );
    let _ = lesson.apply(&mut app.store);
    app.refresh_cells();
    true
}

// ── Lesson content ────────────────────────────────────────────────────

pub const LESSONS: &[Lesson] = &[
    Lesson {

        name: "L1 Movement",
        cells: &[
            (0, 0, "L1: Movement — read this line in the formula bar above, then j to step down through the lesson."),
            (1, 0, "Press l three times. Watch the cell reference in the status bar climb from A2 → D2."),
            (2, 0, "Now press h three times to go back. h = left, l = right, j = down, k = up."),
            (3, 0, "Arrow keys do the same thing — hjkl is just the vim convention."),
            (4, 0, "Try moving to the cell labelled TARGET in column E, three rows below."),
            (7, 4, "TARGET"),
            (8, 0, "Once you reach TARGET and back, press 'gt' to advance to L2."),
            (10, 0, "Tip: 0 jumps to column A, $ jumps to the last filled cell of the row."),
        ],
    },
    Lesson {

        name: "L2 Insert",
        cells: &[
            (0, 0, "L2: Insert mode — type values into cells."),
            (1, 0, "Move to B2 (one cell right of this one). Press 'i', type 'hello', press Enter."),
            (2, 0, "i = caret at start, a = caret at end. I/A are aliases (cells are one line)."),
            (3, 0, "Press 'o' to open a new cell below; 'O' opens above. Both clear the cell first."),
            (4, 0, "Press 's' (substitute) on a non-empty cell to clear and start typing."),
            (5, 0, "Press '=' to jumpstart a formula — opens Insert with '=' already typed. (L3 covers what arrow keys do next.)"),
            (6, 0, "Esc cancels — original value preserved. Enter commits."),
            (7, 0, "Try filling B8..D8 with anything you like."),
            (9, 0, "Done? gt for L3."),
        ],
    },
    Lesson {

        name: "L3 Pointing",
        cells: &[
            (0, 0, "L3: Pointing — arrow keys insert cell references into formulas."),
            (1, 0, "Move to B5. Press '=' then '1+'. Now press Down arrow. The formula auto-completes to '=1+B6'."),
            (2, 0, "Press Down again — '=1+B7'. Press Up — '=1+B5'. The arrow moves the inserted ref."),
            (3, 0, "Press Esc, then '=', '5*'. Now Right arrow → '=5*C5'. Shift+Right extends to '=5*C5:D5'. Shift+Right again → '=5*C5:E5'."),
            (4, 0, "Range extension: plain Arrow moves the ref (single cell). Shift+Arrow extends from the anchor (range)."),
            (5, 0, "Any non-arrow key (typing, Backspace, Enter, Esc) ends pointing. Try typing '+1' after the range — pointing exits and the keystroke types normally."),
            (6, 0, "Pointing only fires at insertable positions: after '=', an operator, comma, or '('. Inside a string or after ')' it doesn't trigger."),
            (8, 0, "Done? gt for L4."),
        ],
    },
    Lesson {

        name: "L4 Counts",
        cells: &[
            (0, 0, "L4: Counts — prefix any motion with a number to repeat."),
            (1, 0, "5j moves down 5 rows. 10l moves right 10 columns. 3w jumps 3 words."),
            (2, 0, "Try 7j to land on row 9; then 4k to come back to row 5."),
            (3, 0, "5G jumps to absolute row 5 (1-indexed). G alone jumps to last filled."),
            (4, 0, "Combine: 2gt switches to sheet 2 directly. {N}gt is absolute."),
            (5, 0, "Esc clears any pending count if you change your mind."),
            (15, 0, "Row 16 — try 16G to come straight here from above."),
            (16, 0, "Done? gt for L5."),
        ],
    },
    Lesson {

        name: "L5 Words",
        cells: &[
            (0, 0, "L5: Word motions — w/b/e jump between filled-run boundaries."),
            (1, 0, "Row 3 below has runs of filled cells. Move to A3, then press w repeatedly."),
            (2, 0, "w = forward to next run start, b = back, e = end of current/next run."),
            (2, 1, "alpha"), (2, 2, "beta"), (2, 4, "gamma"), (2, 5, "delta"),
            (2, 7, "epsilon"),
            (4, 0, "From A3 press w four times. Each press skips past empties to the next filled."),
            (5, 0, "Now press b to walk back through the run starts."),
            (6, 0, "0 jumps to column A. ^ jumps to first filled cell (col B above). $ jumps to last."),
            (8, 0, "Done? gt for L6."),
        ],
    },
    Lesson {

        name: "L6 Sheet jumps",
        cells: &[
            (0, 0, "L6: Whole-sheet motions."),
            (1, 0, "gg jumps to row 1 of the current column. G jumps to last filled (or row 1000 if empty)."),
            (2, 0, "{ and } jump backwards/forwards between filled runs in the current column."),
            (3, 0, "Try this: press G to land on the bottom marker, then gg to come back."),
            (4, 0, "Then press }, }, } to walk down by paragraph."),
            (10, 0, "Run A — three filled rows in column A:"),
            (11, 0, "row 12"),
            (12, 0, "row 13"),
            (13, 0, "row 14"),
            (20, 0, "Run B — another two filled rows:"),
            (21, 0, "row 22"),
            (22, 0, "row 23"),
            (40, 0, "BOTTOM MARKER (used by the G test above)."),
            (41, 0, "Done? gt for L7."),
        ],
    },
    Lesson {

        name: "L7 Viewport",
        cells: &[
            (0, 0, "L7: Viewport scrolling without moving the cursor (or with)."),
            (1, 0, "H/M/L jump cursor to top/middle/bottom of what's visible."),
            (2, 0, "Ctrl+d / Ctrl+u scroll a half page; Ctrl+f / Ctrl+b scroll a full page."),
            (3, 0, "zz centers the cursor row vertically. zt puts cursor row at top, zb at bottom."),
            (4, 0, "zh/zl scroll the viewport one column left/right (cursor stays put)."),
            (5, 0, "Try Ctrl+d a few times, then zz to center, then gg to come back."),
            (50, 0, "Way down here at row 51 — Ctrl+u or gg gets you back."),
            (51, 0, "Done? gt for L8."),
        ],
    },
    Lesson {

        name: "L8 Visual",
        cells: &[
            (0, 0, "L8: Visual mode — select rectangles."),
            (1, 0, "Press v to start a cell-rectangle selection. Move with hjkl to extend."),
            (2, 0, "Press V for V-LINE (whole rows). Press V again to cycle to V-COLUMN (whole columns) — so VV selects the cursor's column."),
            (3, 0, "Inside Visual, 'o' swaps anchor and cursor (move the other corner)."),
            (4, 0, "y yanks the selection. d / x clears it. c clears + drops you into Insert."),
            (5, 0, "Esc returns to Normal without doing anything. gv re-selects last range."),
            (6, 0, "Try this: move to B8, press v, extend to D10 with 2j2l, then y. Then press p somewhere."),
            (7, 1, "1"), (7, 2, "2"), (7, 3, "3"),
            (8, 1, "4"), (8, 2, "5"), (8, 3, "6"),
            (9, 1, "7"), (9, 2, "8"), (9, 3, "9"),
            (12, 0, "Done? gt for L9."),
        ],
    },
    Lesson {

        name: "L9 Operators",
        cells: &[
            (0, 0, "L9: Operators — d (delete), c (change), y (yank), with motions."),
            (1, 0, "dw clears from cursor through the next word. dd clears just the cell."),
            (2, 0, "5dw clears 5 words. cw + type + Esc replaces a word."),
            (3, 0, "x clears one cell, X the cell to the left. D clears to end of row."),
            (4, 0, "y works the same: yw yanks a word, y$ yanks to end of row."),
            (5, 0, "Then p / P pastes the yanked range at the cursor."),
            (6, 0, "Practice on row 8 below: try dw on 'one', then put it back with p."),
            (7, 0, "Practice row:"),
            (7, 1, "one"), (7, 2, "two"), (7, 3, "three"), (7, 4, "four"), (7, 5, "five"),
            (10, 0, "Done? gt for L10."),
        ],
    },
    Lesson {

        name: "L10 Yank paste",
        cells: &[
            (0, 0, "L10: Yank/paste round-trip — formulas survive!"),
            (1, 0, "Move to B3. Yank the 2x2 rect with: vjly. Then move to F3 and press p."),
            (2, 0, "Notice that =B3+1 → =F3+1 — relative refs shift by the paste delta."),
            (2, 1, "10"), (2, 2, "20"),
            (3, 1, "30"), (3, 2, "=B3+1"),
            (5, 0, "yy / Y yanks the cursor cell (V4 changed it from V3 row-yank)."),
            (6, 0, "For row-yank: V (V-LINE) then y. P alongside p both paste at the cursor."),
            (10, 0, "Done? gt for L11."),
        ],
    },
    Lesson {

        name: "L11 Undo repeat",
        cells: &[
            (0, 0, "L11: Undo, redo, dot-repeat."),
            (1, 0, "u undoes the last change. Ctrl+r redoes. Ctrl+Z and Ctrl+Shift+Z also work."),
            (2, 0, "'.' (period) repeats the last change at the current cursor."),
            (3, 0, "Try: cw foo<Esc> on B5, then move to B6 and press . — same change applies."),
            (4, 0, "Or: dw on row 7, then j and . to clear next row's word."),
            (5, 1, "before"),
            (6, 1, "before"),
            (7, 0, "Practice: dw row"),
            (7, 1, "alpha"), (7, 2, "beta"), (7, 3, "gamma"),
            (10, 0, "Done? gt for L12."),
        ],
    },
    Lesson {

        name: "L12 Search",
        cells: &[
            (0, 0, "L12: Search and marks."),
            (1, 0, "Press / and type a pattern, Enter — jumps to the next match. ? searches backward."),
            (2, 0, "n / N step to the next / previous match. * / # search the cursor cell's value."),
            (3, 0, ":noh clears the highlight."),
            (4, 0, "Marks: m{a-z} stores the cursor at letter; '{a-z} jumps to that row, `{a-z} to exact cell."),
            (5, 0, "Try: press 'ma' here, move to row 30, press ''a (apostrophe a) to come back."),
            (10, 1, "needle"),
            (15, 2, "needle"),
            (20, 3, "needle"),
            (25, 4, "haystack"),
            (30, 5, "needle in a haystack"),
            (32, 0, "Try: /needle and step with n/n/n. Then press * to search the cell text. Done? gt for L13."),
        ],
    },
    Lesson {

        name: "L13 Tabs",
        cells: &[
            (0, 0, "L13: Sheets are vim's tabs."),
            (1, 0, "gt = next sheet. gT = previous. {N}gt = jump to sheet N (1-indexed)."),
            (2, 0, "From here, 1gt goes to L1, 8gt to L8, 15gt to L15."),
            (3, 0, ":tabnew Foo creates a new sheet. :tabclose drops the active one."),
            (4, 0, ":tabs (or :sheet ls) lists all sheets — current is wrapped in [brackets]."),
            (5, 0, "Try 1gt then 13gt to flip between L1 and back here."),
            (8, 0, "Done? gt for L14."),
        ],
    },
    Lesson {

        name: "L14 Ex commands",
        cells: &[
            (0, 0, "L14: The : prompt — vim's ex commands."),
            (1, 0, ":q quits. :w with no args is a no-op (vlotus writes immediately)."),
            (2, 0, ":42 jumps cursor to row 42. :A1 (or :goto A1) jumps to a cell."),
            (3, 0, ":w out.csv writes the active sheet to a CSV file (TSV via .tsv)."),
            (4, 0, ":noh clears search highlight. :help shows a one-line summary in the status."),
            (5, 0, ":reset (tutor only) reverts the active lesson to its starting state."),
            (6, 0, "Try :42 to jump down, then :1 to come back."),
            (10, 0, "Done? gt for L15."),
        ],
    },
    Lesson {

        name: "L15 Closing",
        cells: &[
            (0, 0, "L15: Closing exercise — combine everything."),
            (1, 0, "Goal: sort row 5 below alphabetically by editing cells, using only vim bindings."),
            (2, 0, "Hint: yank one cell with yy / cell-Y, move, p. Or use V-LINE + cut/paste."),
            (3, 0, "If you mess up, u undoes; :reset starts over."),
            (5, 1, "delta"),
            (5, 2, "alpha"),
            (5, 3, "echo"),
            (5, 4, "beta"),
            (5, 5, "charlie"),
            (8, 0, "When done, gt for L16 — dates and times."),
            (10, 0, "Reference: :help shows a one-line keymap reminder."),
        ],
    },
    #[cfg(feature = "datetime")]
    Lesson {
        name: "L16 Dates",
        cells: &[
            (0, 0, "L16: Dates and times — when the datetime feature is on, vlotus auto-detects ISO literals."),
            (1, 0, "Move to B2. Type 2025-04-27 and Enter. The cell renders peach — recognised as a jdate."),
            (3, 0, "Move to B4. Type =TODAY() and Enter — recomputes every load."),
            (4, 0, "Try ctrl+; (today's date) or ctrl+shift+; (current datetime) for a literal that pins."),
            (6, 0, "B7: type =B4 - DATE(2025, 1, 1). The result is a jspan, displayed friendly: 'Nd'."),
            (7, 0, "Spans render friendly ('1y 2mo 3d'); the canonical P1Y2M3D round-trips through storage."),
            (9, 0, "Accessors: =YEAR(B2), =MONTH(B2), =DAY(B2). =DATE(2026, 12, 31) constructs a jdate."),
            (10, 0, "Per-cell strftime: move to B2, then :fmt date %a %b %d → 'Sun Apr 27'. :fmt nodate to clear."),
            (12, 0, "Search: / matches against the cell's raw text, so /2025-04-27 finds typed ISO literals."),
            (14, 0, "Done? gt for L1 to wrap around. :help for the keymap."),
        ],
    },
];

#[cfg(test)]
mod tests {
    use super::*;

    fn fresh_store() -> Store {
        Store::open_in_memory().unwrap()
    }

    #[test]
    fn seed_creates_one_sheet_per_lesson() {
        let mut store = fresh_store();
        seed_tutor_db(&mut store).unwrap();
        assert_eq!(store.list_sheets().unwrap().len(), LESSONS.len());
    }

    #[test]
    fn each_lesson_has_at_least_a_title_cell() {
        let mut store = fresh_store();
        seed_tutor_db(&mut store).unwrap();
        for lesson in LESSONS {
            let snap = store.load_sheet(lesson.name).unwrap();
            assert!(
                !snap.cells.is_empty(),
                "lesson {} seeded zero cells",
                lesson.name
            );
            assert!(
                snap.cells.iter().any(|c| c.row_idx == 0 && c.col_idx == 0),
                "lesson {} missing A1",
                lesson.name
            );
        }
    }

    #[test]
    fn lesson_names_are_unique() {
        let mut seen = std::collections::HashSet::new();
        for lesson in LESSONS {
            assert!(seen.insert(lesson.name), "duplicate name: {}", lesson.name);
        }
    }
}
