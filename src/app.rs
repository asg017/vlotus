use std::collections::HashMap;
use std::fs;
use std::path::Path;
use std::sync::Arc;

use lotus_core::{
    complete_with_registry, shift_formula_refs, signature_help, CompletionKind, CompletionList,
    Registry, Sheet, SignatureHelp,
};

use crate::format;
use crate::hyperlink;
use crate::store::{
    coords::{col_idx_to_letters, from_cell_id, to_cell_id},
    CellChange, ColumnMeta, SheetMeta, Store, StoredCell, UndoGroup, UndoOp,
};

/// How many rows/columns to show in the grid.
pub const NUM_COLS: u32 = 8; // A–H
pub const NUM_ROWS: u32 = 20;

/// Default column width in characters. Used when the schema doesn't
/// have a row for the column (off-grid columns past the seeded A–H, or
/// freshly-inserted columns).
pub const DEFAULT_COL_WIDTH: u16 = 12;
/// Min/max user-settable column width. Below 1 there's no content room;
/// above ~80 the column dominates the viewport.
pub const MIN_COL_WIDTH: u16 = 1;
pub const MAX_COL_WIDTH: u16 = 80;

/// Last addressable row/col index. The grid spans 0..=MAX_ROW × 0..=MAX_COL.
pub const MAX_ROW: u32 = 999;
pub const MAX_COL: u32 = 25; // A–Z

/// Top-level interaction mode. Mutually exclusive: at most one of
/// edit / command / visual / search / picker is active at a time.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mode {
    /// Default — vim Normal mode.
    Nav,
    /// In-cell editing — `edit_buf` holds the staged value.
    Edit,
    /// `:` command line — `edit_buf` holds the command being typed.
    Command,
    /// `!` shell prompt — `edit_buf` holds the shell command being
    /// typed. Enter runs it via `shell::run`, sniffs the stdout, and
    /// pastes the result at the cursor.
    Shell,
    /// Vim Visual mode — motions extend the selection rectangle.
    Visual(VisualKind),
    /// Vim `/` or `?` search prompt — `edit_buf` holds the pattern.
    Search(SearchDir),
    /// Modal color-swatch picker overlaying the grid. State (target
    /// axis, grid cursor, hex-input mode) lives on `App::color_picker`.
    ColorPicker,
    /// Modal popup rendering the active patch's changeset. State
    /// (rendered lines + scroll offset) lives on `App::patch_show`.
    PatchShow,
}

#[derive(Debug, Clone)]
pub struct PatchShowState {
    pub lines: Vec<String>,
    pub scroll: usize,
}

/// Which color axis a `Mode::ColorPicker` session is editing.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorPickerKind {
    Fg,
    Bg,
}

/// Live state for the color-swatch picker. Held on `App::color_picker`
/// only while `Mode::ColorPicker` is active.
#[derive(Debug, Clone)]
pub struct ColorPickerState {
    pub kind: ColorPickerKind,
    /// Index into `format::COLOR_PRESETS` of the highlighted swatch.
    pub cursor: usize,
    /// `Some(buffer)` when the user toggled into hex-entry mode (`?`).
    pub hex_input: Option<String>,
    /// Selection rect captured at picker-open time so the apply step
    /// targets the same cells even if the user changed the cursor.
    pub target_rect: (u32, u32, u32, u32),
}

/// Direction of an active or in-progress search.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SearchDir {
    Forward,
    Backward,
}

/// Sub-flavour of Visual mode. `Cell` is vim's character/block visual
/// (a free rectangle); `Row` is vim's line-visual (every column is
/// forced into the selection regardless of cursor x); `Column` is the
/// orthogonal counterpart (every row is forced in regardless of cursor
/// y) — used by mouse column-header click/drag, no keyboard binding yet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VisualKind {
    Cell,
    Row,
    Column,
}

/// Result of executing a `:` command line — drives the run-loop.
///
/// `Quit { force }` carries vim's bang-suffix: `:q!` / `:wq!` / `:x!` set
/// `force = true` to bypass the unsaved-changes / in-memory-session
/// guards in `App::run_command`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CommandOutcome {
    Continue,
    Quit { force: bool },
}

/// A destructive structural op that's been queued and is waiting for the
/// user's `y` to commit. See [sheet.delete.row-confirm] /
/// [sheet.delete.column-confirm].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PendingConfirm {
    DeleteRow(u32),
    DeleteCol(u32),
}

/// Direction for inserts — left/above of the cursor (the new row/col
/// takes the cursor's index) or right/below (the cursor's row/col stays
/// where it is and a blank one slots in next door). See
/// [sheet.column.insert-left-right].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InsertSide {
    AboveOrLeft,
    BelowOrRight,
}

/// Whether a clipboard mark represents a copy or a deferred cut. On a
/// subsequent paste, `Cut` clears the source range; `Copy` leaves it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ClipMarkMode {
    Copy,
    Cut,
}

/// The "marching ants" rectangle painted around a copied or cut range.
/// Cleared by Escape, a fresh copy/cut, a sheet switch, or a paste that
/// consumes a `Cut`.
#[derive(Debug, Clone, Copy)]
pub struct ClipboardMark {
    pub r1: u32,
    pub c1: u32,
    pub r2: u32,
    pub c2: u32,
    pub mode: ClipMarkMode,
}

/// One cell parsed out of a pasted clipboard grid.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PastedCell {
    /// The displayed value (computed). Used as the raw value when the
    /// payload didn't carry a formula marker.
    pub value: String,
    /// Original raw formula text (`=A1*2`). Present only on intra-app
    /// pastes that round-tripped via the `data-sheets-formula` marker.
    pub formula: Option<String>,
}

/// A 2D clipboard grid plus optional source-anchor for intra-app
/// formula round-tripping ([sheet.clipboard.paste-formula-shift]).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PastedGrid {
    /// Top-left of the source range when the payload originated from a
    /// vlotus copy. None for external clipboards (Excel, Sheets, etc.).
    pub source_anchor: Option<(u32, u32)>,
    pub cells: Vec<Vec<PastedCell>>,
}

impl PastedGrid {
    pub fn rows(&self) -> u32 {
        self.cells.len() as u32
    }

    pub fn cols(&self) -> u32 {
        self.cells.iter().map(|r| r.len()).max().unwrap_or(0) as u32
    }

    pub fn is_single_cell(&self) -> bool {
        self.rows() == 1 && self.cols() == 1
    }
}

/// "Pointing" state used by [sheet.editing.formula-ref-pointing].
/// While this is `Some`, an arrow key moves the inserted reference
/// (replaces the token in `edit_buf` with one pointing at the new
/// target) instead of moving the edit caret. Any non-arrow key
/// returns the editor to plain text editing.
/// What kind of ref a [`PointingState`] is tracking. Drives the
/// rewrite logic in `App::rewrite_pointing_text` so a column-axis drag
/// produces `B:E` and a row-axis drag produces `1:5` while a cell drag
/// produces `B2:E5`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PointingKind {
    /// Single cell or cell-range ref (`B2`, `B2:E5`).
    Cell,
    /// Whole-column or column-range ref (`B:B`, `B:E`). Row fields of
    /// `PointingState` are ignored.
    Column,
    /// Whole-row or row-range ref (`1:1`, `1:5`). Column fields of
    /// `PointingState` are ignored.
    Row,
}

#[derive(Debug, Clone, Copy)]
pub struct PointingState {
    /// Byte range in `edit_buf` covered by the inserted ref token.
    pub start: usize,
    pub end: usize,
    /// What kind of ref this span represents — drives the rewrite
    /// formatting (Cell uses `to_cell_id`, Column uses letters, Row
    /// uses 1-indexed numbers).
    pub kind: PointingKind,
    /// Anchor cell/col/row of the ref. Equals `target_*` for a single-
    /// extent ref; during a mouse drag this stays at the original click
    /// point while `target_*` advances, producing range syntax.
    pub anchor_row: u32,
    pub anchor_col: u32,
    /// "Other end" of the ref. For single-cell refs this equals the
    /// anchor; for ranges it's the dragged-to position.
    pub target_row: u32,
    pub target_col: u32,
}

/// State of the formula-editor's completion popup.
/// See [sheet.editing.formula-autocomplete].
#[derive(Debug, Clone)]
pub struct AutocompleteState {
    pub list: CompletionList,
    /// 0-based index into `list.items`. Always in range when non-empty.
    pub selected: usize,
    /// Whether the user has explicitly navigated this popup (Up/Down/
    /// Ctrl+n/Ctrl+p). Until then `selected` is just the auto-default
    /// (0), so Enter should commit the cell rather than inserting the
    /// highlighted item — only Tab / Ctrl+y always accept.
    pub user_selected: bool,
}

/// Vim's internal yank register. Populated by `y` (Visual) and `yy`/`Y`
/// (Normal). `linewise` records whether the source was V-LINE / `yy`;
/// V4 will use it to differentiate `p` vs `P` paste positions, V3 just
/// preserves it for that future use.
#[derive(Debug, Clone)]
pub struct YankRegister {
    pub grid: PastedGrid,
    #[allow(dead_code)] // consumed in a later vim phase
    pub linewise: bool,
}

/// Vim's operator family — the verb half of `{op}{motion}` grammar.
/// `Delete` clears, `Change` clears+Inserts, `Yank` only captures.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Operator {
    Delete,
    Change,
    Yank,
}

/// What `last_edit` actually does when `.` replays it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EditKind {
    /// Clear a rect (no insert).
    Delete,
    /// Clear a rect, then write `text` into the rect's top-left cell.
    Change,
    /// Overwrite the cursor cell with `text` (no clear of any other cells).
    Insert,
}

/// State of an active search. Cleared by `:noh` or by starting a new
/// search. While `Some`, matching cells are tinted in the grid.
#[derive(Debug, Clone)]
pub struct SearchState {
    pub pattern: String,
    pub direction: SearchDir,
    pub case_insensitive: bool,
}

impl SearchState {
    /// Whether `cell_text` contains the pattern under the active casing
    /// rule. Used by both the n/N stepper and the grid renderer.
    pub fn matches(&self, cell_text: &str) -> bool {
        if self.pattern.is_empty() {
            return false;
        }
        if self.case_insensitive {
            cell_text
                .to_lowercase()
                .contains(&self.pattern.to_lowercase())
        } else {
            cell_text.contains(&self.pattern)
        }
    }
}

/// Vim `m` / `'` / `` ` `` are two-keystroke commands. The first key sets
/// this flag; the next letter resolves it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MarkAction {
    /// `m{a-z}` — store the cursor at this letter.
    Set,
    /// `` `{a-z} `` — jump to the marked cell exactly.
    JumpExact,
    /// `'{a-z}` — jump to the marked row, col 0.
    JumpRow,
}

/// The most-recent change that `.` should replay. Captures rect shape
/// (relative to the action's cursor) plus typed text where applicable.
/// Yank operations are not recorded — vim's `.` ignores them.
#[derive(Debug, Clone)]
pub struct EditAction {
    pub kind: EditKind,
    /// Rect anchor offset from the cursor at action time.
    pub anchor_dr: i32,
    pub anchor_dc: i32,
    /// Rect dimensions (≥1).
    pub rect_rows: u32,
    pub rect_cols: u32,
    /// Committed text — populated by Insert and by Change after the user
    /// confirms the edit. Change records with `None` mean the user Esc'd
    /// out before committing; replay just re-clears the rect.
    pub text: Option<String>,
}

/// Snapshot of the most-recent Visual selection so `gv` can re-enter it.
#[derive(Debug, Clone, Copy)]
pub struct LastVisual {
    pub anchor: (u32, u32),
    pub cursor: (u32, u32),
    pub kind: VisualKind,
}

impl ClipboardMark {
    pub fn contains(&self, row: u32, col: u32) -> bool {
        row >= self.r1 && row <= self.r2 && col >= self.c1 && col <= self.c2
    }

    pub fn on_perimeter(&self, row: u32, col: u32) -> bool {
        self.contains(row, col)
            && (row == self.r1 || row == self.r2 || col == self.c1 || col == self.c2)
    }
}

pub struct App {
    pub store: Store,
    /// Every sheet in the workbook, in display order. Always non-empty —
    /// `App::new` creates a default sheet if the database had none.
    pub sheets: Vec<SheetMeta>,
    /// Index into `sheets` of the currently-active tab.
    pub active_sheet: usize,
    pub db_label: String,

    /// 0-based cursor position.
    pub cursor_row: u32,
    pub cursor_col: u32,

    /// Scroll offset (first visible row/col).
    pub scroll_row: u32,
    pub scroll_col: u32,

    /// Number of visible rows/columns (set by UI each frame).
    pub visible_rows: u32,
    pub visible_cols: u32,

    /// Current top-level mode (Nav / Edit / Command).
    pub mode: Mode,
    /// Text buffer used by both Edit and Command modes (cell edit value
    /// or `:` command line, depending on `mode`).
    pub edit_buf: String,
    /// Cursor position within `edit_buf` (byte offset, always on a char
    /// boundary).
    pub edit_cursor: usize,

    /// Cached cell data from SQLite (refreshed after each mutation).
    pub cells: Vec<StoredCell>,
    /// Cached column metadata for the active sheet (per-column width,
    /// name). Refreshed alongside `cells` after mutations.
    pub columns: Vec<ColumnMeta>,

    /// Status message shown at the bottom.
    pub status: String,

    /// Anchor for range selection (set when Shift+Arrow starts a selection).
    /// When Some, the selection spans from anchor to cursor (inclusive).
    pub selection_anchor: Option<(u32, u32)>,

    /// "Marching ants" rectangle painted around a recently-copied or
    /// recently-cut range. See [sheet.clipboard.mark-visual].
    pub clipboard_mark: Option<ClipboardMark>,

    /// Redo history. Popped undo groups are pushed here so `Ctrl+r`
    /// can replay them forward. Cleared on any new mutation.
    /// (The undo side lives on disk in `undo_entry` — see
    /// `Store::pop_undo` / `Store::apply_redo`.)
    pub redo_stack: Vec<UndoGroup>,

    /// Live completion popup state during cell editing.
    /// See [sheet.editing.formula-autocomplete].
    pub autocomplete: Option<AutocompleteState>,
    /// Live signature-help tooltip state during cell editing.
    /// See [sheet.editing.formula-signature-help].
    pub signature: Option<SignatureHelp>,
    /// Long-lived registry handle used by the autocomplete popup so
    /// runtime-registered custom functions (HYPERLINK + the datetime
    /// extension) appear alongside builtins. Built once in `App::new`
    /// from the same `register` calls that `Store::recalculate` uses,
    /// so the popup and the engine see the same function set.
    pub registry: Arc<Registry>,
    /// Active formula-ref-pointing session, if any.
    /// See [sheet.editing.formula-ref-pointing].
    pub pointing: Option<PointingState>,
    /// A queued destructive op waiting for `y` to confirm.
    pub pending_confirm: Option<PendingConfirm>,

    /// Vim count prefix (`5j`, `10G`). Accumulates digit presses; consumed
    /// by the next motion. `None` means "no count" (defaults to 1 for
    /// repeated motions, or absolute target for `gg`/`G`).
    pub pending_count: Option<u32>,
    /// `g` was just pressed; the next key resolves a `g`-prefix command
    /// (`gg`, `gv`, `gt`, `gT`). Cleared after one keypress.
    pub pending_g: bool,
    /// `z` was just pressed; the next key resolves a `z`-prefix command
    /// (`zz`, `zt`, `zb`, `zh`, `zl`).
    pub pending_z: bool,
    /// `f` was just pressed; the next key resolves an `f`-prefix
    /// (format) command. The consumer is placed before the
    /// operator/Nav-only block in `handle_nav_key` so lowercase
    /// second-keys like `fc` reach it instead of being absorbed
    /// by the `c` Change operator.
    pub pending_f: bool,

    /// Vim's unnamed yank register. `y`/`yy`/`Y` populate it, `p`/`P`
    /// drain it. Falls back to the OS clipboard when None.
    pub yank_register: Option<YankRegister>,
    /// Most-recent Visual selection, restored by `gv`.
    pub last_visual: Option<LastVisual>,

    /// Pending vim operator. While `Some`, the next motion or doubled
    /// operator-letter resolves it.
    pub pending_operator: Option<Operator>,
    /// Count captured at the moment the operator key was pressed (e.g.
    /// `5dw` → `Some(5)`). `None` is "no count specified" — distinct from
    /// `Some(1)` because some motions (`G`) treat them differently.
    pub pending_op_count: Option<u32>,

    /// Most-recent recordable change for `.` (V6). Yank doesn't populate
    /// this. Operators set rect-shape immediately; Change actions get
    /// their text populated on Insert commit.
    pub last_edit: Option<EditAction>,

    /// V5: active search state. Highlight is painted in `draw_grid`
    /// while `Some`; cleared by `:noh`.
    pub search: Option<SearchState>,

    /// Live state for `Mode::ColorPicker`. Some only while the picker
    /// is open; `Esc` / `Enter` close it.
    pub color_picker: Option<ColorPickerState>,
    pub patch_show: Option<PatchShowState>,
    /// V5: per-letter marks (`a-z`). Set by `m{letter}` and consumed by
    /// `'{letter}` (jump to row) or `` `{letter} `` (jump to exact cell).
    pub marks: HashMap<char, (u32, u32)>,
    /// Two-keystroke `m`/`'`/`` ` `` — the first key sets this flag, the
    /// second supplies the mark letter.
    pub pending_mark: Option<MarkAction>,

    /// `Ctrl-w` was just pressed; the next key resolves a `Ctrl-w`
    /// prefix command (currently `>` / `<` to grow/shrink the active
    /// column).
    pub pending_ctrl_w: bool,

    /// `:set mouse` toggle. **On by default** — clicking, dragging, and
    /// scroll-wheel "just work" matching Excel / VisiData / sc-im. When
    /// the user wants native terminal text-selection (Cmd/Option+drag →
    /// copy) back, `:set nomouse` releases capture for the rest of the
    /// session. The actual `EnableMouseCapture` / `DisableMouseCapture`
    /// is driven by `run_loop`, which diffs this flag against the live
    /// terminal state each tick.
    pub mouse_enabled: bool,

    /// Cell where the most recent left-button mouse press happened.
    /// Set on `Down(Left)`, cleared on `Up(Left)`. Used by the drag
    /// handler (T2) to know where a drag-to-visual gesture started so
    /// it can anchor the new selection at the original click point.
    /// Separate from `selection_anchor` — that one stays Visual-mode
    /// only; this one is mouse-only state that exists during a press.
    pub drag_anchor: Option<(u32, u32)>,

    /// `(when, cell)` of the most recent left-click. Used by the
    /// double-click handler (T4): a second `Down(Left)` on the same
    /// cell within `DOUBLE_CLICK_MS` enters Edit mode.
    pub last_click: Option<(std::time::Instant, (u32, u32))>,

    /// True once the user has performed any sheet mutation in this
    /// session (cell edit, structural op, undo/redo, column resize).
    /// Drives the in-memory `:q` warning — read-only browsing of an
    /// `:memory:` workbook should not nag on quit. Never reset within
    /// a session; once touched, stays touched.
    pub touched: bool,

    /// True when there are pending writes the user has not flushed.
    /// Always `false` in the default sqlite-backed mode (writes are
    /// immediate, see `apply_changes_recorded`). Wired to the dirty-row
    /// tracking introduced by the CSV mode epic (T2 `1vdapj3i`); `:w`
    /// in that mode flips it back to `false`. Backs `has_unsaved_changes`.
    pub dirty: bool,
}

impl App {
    pub fn new(mut store: Store, db_label: &str) -> Self {
        let sheets = ensure_at_least_one_sheet(&mut store);
        // Bootstrap (default Sheet1 creation) flips the dirty flag.
        // Commit it so a freshly-opened workbook isn't reported as
        // having unsaved changes.
        let _ = store.commit();
        let active_sheet = 0;
        let cells = store
            .load_sheet(&sheets[active_sheet].name)
            .map(|s| s.cells)
            .unwrap_or_default();
        let columns = store
            .load_columns(&sheets[active_sheet].name)
            .unwrap_or_default();

        // Build a scratch sheet purely to materialise the registry the
        // autocomplete popup will consult — same registrations as
        // `Store::recalculate`, so the popup and the engine surface the
        // same function set. Held as an Arc so swapping it onto a fresh
        // recalc Sheet later is allocation-free if we ever want to.
        let registry = build_registry();

        App {
            store,
            sheets,
            active_sheet,
            db_label: db_label.to_string(),
            cursor_row: 0,
            cursor_col: 0,
            scroll_row: 0,
            scroll_col: 0,
            visible_rows: NUM_ROWS,
            visible_cols: NUM_COLS,
            mode: Mode::Nav,
            edit_buf: String::new(),
            edit_cursor: 0,
            cells,
            columns,
            status: String::new(),
            selection_anchor: None,
            clipboard_mark: None,
            redo_stack: Vec::new(),
            autocomplete: None,
            signature: None,
            registry,
            pointing: None,
            pending_confirm: None,
            pending_count: None,
            pending_g: false,
            pending_z: false,
            pending_f: false,
            yank_register: None,
            last_visual: None,
            pending_operator: None,
            pending_op_count: None,
            last_edit: None,
            search: None,
            color_picker: None,
            patch_show: None,
            marks: HashMap::new(),
            pending_mark: None,
            pending_ctrl_w: false,
            mouse_enabled: true,
            drag_anchor: None,
            last_click: None,
            touched: false,
            dirty: false,
        }
    }

    /// True when there are pending writes the user has not flushed.
    /// `Store::is_dirty` covers the SQLite path (set on every cell /
    /// sheet / column / structural mutation, cleared by `:w`). The
    /// `App::dirty` flag is reserved for CSV mode (T2 `1vdapj3i`)
    /// which tracks dirty rows independently of the txn.
    pub fn has_unsaved_changes(&self) -> bool {
        self.dirty || self.store.is_dirty()
    }

    /// `:q` warning gate for an `:memory:` workbook. Tutor sessions
    /// (`db_label == "tutor"`) are intentionally ephemeral and skip
    /// the warning, matching `run_tutor`'s "curriculum stays pristine"
    /// design (`main.rs`).
    pub fn should_warn_in_memory(&self) -> bool {
        self.db_label == ":memory:" && self.touched
    }

    /// Mark the session as having mutations the user might want to
    /// save. Called from every state-mutating App method. One-way —
    /// stays `true` for the rest of the session (matches "no write
    /// since last change" semantics; we don't model "undo back to a
    /// clean state").
    fn mark_touched(&mut self) {
        self.touched = true;
    }

    /// Take `pending_count`, defaulting to 1. Used by repeat-multiplied
    /// motions (`hjkl`, `w`, `b`, `e`, …) — for absolute motions like
    /// `G`/`gg` where None vs Some(n) means different things, take the
    /// raw `pending_count` instead.
    pub fn consume_count(&mut self) -> u32 {
        self.pending_count.take().unwrap_or(1)
    }

    /// Drop every pending vim-prefix flag. Called from Esc and from
    /// keystrokes that abort an in-progress sequence (e.g. an unknown
    /// follow-up to `g`).
    pub fn clear_pending_motion_state(&mut self) {
        self.pending_count = None;
        self.pending_g = false;
        self.pending_z = false;
        self.pending_f = false;
        self.pending_operator = None;
        self.pending_op_count = None;
        self.pending_mark = None;
        self.pending_ctrl_w = false;
    }

    // ── Visual mode (V3) ─────────────────────────────────────────────

    /// Enter Visual mode at the current cursor. The cursor is the new
    /// anchor — single-cell selection initially, expands as motions fire.
    pub fn enter_visual(&mut self, kind: VisualKind) {
        self.mode = Mode::Visual(kind);
        self.selection_anchor = Some((self.cursor_row, self.cursor_col));
        self.status = match kind {
            VisualKind::Cell => "VISUAL — y/d/c/o/Esc, motions extend".into(),
            VisualKind::Row => "V-LINE — y/d/c/Esc".into(),
            VisualKind::Column => "V-COLUMN — y/d/c/Esc".into(),
        };
    }

    /// Promote the current Cell-visual to Row-visual or vice versa.
    /// Anchor stays put.
    pub fn switch_visual_kind(&mut self, kind: VisualKind) {
        if matches!(self.mode, Mode::Visual(_)) {
            self.mode = Mode::Visual(kind);
            self.status = match kind {
                VisualKind::Cell => "VISUAL".into(),
                VisualKind::Row => "V-LINE".into(),
                VisualKind::Column => "V-COLUMN".into(),
            };
        }
    }

    /// Drop Visual mode, save the rectangle for `gv`, and clear the
    /// selection. Returns to Normal.
    pub fn exit_visual(&mut self) {
        if let Mode::Visual(kind) = self.mode {
            if let Some(anchor) = self.selection_anchor {
                self.last_visual = Some(LastVisual {
                    anchor,
                    cursor: (self.cursor_row, self.cursor_col),
                    kind,
                });
            }
        }
        self.mode = Mode::Nav;
        self.selection_anchor = None;
        self.status.clear();
    }

    /// Vim `gv`: re-enter the most-recent Visual selection. Returns true
    /// when a saved selection existed.
    pub fn reselect_last_visual(&mut self) -> bool {
        let Some(last) = self.last_visual else {
            return false;
        };
        self.mode = Mode::Visual(last.kind);
        self.selection_anchor = Some(last.anchor);
        self.cursor_row = last.cursor.0.min(MAX_ROW);
        self.cursor_col = last.cursor.1.min(MAX_COL);
        self.scroll_to_cursor();
        self.status = match last.kind {
            VisualKind::Cell => "VISUAL (re-select)".into(),
            VisualKind::Row => "V-LINE (re-select)".into(),
            VisualKind::Column => "V-COLUMN (re-select)".into(),
        };
        true
    }

    /// Vim Visual `o`: swap anchor and cursor so the other corner of the
    /// rectangle becomes the moveable end.
    pub fn swap_visual_corners(&mut self) {
        if let Some(anchor) = self.selection_anchor {
            let new_anchor = (self.cursor_row, self.cursor_col);
            self.cursor_row = anchor.0;
            self.cursor_col = anchor.1;
            self.selection_anchor = Some(new_anchor);
            self.scroll_to_cursor();
        }
    }

    // ── Yank / paste register (V3) ───────────────────────────────────

    /// Build a `PastedGrid` over the given inclusive rect, preserving
    /// formulas verbatim so the round-trip through `apply_pasted_grid`
    /// shifts refs correctly. Used by V3 visual yank and V4 operators.
    pub fn build_grid_over(&self, r1: u32, c1: u32, r2: u32, c2: u32) -> PastedGrid {
        let mut cells = Vec::with_capacity((r2 - r1 + 1) as usize);
        for r in r1..=r2 {
            let mut row = Vec::with_capacity((c2 - c1 + 1) as usize);
            for c in c1..=c2 {
                let raw = self.get_raw(r, c);
                let value = self.get_display(r, c);
                let formula = if raw.starts_with('=') { Some(raw) } else { None };
                row.push(PastedCell { value, formula });
            }
            cells.push(row);
        }
        PastedGrid {
            source_anchor: Some((r1, c1)),
            cells,
        }
    }

    /// Vim `y` (Visual): yank the current selection (or cursor cell) into
    /// the unnamed register. `linewise` is set when called from V-LINE.
    pub fn yank_selection(&mut self, linewise: bool) -> (u32, u32, u32, u32) {
        let (r1, c1, r2, c2) = self.effective_rect();
        let grid = self.build_grid_over(r1, c1, r2, c2);
        self.yank_register = Some(YankRegister { grid, linewise });
        (r1, c1, r2, c2)
    }

    /// Vim `yy` / `Y`: yank the cursor's row across all columns.
    pub fn yank_row(&mut self) {
        let row = self.cursor_row;
        let grid = self.build_grid_over(row, 0, row, MAX_COL);
        self.yank_register = Some(YankRegister {
            grid,
            linewise: true,
        });
    }

    /// Vim `p` / `P`: drop the unnamed register at the cursor, preserving
    /// formula round-tripping. Returns the dimensions written, or
    /// `Ok((0, 0))` when the register is empty.
    pub fn paste_from_register(&mut self) -> Result<(u32, u32), String> {
        let Some(reg) = self.yank_register.clone() else {
            return Ok((0, 0));
        };
        self.apply_pasted_grid(&reg.grid, None)
    }

    // ── Search & marks (V5) ──────────────────────────────────────────

    /// Open the `/` or `?` search prompt with an empty pattern.
    pub fn start_search(&mut self, direction: SearchDir) {
        self.mode = Mode::Search(direction);
        self.edit_buf.clear();
        self.edit_cursor = 0;
        self.status.clear();
    }

    /// Cancel the search prompt without committing a pattern.
    pub fn cancel_search(&mut self) {
        self.mode = Mode::Nav;
        self.edit_buf.clear();
        self.edit_cursor = 0;
    }

    /// Commit the pattern in `edit_buf` and jump to the first match. Sets
    /// `search` so n/N can step through subsequent matches.
    pub fn commit_search(&mut self, direction: SearchDir) {
        let pattern = self.edit_buf.trim().to_string();
        self.mode = Mode::Nav;
        self.edit_buf.clear();
        self.edit_cursor = 0;
        if pattern.is_empty() {
            self.search = None;
            return;
        }
        self.search = Some(SearchState {
            pattern,
            direction,
            // Smart-case: lowercase pattern → case insensitive, else strict.
            case_insensitive: true,
        });
        self.search_step(direction);
    }

    /// Vim `n`/`N`: step to the next/prev match in the search direction
    /// (`n`) or its reverse (`N`).
    pub fn search_step(&mut self, dir: SearchDir) {
        let Some(state) = self.search.as_ref() else {
            self.status = "No search pattern".into();
            return;
        };
        let pattern = state.pattern.clone();
        let case_insensitive = state.case_insensitive;
        let effective_dir = match (state.direction, dir) {
            (SearchDir::Forward, SearchDir::Forward) => SearchDir::Forward,
            (SearchDir::Backward, SearchDir::Backward) => SearchDir::Forward,
            _ => SearchDir::Backward,
        };
        let from = (self.cursor_row, self.cursor_col);
        match self.find_match(from, &pattern, case_insensitive, effective_dir) {
            Some((r, c)) => {
                self.jump_cursor_to(r, c);
                self.status = format!("/{pattern}");
            }
            None => {
                self.status = format!("Pattern not found: {pattern}");
            }
        }
    }

    /// Linear scan for the next/prev cell whose computed value contains
    /// `pattern`. Wraps around at the grid edge.
    fn find_match(
        &self,
        from: (u32, u32),
        pattern: &str,
        case_insensitive: bool,
        dir: SearchDir,
    ) -> Option<(u32, u32)> {
        let needle = if case_insensitive {
            pattern.to_lowercase()
        } else {
            pattern.to_string()
        };
        let test = |cell: &StoredCell| -> bool {
            let display = cell.computed_value.as_deref().unwrap_or("");
            let hay = if case_insensitive {
                display.to_lowercase()
            } else {
                display.to_string()
            };
            hay.contains(&needle)
        };

        let mut hits: Vec<(u32, u32)> = self
            .cells
            .iter()
            .filter(|c| test(c))
            .map(|c| (c.row_idx, c.col_idx))
            .collect();
        if hits.is_empty() {
            return None;
        }
        hits.sort();
        match dir {
            SearchDir::Forward => hits
                .iter()
                .find(|p| **p > from)
                .copied()
                .or_else(|| hits.first().copied()),
            SearchDir::Backward => hits
                .iter()
                .rev()
                .find(|p| **p < from)
                .copied()
                .or_else(|| hits.last().copied()),
        }
    }

    /// Vim `*` / `#`: search for the cursor cell's exact computed value.
    pub fn search_current_cell(&mut self, dir: SearchDir) {
        let value = self.get_display(self.cursor_row, self.cursor_col);
        if value.trim().is_empty() {
            self.status = "Empty cell".into();
            return;
        }
        self.search = Some(SearchState {
            pattern: value,
            direction: dir,
            case_insensitive: false,
        });
        self.search_step(dir);
    }

    /// Vim `m{a-z}`: stash the cursor at the given letter.
    pub fn set_mark(&mut self, letter: char) {
        if !letter.is_ascii_alphabetic() {
            self.status = format!("Invalid mark: {letter}");
            return;
        }
        self.marks
            .insert(letter.to_ascii_lowercase(), (self.cursor_row, self.cursor_col));
        self.status = format!("Marked '{letter}");
    }

    /// Vim `` `{a-z} `` (exact cell) / `'{a-z}` (row only).
    pub fn jump_to_mark(&mut self, letter: char, row_only: bool) {
        if !letter.is_ascii_alphabetic() {
            self.status = format!("Invalid mark: {letter}");
            return;
        }
        match self.marks.get(&letter.to_ascii_lowercase()).copied() {
            Some((r, c)) => {
                let target_col = if row_only { 0 } else { c };
                self.jump_cursor_to(r, target_col);
            }
            None => self.status = format!("Mark not set: '{letter}"),
        }
    }

    /// Vim `:noh`: clear the active search highlight.
    pub fn clear_search(&mut self) {
        if self.search.take().is_some() {
            self.status = "Cleared search".into();
        }
    }

    /// Resolve `:colwidth <arg>` (numeric or `auto`) for the given column.
    /// Updates status; the dispatcher returns Continue regardless.
    pub fn handle_colwidth(&mut self, col_idx: u32, arg: &str) {
        let letter: String = to_cell_id(0, col_idx)
            .chars()
            .take_while(|c| c.is_ascii_alphabetic())
            .collect();
        if arg.eq_ignore_ascii_case("auto") {
            match self.autofit_column(col_idx) {
                Ok(w) => self.status = format!("Column {letter} → {w}"),
                Err(e) => self.status = format!("Error: {e}"),
            }
            return;
        }
        match arg.parse::<u16>() {
            Ok(n) if (MIN_COL_WIDTH..=MAX_COL_WIDTH).contains(&n) => {
                match self.set_column_width(col_idx, n) {
                    Ok(()) => self.status = format!("Column {letter} → {n}"),
                    Err(e) => self.status = format!("Error: {e}"),
                }
            }
            Ok(_) => {
                self.status = format!(
                    "Width must be between {MIN_COL_WIDTH} and {MAX_COL_WIDTH}"
                );
            }
            Err(_) => self.status = format!("Bad width: {arg}"),
        }
    }

    /// V8: parse `:goto A1` / `:42` / etc. into a cursor jump. Returns
    /// false on an unrecognised target.
    pub fn jump_to_target(&mut self, target: &str) -> bool {
        if let Ok(row) = target.parse::<u32>() {
            self.goto_row(row.saturating_sub(1));
            return true;
        }
        if let Some((r, c)) = from_cell_id(&target.to_uppercase()) {
            self.jump_cursor_to(r, c);
            return true;
        }
        false
    }

    /// V8 `:w <path>`: write the active sheet to disk. Format dispatched
    /// by extension. Empty rows are skipped to keep output compact.
    pub fn export_sheet(&self, path: &str) -> Result<(), String> {
        let ext = Path::new(path)
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_ascii_lowercase();
        let separator = match ext.as_str() {
            "tsv" => '\t',
            "csv" => ',',
            other => return Err(format!("Unknown export format: .{other}")),
        };
        let max_row = self.cells.iter().map(|c| c.row_idx).max();
        let max_col = self.cells.iter().map(|c| c.col_idx).max();
        let (last_row, last_col) = match (max_row, max_col) {
            (Some(r), Some(c)) => (r, c),
            _ => {
                fs::write(path, "").map_err(|e| e.to_string())?;
                return Ok(());
            }
        };
        let mut out = String::new();
        for r in 0..=last_row {
            let mut row_cells = Vec::with_capacity((last_col + 1) as usize);
            for c in 0..=last_col {
                let value = self.get_display(r, c);
                row_cells.push(quote_for_format(&value, separator));
            }
            out.push_str(&row_cells.join(&separator.to_string()));
            out.push('\n');
        }
        fs::write(path, out).map_err(|e| e.to_string())?;
        Ok(())
    }

    /// Vim `.` (V6): replay `last_edit` at the current cursor. Returns
    /// true when an action was replayed.
    pub fn repeat_last_edit(&mut self) -> bool {
        let Some(action) = self.last_edit.clone() else {
            return false;
        };
        let r1 = (self.cursor_row as i32 + action.anchor_dr).max(0) as u32;
        let c1 = (self.cursor_col as i32 + action.anchor_dc).max(0) as u32;
        let r2 = (r1 + action.rect_rows.saturating_sub(1)).min(MAX_ROW);
        let c2 = (c1 + action.rect_cols.saturating_sub(1)).min(MAX_COL);

        match action.kind {
            EditKind::Delete => {
                let _ = self.clear_rect(r1, c1, r2, c2);
                self.cursor_row = r1;
                self.cursor_col = c1;
                self.status = "Repeated".into();
            }
            EditKind::Change => {
                let _ = self.clear_rect(r1, c1, r2, c2);
                if let Some(text) = action.text.as_deref() {
                    let _ = self.apply_changes_recorded(&[Self::cell_change_from_typed(
                        r1, c1, text,
                    )]);
                }
                self.cursor_row = r1;
                self.cursor_col = c1;
                self.status = "Repeated".into();
            }
            EditKind::Insert => {
                let text = action.text.clone().unwrap_or_default();
                let _ = self.apply_changes_recorded(&[Self::cell_change_from_typed(
                    self.cursor_row,
                    self.cursor_col,
                    &text,
                )]);
                self.status = "Repeated".into();
            }
        }
        true
    }

    /// Clear every cell in the inclusive `(r1, c1)`–`(r2, c2)` rectangle as
    /// a single undo entry. Used by Visual `d` / `c` and (later) `dd`.
    pub fn clear_rect(&mut self, r1: u32, c1: u32, r2: u32, c2: u32) -> Result<(), String> {
        let mut changes = Vec::with_capacity(((r2 - r1 + 1) * (c2 - c1 + 1)) as usize);
        for r in r1..=r2 {
            for c in c1..=c2 {
                changes.push(CellChange {
                    row_idx: r,
                    col_idx: c,
                    raw_value: String::new(),
                    format_json: None,
                });
            }
        }
        self.apply_changes_recorded(&changes)
    }

    /// Convenience accessor — every code path used to reach into
    /// `self.sheet_id`; now they go through here so the active sheet
    /// can change at runtime.
    pub fn active_sheet_name(&self) -> &str {
        &self.sheets[self.active_sheet].name
    }

    // ── Multi-sheet (tabs) ───────────────────────────────────────────

    /// Common reset after a sheet swap: load the new sheet's cells, drop
    /// per-sheet state (selection, clipboard mark, undo history, etc.).
    fn after_sheet_change(&mut self) {
        self.cells = self
            .store
            .load_sheet(self.active_sheet_name())
            .map(|s| s.cells)
            .unwrap_or_default();
        self.columns = self
            .store
            .load_columns(self.active_sheet_name())
            .unwrap_or_default();
        self.cursor_row = 0;
        self.cursor_col = 0;
        self.scroll_row = 0;
        self.scroll_col = 0;
        self.selection_anchor = None;
        // [sheet.clipboard.sheet-switch-clears-mark] mark is per-sheet.
        self.clipboard_mark = None;
        // Undo/redo history is per-session, but reusing it across sheets
        // would let an undo apply changes to the wrong sheet — drop it.
        let _ = self.store.clear_undo_log();
        self.redo_stack.clear();
    }

    /// Switch to the sheet at `idx` if in range. No-op otherwise.
    pub fn switch_sheet(&mut self, idx: usize) {
        if idx >= self.sheets.len() || idx == self.active_sheet {
            return;
        }
        self.active_sheet = idx;
        self.after_sheet_change();
        self.status = format!("Switched to '{}'", self.active_sheet_name());
    }

    // [sheet.tabs.keyboard-switch]
    pub fn next_sheet(&mut self) {
        let n = self.sheets.len();
        if n > 1 {
            self.switch_sheet((self.active_sheet + 1) % n);
        }
    }

    pub fn prev_sheet(&mut self) {
        let n = self.sheets.len();
        if n > 1 {
            self.switch_sheet((self.active_sheet + n - 1) % n);
        }
    }

    // [sheet.tabs.add]
    /// Create a new sheet (with default columns A–H) and switch to it.
    pub fn add_sheet(&mut self, name: &str) -> Result<(), String> {
        if self.sheets.iter().any(|s| s.name == name) {
            return Err(format!("sheet '{name}' already exists"));
        }
        self.store.create_sheet(name).map_err(|e| e.to_string())?;
        self.sheets.push(SheetMeta {
            name: name.to_string(),
            sort_order: 0,
            color: None,
        });
        self.active_sheet = self.sheets.len() - 1;
        self.after_sheet_change();
        self.mark_touched();
        Ok(())
    }

    // ── Row / column structural ops ──────────────────────────────────

    // [sheet.column.insert-left-right]
    /// Insert `count` blank rows at the cursor position (`above`) or
    /// just below it (`below`). Caller controls direction.
    pub fn insert_rows_at_cursor(&mut self, side: InsertSide, count: u32) -> Result<(), String> {
        let at = match side {
            InsertSide::AboveOrLeft => self.cursor_row,
            InsertSide::BelowOrRight => self.cursor_row + 1,
        };
        let name = self.active_sheet_name().to_string();
        self.store
            .insert_rows(&name, at, count)
            .map_err(|e| e.to_string())?;
        self.refresh_cells();
        // Inserts invalidate undo/redo (the engine rewrites every formula
        // and shifts every cell — we'd need a much richer snapshot to
        // undo all that). Drop the history rather than letting a partial
        // undo desync the sheet from history.
        let _ = self.store.clear_undo_log();
        self.redo_stack.clear();
        self.mark_touched();
        Ok(())
    }

    // [sheet.column.insert-left-right]
    pub fn insert_cols_at_cursor(&mut self, side: InsertSide, count: u32) -> Result<(), String> {
        let at = match side {
            InsertSide::AboveOrLeft => self.cursor_col,
            InsertSide::BelowOrRight => self.cursor_col + 1,
        };
        let name = self.active_sheet_name().to_string();
        self.store
            .insert_cols(&name, at, count)
            .map_err(|e| e.to_string())?;
        self.refresh_cells();
        let _ = self.store.clear_undo_log();
        self.redo_stack.clear();
        self.mark_touched();
        Ok(())
    }

    /// Queue a row delete pending confirmation. Run-loop watches
    /// `pending_confirm` and intercepts the next y/n keystroke.
    pub fn request_delete_row(&mut self) {
        self.pending_confirm = Some(PendingConfirm::DeleteRow(self.cursor_row));
        self.status = format!("Delete row {}? (y/n)", self.cursor_row + 1);
    }

    pub fn request_delete_col(&mut self) {
        self.pending_confirm = Some(PendingConfirm::DeleteCol(self.cursor_col));
        let letter: String = to_cell_id(0, self.cursor_col)
            .chars()
            .take_while(|c| c.is_ascii_alphabetic())
            .collect();
        self.status = format!("Delete column {letter}? (y/n)");
    }

    /// Drop a queued confirm without executing.
    pub fn cancel_pending_confirm(&mut self) {
        if self.pending_confirm.take().is_some() {
            self.status = "Cancelled".into();
        }
    }

    // [sheet.delete.row-confirm]
    // [sheet.delete.column-confirm]
    /// Apply a previously-queued destructive op (delete row/col). Caller
    /// invokes this when the user presses `y`. No-op when nothing is
    /// pending.
    pub fn confirm_pending(&mut self) {
        let action = match self.pending_confirm.take() {
            Some(a) => a,
            None => return,
        };
        let name = self.active_sheet_name().to_string();
        let result: Result<(), String> = match action {
            PendingConfirm::DeleteRow(r) => self
                .store
                .delete_rows(&name, r, 1)
                .map_err(|e| e.to_string()),
            PendingConfirm::DeleteCol(c) => self
                .store
                .delete_cols(&name, c, 1)
                .map_err(|e| e.to_string()),
        };
        match result {
            Ok(()) => {
                self.refresh_cells();
                let _ = self.store.clear_undo_log();
                self.redo_stack.clear();
                self.status = match action {
                    PendingConfirm::DeleteRow(r) => format!("Deleted row {}", r + 1),
                    PendingConfirm::DeleteCol(_) => "Deleted column".into(),
                };
                // Clamp cursor in case it now points off-grid.
                self.cursor_row = self.cursor_row.min(MAX_ROW);
                self.cursor_col = self.cursor_col.min(MAX_COL);
                self.mark_touched();
            }
            Err(e) => self.status = format!("Error: {e}"),
        }
    }

    // [sheet.tabs.rename]
    /// Rename the active sheet to `new_name`. Rejects empty names and
    /// collisions with another existing sheet; renaming to the current
    /// name is a silent no-op. Drops the undo log because stored undo
    /// entries reference cells by `sheet_name` and would dangle after
    /// the rename.
    pub fn rename_active_sheet(&mut self, new_name: &str) -> Result<(), String> {
        let new_name = new_name.trim();
        if new_name.is_empty() {
            return Err("Sheet name cannot be empty".into());
        }
        let old = self.active_sheet_name().to_string();
        if new_name == old {
            return Ok(());
        }
        if self.sheets.iter().any(|s| s.name == new_name) {
            return Err(format!("sheet '{new_name}' already exists"));
        }
        self.store
            .rename_sheet(&old, new_name)
            .map_err(|e| e.to_string())?;
        self.sheets[self.active_sheet].name = new_name.to_string();
        let _ = self.store.clear_undo_log();
        self.redo_stack.clear();
        self.mark_touched();
        Ok(())
    }

    // [sheet.tabs.delete]
    /// Delete the active sheet. Refuses if it's the last one (per the
    /// spec: a workbook always has at least one sheet).
    pub fn delete_active_sheet(&mut self) -> Result<(), String> {
        if self.sheets.len() <= 1 {
            return Err("Cannot delete the last sheet".into());
        }
        let name = self.active_sheet_name().to_string();
        self.store.delete_sheet(&name).map_err(|e| e.to_string())?;
        self.sheets.remove(self.active_sheet);
        if self.active_sheet >= self.sheets.len() {
            self.active_sheet = self.sheets.len() - 1;
        }
        self.after_sheet_change();
        self.mark_touched();
        Ok(())
    }

    /// The current selection rectangle, falling back to a 1×1 rect at the
    /// cursor when nothing is range-selected.
    pub fn effective_rect(&self) -> (u32, u32, u32, u32) {
        self.selection_range().unwrap_or((
            self.cursor_row,
            self.cursor_col,
            self.cursor_row,
            self.cursor_col,
        ))
    }

    /// Snapshot the current raw + format for each cell in `changes`,
    /// returning [`UndoOp::CellChange`] entries that would restore the
    /// pre-edit state. Mirrors the prior `capture_undo_entry` behaviour
    /// including the `"null"` clear-sentinel for "prior format was
    /// absent" so undo actually drops a format added by the change.
    fn snapshot_undo_ops(&self, changes: &[CellChange]) -> Vec<UndoOp> {
        let sheet_name = self.active_sheet_name().to_string();
        changes
            .iter()
            .map(|ch| {
                let prior_format = self.get_format_json_raw(ch.row_idx, ch.col_idx);
                let format_json = match (prior_format, ch.format_json.as_deref()) {
                    (Some(prior), _) => Some(prior),
                    (None, Some(_)) => Some("null".to_string()),
                    (None, None) => None,
                };
                let prior_raw = self.get_raw(ch.row_idx, ch.col_idx);
                UndoOp::CellChange {
                    sheet_name: sheet_name.clone(),
                    row: ch.row_idx,
                    col: ch.col_idx,
                    // An empty raw means "no cell row" (the in-memory
                    // model collapses missing/empty for App's getters)
                    // — encode as None so undo deletes rather than
                    // re-inserts with empty raw.
                    raw: if prior_raw.is_empty() { None } else { Some(prior_raw) },
                    format_json,
                }
            })
            .collect()
    }

    /// Apply `changes` and (on success) record the inverse ops in
    /// `undo_entry` and clear the redo stack. All user-driven
    /// mutations (edit commit, delete, paste, range clear) flow
    /// through here.
    fn apply_changes_recorded(&mut self, changes: &[CellChange]) -> Result<(), String> {
        if changes.is_empty() {
            return Ok(());
        }
        let inverses = self.snapshot_undo_ops(changes);
        let name = self.active_sheet_name().to_string();
        self.store
            .apply(&name, changes)
            .map_err(|e| e.to_string())?;
        self.store
            .record_undo_group(&inverses)
            .map_err(|e| e.to_string())?;
        self.refresh_cells();
        self.mark_touched();
        // [sheet.undo.redo] a new action erases the redo branch.
        self.redo_stack.clear();
        Ok(())
    }

    // [sheet.undo.cmd-z]
    /// Pop the latest undo group from the on-disk log, replay its
    /// inverses, and push the freshly-captured forward state onto the
    /// in-memory redo stack. Returns true if anything was undone.
    pub fn undo(&mut self) -> bool {
        let popped = match self.store.pop_undo() {
            Ok(Some(group)) => group,
            Ok(None) => return false,
            Err(_) => return false,
        };
        self.refresh_cells();
        self.mark_touched();
        self.redo_stack.push(popped);
        true
    }

    // [sheet.undo.redo]
    /// Pop the in-memory redo stack, replay it, and let the Store
    /// re-record the inverses onto the undo log. Returns true if
    /// anything was redone.
    pub fn redo(&mut self) -> bool {
        let group = match self.redo_stack.pop() {
            Some(g) => g,
            None => return false,
        };
        if self.store.apply_redo(&group).is_err() {
            self.redo_stack.push(group);
            return false;
        }
        self.refresh_cells();
        self.mark_touched();
        true
    }

    // [sheet.clipboard.mark-visual]
    /// Replace any existing clipboard mark with one over the current
    /// selection (or single cell). Mode controls paste-time behaviour.
    pub fn set_clipboard_mark(&mut self, mode: ClipMarkMode) {
        let (r1, c1, r2, c2) = self.effective_rect();
        self.clipboard_mark = Some(ClipboardMark {
            r1,
            c1,
            r2,
            c2,
            mode,
        });
    }

    // [sheet.clipboard.escape-cancels-mark]
    /// Drop the clipboard mark without touching the OS clipboard. Returns
    /// true if a mark was active (so the caller knows Esc consumed something).
    pub fn clear_clipboard_mark(&mut self) -> bool {
        self.clipboard_mark.take().is_some()
    }

    // [sheet.navigation.arrow]
    // [sheet.navigation.tab-nav-move]
    pub fn move_cursor(&mut self, dr: i32, dc: i32) {
        let new_row = (self.cursor_row as i32 + dr).max(0) as u32;
        let new_col = (self.cursor_col as i32 + dc).max(0) as u32;
        self.cursor_row = new_row.min(MAX_ROW);
        self.cursor_col = new_col.min(MAX_COL);

        // Visual mode keeps its anchor — motions extend instead of reset.
        if !matches!(self.mode, Mode::Visual(_)) {
            self.selection_anchor = None;
        }

        self.scroll_to_cursor();
        self.status.clear();
    }

    // [sheet.navigation.shift-arrow-extend]
    /// Move cursor while extending selection. Sets anchor on first shift-move.
    pub fn move_cursor_selecting(&mut self, dr: i32, dc: i32) {
        if self.selection_anchor.is_none() {
            self.selection_anchor = Some((self.cursor_row, self.cursor_col));
        }

        let new_row = (self.cursor_row as i32 + dr).max(0) as u32;
        let new_col = (self.cursor_col as i32 + dc).max(0) as u32;
        self.cursor_row = new_row.min(MAX_ROW);
        self.cursor_col = new_col.min(MAX_COL);

        self.scroll_to_cursor();
    }

    /// Whether the cell's raw value is non-empty after trimming. Matches the
    /// `filled` predicate from the navigation specs.
    pub fn is_filled(&self, row: u32, col: u32) -> bool {
        !self.get_raw(row, col).trim().is_empty()
    }

    // [sheet.navigation.arrow-jump]
    /// Ctrl+Arrow: jump to the content boundary (or grid edge if no content
    /// in the direction of travel). Solo-selects the target.
    pub fn move_cursor_jump(&mut self, dr: i32, dc: i32) {
        let (r, c) = compute_jump_target(
            self.cursor_row,
            self.cursor_col,
            dr,
            dc,
            MAX_ROW,
            MAX_COL,
            |row, col| self.is_filled(row, col),
        );
        self.cursor_row = r;
        self.cursor_col = c;
        if !matches!(self.mode, Mode::Visual(_)) {
            self.selection_anchor = None;
        }
        self.scroll_to_cursor();
        self.status.clear();
    }

    // [sheet.navigation.shift-arrow-jump-extend]
    /// Shift+Ctrl+Arrow: jump to the content boundary AND extend the
    /// selection rectangle. Anchor stays put.
    pub fn move_cursor_jump_selecting(&mut self, dr: i32, dc: i32) {
        if self.selection_anchor.is_none() {
            self.selection_anchor = Some((self.cursor_row, self.cursor_col));
        }
        let (r, c) = compute_jump_target(
            self.cursor_row,
            self.cursor_col,
            dr,
            dc,
            MAX_ROW,
            MAX_COL,
            |row, col| self.is_filled(row, col),
        );
        self.cursor_row = r;
        self.cursor_col = c;
        self.scroll_to_cursor();
    }

    /// Ctrl+A: expand the selection to the contiguous rectangular block
    /// of non-empty cells (8-connected) containing the cursor. On an empty
    /// cell, selects the entire sheet (`A1:Z1000`). Always lands in
    /// `Visual::Cell` regardless of the prior mode.
    pub fn select_data_region(&mut self) {
        let populated: std::collections::HashSet<(u32, u32)> = self
            .cells
            .iter()
            .filter(|c| !c.raw_value.trim().is_empty())
            .map(|c| (c.row_idx, c.col_idx))
            .collect();

        let (r1, c1, r2, c2) =
            match compute_data_region(self.cursor_row, self.cursor_col, &populated) {
                Some(rect) => rect,
                None => (0, 0, MAX_ROW, MAX_COL),
            };

        self.mode = Mode::Visual(VisualKind::Cell);
        self.selection_anchor = Some((r1, c1));
        self.cursor_row = r2;
        self.cursor_col = c2;
        self.scroll_to_cursor();
        self.status = format!(
            "Selected {}:{}",
            to_cell_id(r1, c1),
            to_cell_id(r2, c2),
        );
    }

    // ── Vim motions (V2) ─────────────────────────────────────────────

    /// Common cursor jump used by every absolute motion below. Clamps,
    /// drops the selection anchor (unless in Visual mode, where motions
    /// extend the existing selection), and scrolls into view.
    pub fn jump_cursor_to(&mut self, row: u32, col: u32) {
        self.cursor_row = row.min(MAX_ROW);
        self.cursor_col = col.min(MAX_COL);
        if !matches!(self.mode, Mode::Visual(_)) {
            self.selection_anchor = None;
        }
        self.scroll_to_cursor();
        self.status.clear();
    }

    /// Vim `0`: column 0 of current row.
    pub fn move_to_row_start(&mut self) {
        self.jump_cursor_to(self.cursor_row, 0);
    }

    /// Vim `^`: first filled cell of the current row, falling back to col 0.
    pub fn move_to_first_filled_in_row(&mut self) {
        let row = self.cursor_row;
        let target = (0..=MAX_COL).find(|&c| self.is_filled(row, c)).unwrap_or(0);
        self.jump_cursor_to(row, target);
    }

    /// Vim `$`: last filled cell of the current row, falling back to MAX_COL.
    pub fn move_to_last_filled_in_row(&mut self) {
        let row = self.cursor_row;
        let target = (0..=MAX_COL)
            .rev()
            .find(|&c| self.is_filled(row, c))
            .unwrap_or(MAX_COL);
        self.jump_cursor_to(row, target);
    }

    /// Vim `gg` (no count): row 0 of current column.
    pub fn goto_first_row(&mut self) {
        self.jump_cursor_to(0, self.cursor_col);
    }

    /// Vim `G` (no count): last row anywhere in the sheet that has data,
    /// falling back to row 0 on a completely empty sheet. Global rather
    /// than column-aware so an empty column doesn't fling the cursor to
    /// row 1000 just because nothing is filled below the current cell.
    /// With a count, the dispatcher calls `goto_row` instead.
    pub fn goto_last_filled_row(&mut self) {
        let target = self.cells.iter().map(|c| c.row_idx).max().unwrap_or(0);
        self.jump_cursor_to(target, self.cursor_col);
    }

    /// Vim `{count}G` / `{count}gg`: jump to absolute row `n` (0-indexed).
    /// Caller is responsible for converting from the 1-indexed user input.
    pub fn goto_row(&mut self, row: u32) {
        self.jump_cursor_to(row, self.cursor_col);
    }

    /// Vim `{`: previous non-empty row in current column. Walks back to the
    /// start of the run if currently inside one.
    pub fn paragraph_backward(&mut self) {
        let col = self.cursor_col;
        let mut r = self.cursor_row;
        // Step back through any empties (or out of the current run, if at
        // its start).
        loop {
            if r == 0 {
                self.jump_cursor_to(0, col);
                return;
            }
            r -= 1;
            if self.is_filled(r, col) {
                break;
            }
        }
        // Walk back to the first cell of the contiguous run.
        while r > 0 && self.is_filled(r - 1, col) {
            r -= 1;
        }
        self.jump_cursor_to(r, col);
    }

    /// Vim `}`: next non-empty row in current column. Walks forward to the
    /// end of the run if currently inside one.
    pub fn paragraph_forward(&mut self) {
        let col = self.cursor_col;
        let mut r = self.cursor_row;
        loop {
            if r >= MAX_ROW {
                self.jump_cursor_to(MAX_ROW, col);
                return;
            }
            r += 1;
            if self.is_filled(r, col) {
                break;
            }
        }
        while r < MAX_ROW && self.is_filled(r + 1, col) {
            r += 1;
        }
        self.jump_cursor_to(r, col);
    }

    /// Vim `H`: top of viewport.
    pub fn move_to_viewport_top(&mut self) {
        self.jump_cursor_to(self.scroll_row, self.cursor_col);
    }

    /// Vim `M`: middle of viewport.
    pub fn move_to_viewport_middle(&mut self) {
        let r = self.scroll_row + self.visible_rows / 2;
        self.jump_cursor_to(r, self.cursor_col);
    }

    /// Vim `L`: bottom of viewport.
    pub fn move_to_viewport_bottom(&mut self) {
        let r = self.scroll_row + self.visible_rows.saturating_sub(1);
        self.jump_cursor_to(r, self.cursor_col);
    }

    /// Move the viewport by `(dr, dc)` rows/columns *without touching
    /// the cursor*. Used by the mouse wheel (T3) — Excel / VisiData
    /// convention: scrolling reveals new content; the cursor doesn't
    /// follow until the user explicitly moves it.
    pub fn scroll_viewport(&mut self, dr: i32, dc: i32) {
        let new_row = (self.scroll_row as i64 + dr as i64).clamp(0, MAX_ROW as i64);
        let new_col = (self.scroll_col as i64 + dc as i64).clamp(0, MAX_COL as i64);
        self.scroll_row = new_row as u32;
        self.scroll_col = new_col as u32;
    }

    /// Vim `Ctrl+d` / `Ctrl+u`: half-page scroll. `dir` is +1 for down, -1 for up.
    pub fn scroll_half_page(&mut self, dir: i32) {
        let half = (self.visible_rows / 2).max(1) as i32;
        let new = (self.cursor_row as i32 + dir * half).max(0) as u32;
        self.jump_cursor_to(new, self.cursor_col);
    }

    /// Vim `Ctrl+f` / `Ctrl+b`: full-page scroll.
    pub fn scroll_full_page(&mut self, dir: i32) {
        let page = self.visible_rows.max(1) as i32;
        let new = (self.cursor_row as i32 + dir * page).max(0) as u32;
        self.jump_cursor_to(new, self.cursor_col);
    }

    /// Vim `zz`: re-scroll so the cursor's row is at the viewport's middle.
    pub fn scroll_cursor_to_middle(&mut self) {
        let half = self.visible_rows / 2;
        self.scroll_row = self.cursor_row.saturating_sub(half);
    }
    /// Vim `zt`: cursor's row to top of viewport.
    pub fn scroll_cursor_to_top(&mut self) {
        self.scroll_row = self.cursor_row;
    }
    /// Vim `zb`: cursor's row to bottom of viewport.
    pub fn scroll_cursor_to_bottom(&mut self) {
        self.scroll_row = self
            .cursor_row
            .saturating_sub(self.visible_rows.saturating_sub(1));
    }
    /// Vim `zh`: scroll viewport one column left, leave cursor in place.
    pub fn scroll_viewport_left(&mut self) {
        self.scroll_col = self.scroll_col.saturating_sub(1);
    }
    /// Vim `zl`: scroll viewport one column right.
    pub fn scroll_viewport_right(&mut self) {
        self.scroll_col = (self.scroll_col + 1).min(MAX_COL);
    }

    /// Vim `w`: forward to start of next "word" (filled run).
    pub fn word_forward(&mut self) {
        let (r, c) = compute_word_forward(
            self.cursor_row,
            self.cursor_col,
            MAX_COL,
            |row, col| self.is_filled(row, col),
        );
        self.jump_cursor_to(r, c);
    }

    /// Vim `b`: backward to start of current/previous word.
    pub fn word_backward(&mut self) {
        let (r, c) = compute_word_backward(
            self.cursor_row,
            self.cursor_col,
            |row, col| self.is_filled(row, col),
        );
        self.jump_cursor_to(r, c);
    }

    /// Vim `e`: forward to end of current/next word. Equivalent to the
    /// existing Ctrl+Right `compute_jump_target` rules: from a filled cell
    /// it walks to the end of the run, from empty it skips to the next
    /// filled cell.
    pub fn word_end(&mut self) {
        let (r, c) = compute_jump_target(
            self.cursor_row,
            self.cursor_col,
            0,
            1,
            MAX_ROW,
            MAX_COL,
            |row, col| self.is_filled(row, col),
        );
        self.jump_cursor_to(r, c);
    }

    fn scroll_to_cursor(&mut self) {
        if self.cursor_row < self.scroll_row {
            self.scroll_row = self.cursor_row;
        } else if self.cursor_row >= self.scroll_row + self.visible_rows {
            self.scroll_row = self.cursor_row - self.visible_rows + 1;
        }
        if self.cursor_col < self.scroll_col {
            self.scroll_col = self.cursor_col;
        } else if self.cursor_col >= self.scroll_col + self.visible_cols {
            self.scroll_col = self.cursor_col - self.visible_cols + 1;
        }
    }

    /// Returns the selected range as (min_row, min_col, max_row, max_col), or None.
    /// In V-LINE mode the column extents are forced to span the full grid;
    /// in V-COLUMN mode the row extents are forced.
    pub fn selection_range(&self) -> Option<(u32, u32, u32, u32)> {
        let (ar, ac) = self.selection_anchor?;
        let r1 = ar.min(self.cursor_row);
        let r2 = ar.max(self.cursor_row);
        if matches!(self.mode, Mode::Visual(VisualKind::Row)) {
            return Some((r1, 0, r2, MAX_COL));
        }
        let c1 = ac.min(self.cursor_col);
        let c2 = ac.max(self.cursor_col);
        if matches!(self.mode, Mode::Visual(VisualKind::Column)) {
            return Some((0, c1, MAX_ROW, c2));
        }
        Some((r1, c1, r2, c2))
    }

    /// Returns true if the given cell is inside the current selection.
    pub fn is_selected(&self, row: u32, col: u32) -> bool {
        if let Some((r1, c1, r2, c2)) = self.selection_range() {
            row >= r1 && row <= r2 && col >= c1 && col <= c2
        } else {
            false
        }
    }

    /// Build TSV text for the selected range (or single cell if no selection).
    pub fn yank_tsv(&self) -> String {
        let (r1, c1, r2, c2) = self
            .selection_range()
            .unwrap_or((self.cursor_row, self.cursor_col, self.cursor_row, self.cursor_col));

        let mut lines = Vec::new();
        for row in r1..=r2 {
            let mut cols = Vec::new();
            for col in c1..=c2 {
                cols.push(self.get_display(row, col));
            }
            lines.push(cols.join("\t"));
        }
        lines.join("\n")
    }

    /// Build an HTML table for the selected range (or single cell), with
    /// the app-private formula round-trip channel (see
    /// [sheet.clipboard.copy] / [sheet.clipboard.paste-formula-shift]):
    ///
    /// - The `<table>` carries `data-sheets-source-anchor="<top-left>"`
    ///   so a paste back into vlotus can compute the source→target
    ///   delta for relative-ref shifting.
    /// - Each `<td>` whose source cell holds a formula carries
    ///   `data-sheets-formula="=..."` with the raw formula text.
    /// - External apps ignore unknown attributes — the visible value is
    ///   unchanged, so paste into Excel / Sheets / docs / email behaves
    ///   identically to a plain HTML table.
    pub fn yank_html(&self) -> String {
        let (r1, c1, r2, c2) = self.effective_rect();

        let anchor_id = to_cell_id(r1, c1);
        let mut html = format!(
            "<table style=\"border-collapse:collapse;border:none\" \
             data-sheets-source-anchor=\"{anchor_id}\">\n<tbody>\n"
        );
        for row in r1..=r2 {
            html.push_str("<tr>");
            for col in c1..=c2 {
                let val = self.get_display(row, col);
                let raw = self.get_raw(row, col);
                let escaped_val = html_escape(&val);
                if raw.starts_with('=') {
                    let escaped_formula = html_escape(&raw);
                    html.push_str(&format!(
                        "<td style=\"border:1px solid #cccccc;padding:2px 3px\" \
                         data-sheets-formula=\"{escaped_formula}\">{escaped_val}</td>"
                    ));
                } else {
                    html.push_str(&format!(
                        "<td style=\"border:1px solid #cccccc;padding:2px 3px\">{escaped_val}</td>"
                    ));
                }
            }
            html.push_str("</tr>\n");
        }
        html.push_str("</tbody>\n</table>");
        html
    }

    /// Compute stats (count, sum, min, max, mean) for numeric cells in the selection.
    pub fn selection_stats(&self) -> Option<String> {
        let (r1, c1, r2, c2) = self.selection_range()?;
        let mut nums: Vec<f64> = Vec::new();
        let mut total_cells: u32 = 0;

        for row in r1..=r2 {
            for col in c1..=c2 {
                total_cells += 1;
                let display = self.get_display(row, col);
                if let Ok(v) = display.parse::<f64>() {
                    nums.push(v);
                }
            }
        }

        if nums.is_empty() {
            return Some(format!("Count: {total_cells}"));
        }

        let count = nums.len();
        let min = nums.iter().cloned().fold(f64::INFINITY, f64::min);
        let max = nums.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
        let sum: f64 = nums.iter().sum();
        let mean = sum / count as f64;

        Some(format!(
            "Count: {total_cells}  Sum: {sum}  Min: {min}  Max: {max}  Mean: {mean:.2}"
        ))
    }

    // [sheet.editing.f2-or-enter]
    pub fn start_edit(&mut self) {
        self.mode = Mode::Edit;
        // Pre-fill with current raw value
        self.edit_buf = self.get_raw(self.cursor_row, self.cursor_col);
        self.edit_cursor = self.edit_buf.len();
        self.status = "EDIT — Enter to confirm, Esc to cancel".into();
        self.refresh_completions();
    }

    /// Vim `i`/`I`: enter Insert with the current raw value pre-filled and
    /// the caret at the start. `a`/`A` map onto `start_edit` (caret at end).
    pub fn start_edit_at_start(&mut self) {
        self.mode = Mode::Edit;
        self.edit_buf = self.get_raw(self.cursor_row, self.cursor_col);
        self.edit_cursor = 0;
        self.status = "EDIT — Enter to confirm, Esc to cancel".into();
        self.refresh_completions();
    }

    /// Vim `o`/`O`/`s`/`S`: enter Insert with an empty buffer. The current
    /// cell value is replaced wholesale on commit (Esc still cancels).
    pub fn start_edit_blank(&mut self) {
        self.mode = Mode::Edit;
        self.edit_buf.clear();
        self.edit_cursor = 0;
        self.status = "EDIT — Enter to confirm, Esc to cancel".into();
        self.refresh_completions();
    }

    // [sheet.navigation.enter-commit-down]
    /// Commit the edit and move down (Enter behaviour).
    pub fn confirm_edit(&mut self) {
        self.confirm_edit_advancing(1, 0);
    }

    /// Build a `CellChange` from a user-typed buffer. Recognises `$1.25`
    /// style input as USD with the `$` peeled off the stored raw value;
    /// formula buffers (`=…`) and plain text are passed through unchanged.
    /// Used by edit-commit and `.` dot-repeat so both honour auto-detect.
    fn cell_change_from_typed(row: u32, col: u32, buf: &str) -> CellChange {
        if !buf.starts_with('=') {
            if let Some((numeric, fmt)) = format::try_parse_typed_input(buf) {
                return CellChange {
                    row_idx: row,
                    col_idx: col,
                    raw_value: numeric,
                    format_json: Some(format::to_format_json(&fmt)),
                };
            }
        }
        CellChange {
            row_idx: row,
            col_idx: col,
            raw_value: buf.to_string(),
            format_json: None,
        }
    }

    // [sheet.navigation.tab-commit-right]
    /// Commit the edit and move one column right (Tab) or left (Shift+Tab).
    /// At the relevant grid edge, commit but stay put — `move_cursor`'s
    /// clamp handles the no-op.
    pub fn confirm_edit_advancing(&mut self, dr: i32, dc: i32) {
        self.mode = Mode::Nav;
        let raw = self.edit_buf.clone();
        let change = Self::cell_change_from_typed(self.cursor_row, self.cursor_col, &raw);
        match self.apply_changes_recorded(&[change]) {
            Ok(()) => self.status = "Saved".into(),
            Err(e) => self.status = format!("Error: {e}"),
        }
        // V6: dot-repeat. Either complete an in-flight Change record (set
        // by apply_operator before it dropped us into Insert) or start a
        // fresh Insert record.
        match self.last_edit.as_mut() {
            Some(action) if action.kind == EditKind::Change && action.text.is_none() => {
                action.text = Some(raw);
            }
            _ => {
                self.last_edit = Some(EditAction {
                    kind: EditKind::Insert,
                    anchor_dr: 0,
                    anchor_dc: 0,
                    rect_rows: 1,
                    rect_cols: 1,
                    text: Some(raw),
                });
            }
        }
        self.edit_buf.clear();
        self.edit_cursor = 0;
        self.autocomplete = None;
        self.signature = None;
        self.pointing = None;
        self.move_cursor(dr, dc);
    }

    pub fn cancel_edit(&mut self) {
        self.mode = Mode::Nav;
        self.edit_buf.clear();
        self.edit_cursor = 0;
        self.status.clear();
        self.autocomplete = None;
        self.signature = None;
        self.pointing = None;
        // V6: a half-built Change record (text=None) means apply_operator
        // already cleared the rect but the user Esc'd before typing.
        // Replay that as a plain Delete instead of a no-text Change.
        if let Some(action) = self.last_edit.as_mut() {
            if action.kind == EditKind::Change && action.text.is_none() {
                action.kind = EditKind::Delete;
            }
        }
    }

    // ── Formula-editor completions / signature help ────────────────

    // [sheet.editing.formula-autocomplete]
    // [sheet.editing.formula-signature-help]
    /// Recompute autocomplete + signature_help from the current edit buffer.
    /// Cheap (re-tokenizes only `edit_buf[..edit_cursor]`); call after every
    /// edit mutation. Outside Edit mode this clears both.
    pub fn refresh_completions(&mut self) {
        if self.mode != Mode::Edit {
            self.autocomplete = None;
            self.signature = None;
            return;
        }
        // Named ranges are not yet supported in the TUI — pass an empty list.
        let names: [&str; 0] = [];
        let list = complete_with_registry(&self.registry, &self.edit_buf, self.edit_cursor, &names);
        // Only open the popup once a function-name prefix is actually being
        // typed. Without this gate, an empty prefix (cursor right after `=`,
        // `(`, `,`, or an operator) would dump every function into the popup
        // and Up/Down would steer that list instead of falling through to
        // pointing mode for the cell above/below.
        let has_prefix = list.replace_start < list.replace_end;
        self.autocomplete = if list.items.is_empty() || !has_prefix {
            None
        } else {
            // Only carry the prior selection forward if the user actually
            // navigated; otherwise typing more should land back on the
            // top-ranked item. `user_selected` similarly persists only
            // while the popup remains open and the user has navigated.
            let (prev_selected, prev_user) = self
                .autocomplete
                .as_ref()
                .map(|a| (a.selected, a.user_selected))
                .unwrap_or((0, false));
            let selected = if prev_user {
                prev_selected.min(list.items.len() - 1)
            } else {
                0
            };
            Some(AutocompleteState {
                list,
                selected,
                user_selected: prev_user,
            })
        };
        self.signature = signature_help(&self.edit_buf, self.edit_cursor);
    }

    pub fn autocomplete_select_next(&mut self) {
        if let Some(a) = &mut self.autocomplete {
            if !a.list.items.is_empty() {
                a.selected = (a.selected + 1) % a.list.items.len();
                a.user_selected = true;
            }
        }
    }

    pub fn autocomplete_select_prev(&mut self) {
        if let Some(a) = &mut self.autocomplete {
            let n = a.list.items.len();
            if n > 0 {
                a.selected = (a.selected + n - 1) % n;
                a.user_selected = true;
            }
        }
    }

    /// Accept the highlighted completion: replace the prefix being typed
    /// with the item's `insert` text. For functions, append `(` if the
    /// insert doesn't already include one. Returns true if anything was
    /// inserted.
    pub fn autocomplete_accept(&mut self) -> bool {
        let Some(state) = self.autocomplete.take() else {
            return false;
        };
        let item = match state.list.items.get(state.selected) {
            Some(it) => it.clone(),
            None => return false,
        };
        let mut insert_text = item.insert;
        if matches!(item.kind, CompletionKind::Function) && !insert_text.contains('(') {
            insert_text.push('(');
        }
        let start = state.list.replace_start.min(self.edit_buf.len());
        let end = state.list.replace_end.min(self.edit_buf.len());
        self.edit_buf.replace_range(start..end, &insert_text);
        self.edit_cursor = start + insert_text.len();
        // Pick up the new function call's signature help, if any.
        self.signature = signature_help(&self.edit_buf, self.edit_cursor);
        // After accepting a function we expect to type arguments; the
        // ranked completions for "" are huge and not useful, so leave
        // the autocomplete dismissed until the user types more.
        true
    }

    pub fn autocomplete_dismiss(&mut self) {
        self.autocomplete = None;
    }

    // ── Formula-ref pointing ─────────────────────────────────────────

    // [sheet.editing.formula-ref-pointing]
    /// Try to start (or continue) a pointing session. Returns true if the
    /// arrow key was consumed by pointing — caller should suppress the
    /// normal edit-caret movement.
    ///
    /// `extend` mirrors mouse-drag behaviour: while a Cell-kind pointing
    /// session is active, Shift+Arrow keeps the anchor pinned and only
    /// advances the target, producing range syntax (`B2:B3`). Plain Arrow
    /// (`extend == false`) collapses the range and resets the anchor to
    /// the new target. On entry to pointing, `extend` is irrelevant —
    /// there's no anchor yet, so the first ref is always single-cell with
    /// `anchor == target`, ready for follow-up Shift+Arrows.
    pub fn try_pointing_arrow(&mut self, dr: i32, dc: i32, extend: bool) -> bool {
        if self.mode != Mode::Edit {
            return false;
        }
        // Already pointing: extend the range, or move the ref.
        if let Some(p) = self.pointing {
            let new_row = ((p.target_row as i32 + dr).max(0) as u32).min(MAX_ROW);
            let new_col = ((p.target_col as i32 + dc).max(0) as u32).min(MAX_COL);
            if extend && p.kind == PointingKind::Cell {
                // Pin anchor, advance target only. rewrite_pointing_text
                // handles min/max normalisation and the collapse-when-
                // target==anchor case.
                let mut q = p;
                q.target_row = new_row;
                q.target_col = new_col;
                self.pointing = Some(q);
                self.rewrite_pointing_text();
            } else {
                // Plain arrow (or non-Cell kind, which is mouse-only) —
                // collapse to a fresh single-cell ref at the new target.
                self.replace_pointing_ref(new_row, new_col);
            }
            return true;
        }
        // Otherwise, only enter pointing when the caret is at an
        // insertable position (right after an operator/comma/(/=).
        if !is_insertable_at(&self.edit_buf, self.edit_cursor) {
            return false;
        }
        let start_row = ((self.cursor_row as i32 + dr).max(0) as u32).min(MAX_ROW);
        let start_col = ((self.cursor_col as i32 + dc).max(0) as u32).min(MAX_COL);
        let ref_text = to_cell_id(start_row, start_col);
        let start = self.edit_cursor;
        self.edit_buf.insert_str(start, &ref_text);
        let end = start + ref_text.len();
        self.edit_cursor = end;
        self.pointing = Some(PointingState {
            start,
            end,
            kind: PointingKind::Cell,
            anchor_row: start_row,
            anchor_col: start_col,
            target_row: start_row,
            target_col: start_col,
        });
        // Caret moved programmatically — refresh so signature_help follows.
        self.refresh_completions();
        true
    }

    /// Insert a cell ref at the current edit caret (or replace the
    /// existing pointing-mode ref). Returns `true` if it inserted /
    /// replaced — caller can early-return — or `false` if the caret
    /// isn't in a position where a ref would be syntactically valid.
    /// Mirrors `try_pointing_arrow` but with an explicit target instead
    /// of a delta from the cursor; T8 wires this into mouse clicks
    /// during formula editing.
    pub fn insert_ref_at_caret(&mut self, target_row: u32, target_col: u32) -> bool {
        if self.mode != Mode::Edit {
            return false;
        }
        if self.pointing.is_some() {
            self.replace_pointing_ref(target_row, target_col);
            return true;
        }
        if !is_insertable_at(&self.edit_buf, self.edit_cursor) {
            return false;
        }
        let ref_text = to_cell_id(target_row, target_col);
        let start = self.edit_cursor;
        self.edit_buf.insert_str(start, &ref_text);
        let end = start + ref_text.len();
        self.edit_cursor = end;
        self.pointing = Some(PointingState {
            start,
            end,
            kind: PointingKind::Cell,
            anchor_row: target_row,
            anchor_col: target_col,
            target_row,
            target_col,
        });
        self.refresh_completions();
        true
    }

    /// Extend the active pointing-mode ref's target end. T9/T12/T13
    /// drag handlers pass the new (row, col); the rewrite logic uses
    /// only the dimensions appropriate to the pointing kind (rows for
    /// Row, cols for Column, both for Cell).
    pub fn drag_pointing_target(&mut self, target_row: u32, target_col: u32) {
        let Some(mut p) = self.pointing else { return };
        p.target_row = target_row;
        p.target_col = target_col;
        self.pointing = Some(p);
        self.rewrite_pointing_text();
    }

    /// Insert a whole-column ref (`B:B`) at the edit caret and start a
    /// `PointingKind::Column` session so a follow-up column-header drag
    /// (T12) can extend it into `B:E`. Replaces any active pointing of
    /// any other kind (cell, row) with the column ref.
    pub fn insert_col_ref_at_caret(&mut self, col: u32) -> bool {
        self.insert_axis_ref_at_caret(PointingKind::Column, 0, col)
    }

    /// Insert a whole-row ref (`1:1`, 1-indexed for display) at the
    /// edit caret and start a `PointingKind::Row` session so a follow-
    /// up row-header drag (T13) can extend it into `1:5`.
    pub fn insert_row_ref_at_caret(&mut self, row: u32) -> bool {
        self.insert_axis_ref_at_caret(PointingKind::Row, row, 0)
    }

    /// Shared T10/T11 implementation. Replaces an active pointing span
    /// (regardless of kind) with the new axis ref, or — if no pointing
    /// is active — gates on `is_insertable_at` before inserting. Either
    /// way, on success a fresh pointing session of `kind` is registered
    /// at the inserted span so a drag can extend it.
    fn insert_axis_ref_at_caret(
        &mut self,
        kind: PointingKind,
        anchor_row: u32,
        anchor_col: u32,
    ) -> bool {
        if self.mode != Mode::Edit {
            return false;
        }
        let (start, _) = if let Some(p) = self.pointing {
            // Existing pointing span — replace it (regardless of kind).
            self.edit_buf.replace_range(p.start..p.end, "");
            self.edit_cursor = p.start;
            (p.start, p.end)
        } else {
            if !is_insertable_at(&self.edit_buf, self.edit_cursor) {
                return false;
            }
            (self.edit_cursor, self.edit_cursor)
        };
        // Compute the ref text via rewrite_pointing_text so single-vs-
        // range formatting stays in one place.
        self.pointing = Some(PointingState {
            start,
            end: start, // empty span; rewrite fills it in
            kind,
            anchor_row,
            anchor_col,
            target_row: anchor_row,
            target_col: anchor_col,
        });
        self.rewrite_pointing_text();
        true
    }

    /// Replace the pointing-mode reference span with one for the new
    /// target cell. Collapses any in-flight range back to a single-cell
    /// ref (anchor = target = new cell) and forces kind back to Cell.
    /// No-op outside pointing mode.
    fn replace_pointing_ref(&mut self, target_row: u32, target_col: u32) {
        let Some(mut p) = self.pointing else {
            return;
        };
        p.kind = PointingKind::Cell;
        p.anchor_row = target_row;
        p.anchor_col = target_col;
        p.target_row = target_row;
        p.target_col = target_col;
        self.pointing = Some(p);
        self.rewrite_pointing_text();
    }

    /// Re-render the pointing span in `edit_buf` from `anchor_*` /
    /// `target_*` and the pointing `kind`. Cell-kind uses single or
    /// `B2:E5` syntax; Column-kind uses `B:E` (or `B:B` collapsed);
    /// Row-kind uses `1:5` (or `1:1`). Updates `edit_cursor` and
    /// `p.end` to the new span end.
    fn rewrite_pointing_text(&mut self) {
        let Some(mut p) = self.pointing else { return };
        let new_text = match p.kind {
            PointingKind::Cell => {
                if (p.anchor_row, p.anchor_col) == (p.target_row, p.target_col) {
                    to_cell_id(p.target_row, p.target_col)
                } else {
                    let r1 = p.anchor_row.min(p.target_row);
                    let c1 = p.anchor_col.min(p.target_col);
                    let r2 = p.anchor_row.max(p.target_row);
                    let c2 = p.anchor_col.max(p.target_col);
                    format!("{}:{}", to_cell_id(r1, c1), to_cell_id(r2, c2))
                }
            }
            PointingKind::Column => {
                let c1 = p.anchor_col.min(p.target_col);
                let c2 = p.anchor_col.max(p.target_col);
                let l1 = col_idx_to_letters(c1);
                let l2 = col_idx_to_letters(c2);
                format!("{l1}:{l2}")
            }
            PointingKind::Row => {
                let r1 = p.anchor_row.min(p.target_row) + 1;
                let r2 = p.anchor_row.max(p.target_row) + 1;
                format!("{r1}:{r2}")
            }
        };
        self.edit_buf.replace_range(p.start..p.end, &new_text);
        let new_end = p.start + new_text.len();
        self.edit_cursor = new_end;
        p.end = new_end;
        self.pointing = Some(p);
        self.refresh_completions();
    }

    /// Stop the pointing session without altering the inserted ref.
    /// Called whenever any non-arrow editing key fires.
    pub fn exit_pointing(&mut self) {
        self.pointing = None;
    }

    // ── Command mode (`:` prompt) ───────────────────────────────────

    /// Open the `:` command prompt with an empty buffer.
    pub fn start_command(&mut self) {
        self.mode = Mode::Command;
        self.edit_buf.clear();
        self.edit_cursor = 0;
        self.status.clear();
    }

    /// Drop the `:` command prompt without executing.
    pub fn cancel_command(&mut self) {
        self.mode = Mode::Nav;
        self.edit_buf.clear();
        self.edit_cursor = 0;
    }

    // ── Shell mode (`!` prompt) ─────────────────────────────────────

    /// Open the `!` shell prompt with an empty buffer.
    pub fn start_shell(&mut self) {
        self.mode = Mode::Shell;
        self.edit_buf.clear();
        self.edit_cursor = 0;
        self.status.clear();
    }

    /// Drop the `!` shell prompt without executing.
    pub fn cancel_shell(&mut self) {
        self.mode = Mode::Nav;
        self.edit_buf.clear();
        self.edit_cursor = 0;
    }

    /// Run the buffered shell command. Captures stdout via
    /// [`shell::run`], sniffs it via [`shell::detect_payload`], and
    /// pastes the resulting grid at the cursor with the first row as
    /// headers. The whole paste is one undo group (inherited from
    /// `apply_pasted_grid`).
    pub fn run_shell(&mut self) -> CommandOutcome {
        let cmd = self.edit_buf.trim().to_string();
        self.mode = Mode::Nav;
        self.edit_buf.clear();
        self.edit_cursor = 0;
        if cmd.is_empty() {
            return CommandOutcome::Continue;
        }

        let stdout = match crate::shell::run(&cmd) {
            Ok(s) => s,
            Err(e) => {
                self.status = format!("!: {e}");
                return CommandOutcome::Continue;
            }
        };

        let Some(grid) = crate::shell::detect_payload(&stdout) else {
            self.status = "!: empty output".into();
            return CommandOutcome::Continue;
        };

        self.insert_shell_payload(&grid);
        CommandOutcome::Continue
    }

    /// Paste a sniffed shell payload at the cursor. Factored out from
    /// `run_shell` so tests can exercise the paste path without
    /// spawning a subprocess.
    pub fn insert_shell_payload(&mut self, grid: &PastedGrid) {
        match self.apply_pasted_grid(grid, None) {
            Ok((rows, cols)) if rows == 0 && cols == 0 => {
                self.status = "!: empty output".into();
            }
            Ok((rows, cols)) => {
                self.status = format!("!: pasted {rows}×{cols}");
            }
            Err(e) => self.status = format!("!: {e}"),
        }
    }

    /// Execute the buffered `:` command and return what the run-loop
    /// should do (`Continue` or `Quit`).
    pub fn run_command(&mut self) -> CommandOutcome {
        let cmd = self.edit_buf.trim().to_string();
        self.mode = Mode::Nav;
        self.edit_buf.clear();
        self.edit_cursor = 0;

        // Multi-token state-mutating commands run here so they can borrow
        // `&mut self`. Pure single-token commands stay in `execute_command`.
        let parts: Vec<&str> = cmd.split_whitespace().collect();
        match parts.as_slice() {
            // [sheet.tabs.add]
            ["sheet", "new", rest @ ..] => {
                let name = if rest.is_empty() {
                    format!("Sheet{}", self.sheets.len() + 1)
                } else {
                    rest.join(" ")
                };
                match self.add_sheet(&name) {
                    Ok(()) => self.status = format!("Added sheet '{name}'"),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            // [sheet.tabs.delete]
            ["sheet", "del" | "delete"] => {
                let name = self.active_sheet_name().to_string();
                match self.delete_active_sheet() {
                    Ok(()) => self.status = format!("Deleted sheet '{name}'"),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            // [sheet.tabs.rename]
            ["sheet", "rename" | "ren", rest @ ..] => {
                if rest.is_empty() {
                    self.status = "Use :sheet rename <newname>".into();
                    return CommandOutcome::Continue;
                }
                let new_name = rest.join(" ");
                let old = self.active_sheet_name().to_string();
                match self.rename_active_sheet(&new_name) {
                    Ok(()) => self.status = format!("Renamed '{old}' → '{new_name}'"),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            // [sheet.column.insert-left-right]
            ["row", "ins" | "insert", side] => {
                let r = match parse_insert_side(side) {
                    Some(InsertSide::AboveOrLeft) => InsertSide::AboveOrLeft,
                    Some(InsertSide::BelowOrRight) => InsertSide::BelowOrRight,
                    None => {
                        self.status = "Use :row ins above|below".into();
                        return CommandOutcome::Continue;
                    }
                };
                match self.insert_rows_at_cursor(r, 1) {
                    Ok(()) => self.status = "Inserted row".into(),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            ["col", "ins" | "insert", side] => {
                let s = match parse_insert_side(side) {
                    Some(s) => s,
                    None => {
                        self.status = "Use :col ins left|right".into();
                        return CommandOutcome::Continue;
                    }
                };
                match self.insert_cols_at_cursor(s, 1) {
                    Ok(()) => self.status = "Inserted column".into(),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            // [sheet.delete.row-confirm]
            ["row", "del" | "delete"] => {
                self.request_delete_row();
                return CommandOutcome::Continue;
            }
            // [sheet.delete.column-confirm]
            ["col", "del" | "delete"] => {
                self.request_delete_col();
                return CommandOutcome::Continue;
            }
            ["sheet", "list" | "ls"] | ["tabs"] => {
                let names: Vec<String> = self
                    .sheets
                    .iter()
                    .enumerate()
                    .map(|(i, s)| {
                        if i == self.active_sheet {
                            format!("[{}]", s.name)
                        } else {
                            s.name.clone()
                        }
                    })
                    .collect();
                self.status = names.join("  ");
                return CommandOutcome::Continue;
            }
            // V7 vim `:tab*` aliases for the existing sheet plumbing.
            ["tabnew", rest @ ..] | ["tabe" | "tabedit", rest @ ..] => {
                let name = if rest.is_empty() {
                    format!("Sheet{}", self.sheets.len() + 1)
                } else {
                    rest.join(" ")
                };
                match self.add_sheet(&name) {
                    Ok(()) => self.status = format!("Added sheet '{name}'"),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            ["tabnext"] | ["tabn"] => {
                self.next_sheet();
                return CommandOutcome::Continue;
            }
            ["tabprev"] | ["tabp"] | ["tabprevious"] | ["tabN"] => {
                self.prev_sheet();
                return CommandOutcome::Continue;
            }
            ["tabfirst"] | ["tabfir"] | ["tabrewind"] | ["tabr"] => {
                self.switch_sheet(0);
                return CommandOutcome::Continue;
            }
            ["tablast"] | ["tabl"] => {
                self.switch_sheet(self.sheets.len() - 1);
                return CommandOutcome::Continue;
            }
            ["tabclose"] | ["tabc"] => {
                let name = self.active_sheet_name().to_string();
                match self.delete_active_sheet() {
                    Ok(()) => self.status = format!("Deleted sheet '{name}'"),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            // V5 vim `:noh` clears the search highlight.
            ["noh"] | ["nohlsearch"] => {
                self.clear_search();
                return CommandOutcome::Continue;
            }
            // Tutor: `:next` / `:prev` shorthand for sheet nav, and
            // `:reset` reverts the active sheet to its lesson seed.
            ["next"] => {
                self.next_sheet();
                return CommandOutcome::Continue;
            }
            ["prev"] | ["previous"] => {
                self.prev_sheet();
                return CommandOutcome::Continue;
            }
            ["reset"] => {
                if crate::tutor::reset_active_sheet(self) {
                    self.status = "Lesson reset".into();
                } else {
                    self.status = "Not a tutor sheet".into();
                }
                return CommandOutcome::Continue;
            }
            // V8 `:colwidth <n>` / `:colwidth <letter> <n>` /
            // `:colwidth auto` / `:colwidth <letter> auto`. Sets a
            // column's width (1..=80) in characters.
            ["colwidth", arg] => {
                let col_idx = self.cursor_col;
                self.handle_colwidth(col_idx, arg);
                return CommandOutcome::Continue;
            }
            ["colwidth", letter, arg] => {
                let cell_id = format!("{}1", letter.to_uppercase());
                match from_cell_id(&cell_id) {
                    Some((_, col_idx)) => self.handle_colwidth(col_idx, arg),
                    None => self.status = format!("Unknown column: {letter}"),
                }
                return CommandOutcome::Continue;
            }
            // Cell formatting. `:fmt usd [N]` applies USD with the given
            // decimals (default 2). `:fmt+` / `:fmt-` bump decimals.
            // `:fmt clear` (alias `:fmt none`) drops every format axis.
            // All :fmt commands merge — they leave other axes alone.
            ["fmt", "usd"] => {
                self.apply_format_update(|f| {
                    f.number = Some(format::NumberFormat::Usd { decimals: 2 });
                });
                if !self.status.starts_with("Error")
                    && self.status != "No cells to format"
                {
                    self.status = "USD format applied (decimals=2)".into();
                }
                return CommandOutcome::Continue;
            }
            ["fmt", "usd", n] => {
                match n.parse::<u8>() {
                    Ok(decimals) if decimals <= 10 => {
                        self.apply_format_update(|f| {
                            f.number = Some(format::NumberFormat::Usd { decimals });
                        });
                        if !self.status.starts_with("Error")
                            && self.status != "No cells to format"
                        {
                            self.status =
                                format!("USD format applied (decimals={decimals})");
                        }
                    }
                    _ => self.status = format!("Bad decimals (0-10): {n}"),
                }
                return CommandOutcome::Continue;
            }
            ["fmt+"] => {
                self.bump_format_decimals(1);
                return CommandOutcome::Continue;
            }
            ["fmt-"] => {
                self.bump_format_decimals(-1);
                return CommandOutcome::Continue;
            }
            ["fmt", "clear"] | ["fmt", "none"] => {
                self.clear_format_in_selection();
                return CommandOutcome::Continue;
            }
            // Text styles. Each command sets the named flag; the `no…`
            // form clears it. Other format axes are preserved.
            ["fmt", "bold"] => {
                self.apply_format_update(|f| f.bold = true);
                return CommandOutcome::Continue;
            }
            ["fmt", "nobold"] => {
                self.apply_format_update(|f| f.bold = false);
                return CommandOutcome::Continue;
            }
            ["fmt", "italic"] => {
                self.apply_format_update(|f| f.italic = true);
                return CommandOutcome::Continue;
            }
            ["fmt", "noitalic"] => {
                self.apply_format_update(|f| f.italic = false);
                return CommandOutcome::Continue;
            }
            ["fmt", "underline"] => {
                self.apply_format_update(|f| f.underline = true);
                return CommandOutcome::Continue;
            }
            ["fmt", "nounderline"] => {
                self.apply_format_update(|f| f.underline = false);
                return CommandOutcome::Continue;
            }
            ["fmt", "strike"] => {
                self.apply_format_update(|f| f.strike = true);
                return CommandOutcome::Continue;
            }
            ["fmt", "nostrike"] => {
                self.apply_format_update(|f| f.strike = false);
                return CommandOutcome::Continue;
            }
            // Alignment. `:fmt auto` (or noalign) clears the explicit
            // override so classify_display picks (number = right,
            // boolean = center, text/error = left).
            ["fmt", "left"] => {
                self.apply_format_update(|f| f.align = Some(format::Align::Left));
                return CommandOutcome::Continue;
            }
            ["fmt", "center"] => {
                self.apply_format_update(|f| f.align = Some(format::Align::Center));
                return CommandOutcome::Continue;
            }
            ["fmt", "right"] => {
                self.apply_format_update(|f| f.align = Some(format::Align::Right));
                return CommandOutcome::Continue;
            }
            ["fmt", "auto"] | ["fmt", "noalign"] => {
                self.apply_format_update(|f| f.align = None);
                return CommandOutcome::Continue;
            }
            // Per-cell strftime for `lotus-datetime` cells. `:fmt date
            // %Y-%m-%d` overrides the handler's default friendly form.
            // `:fmt nodate` clears the override. Ignored for non-datetime
            // values (the override stays attached but only renders when
            // the cell becomes a datetime).
            #[cfg(feature = "datetime")]
            ["fmt", "date", rest @ ..] if !rest.is_empty() => {
                let pattern = rest.join(" ");
                self.apply_format_update(|f| f.date = Some(pattern.clone()));
                if !self.status.starts_with("Error")
                    && self.status != "No cells to format"
                {
                    self.status = format!("Date format applied: {pattern}");
                }
                return CommandOutcome::Continue;
            }
            #[cfg(feature = "datetime")]
            ["fmt", "date"] => {
                self.status = "Use :fmt date <strftime>  (e.g. :fmt date %a %b %d)".into();
                return CommandOutcome::Continue;
            }
            #[cfg(feature = "datetime")]
            ["fmt", "nodate"] => {
                self.apply_format_update(|f| f.date = None);
                return CommandOutcome::Continue;
            }
            // Insert today's date / current datetime as a literal at the
            // cursor (Excel ctrl+; / ctrl+shift+; convention). The
            // datetime extension auto-detects the resulting ISO string
            // so the cell lands as jdate / jdatetime.
            #[cfg(feature = "datetime")]
            ["today"] => {
                self.insert_today_literal();
                return CommandOutcome::Continue;
            }
            #[cfg(feature = "datetime")]
            ["now"] => {
                self.insert_now_literal();
                return CommandOutcome::Continue;
            }
            // Percent number format. Default 0 decimals (matches gsheets's
            // `Ctrl+Shift+5` behavior: `0.05` → `5%`).
            ["fmt", "percent"] => {
                self.apply_format_update(|f| {
                    f.number = Some(format::NumberFormat::Percent { decimals: 0 });
                });
                return CommandOutcome::Continue;
            }
            ["fmt", "percent", n] => {
                match n.parse::<u8>() {
                    Ok(decimals) if decimals <= 10 => {
                        self.apply_format_update(|f| {
                            f.number =
                                Some(format::NumberFormat::Percent { decimals });
                        });
                    }
                    _ => self.status = format!("Bad decimals (0-10): {n}"),
                }
                return CommandOutcome::Continue;
            }
            // Colors. Preset names (case-insensitive) and hex `#rgb` /
            // `#rrggbb` both accepted. `:fmt nofg` / `:fmt nobg` clear
            // just that axis (other format axes preserved).
            ["fmt", "fg", arg] => {
                match format::parse_color_arg(arg) {
                    Some(c) => self.apply_format_update(|f| f.fg = Some(c)),
                    None => self.status = format!("Bad color: {arg}"),
                }
                return CommandOutcome::Continue;
            }
            ["fmt", "bg", arg] => {
                match format::parse_color_arg(arg) {
                    Some(c) => self.apply_format_update(|f| f.bg = Some(c)),
                    None => self.status = format!("Bad color: {arg}"),
                }
                return CommandOutcome::Continue;
            }
            ["fmt", "nofg"] => {
                self.apply_format_update(|f| f.fg = None);
                return CommandOutcome::Continue;
            }
            ["fmt", "nobg"] => {
                self.apply_format_update(|f| f.bg = None);
                return CommandOutcome::Continue;
            }
            // `:set mouse` / `:set nomouse` / `:set mouse?` — toggle
            // mouse capture. **On by default** (App::new); `:set nomouse`
            // releases capture for the rest of the session so the
            // terminal's native text-selection (Cmd/Option+drag → copy)
            // works again. The actual EnableMouseCapture / DisableMouseCapture
            // call is driven by `run_loop`, which diffs this flag each tick.
            ["set", "mouse"] => {
                self.mouse_enabled = true;
                self.status = "mouse capture enabled".into();
                return CommandOutcome::Continue;
            }
            ["set", "nomouse"] => {
                self.mouse_enabled = false;
                self.status = "mouse capture disabled".into();
                return CommandOutcome::Continue;
            }
            ["set", "mouse?"] => {
                self.status = if self.mouse_enabled {
                    "mouse=on".into()
                } else {
                    "mouse=off".into()
                };
                return CommandOutcome::Continue;
            }
            // V8 `:w <path>` exports the active sheet by extension.
            ["w" | "write", path] => {
                match self.export_sheet(path) {
                    Ok(()) => self.status = format!("Wrote {path}"),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            // T6 dirty-buffer commit/rollback. `:w` / `:wq` / `:x` /
            // bang-variants explicitly commit the long-lived
            // `BEGIN IMMEDIATE` txn; `:q!` discards it.
            ["w" | "write"] => {
                match self.store.commit() {
                    Ok(()) => self.status = "saved".into(),
                    Err(e) => self.status = format!("Error: {e}"),
                }
                return CommandOutcome::Continue;
            }
            ["wq" | "x"] => {
                if let Err(e) = self.store.commit() {
                    self.status = format!("Error: {e}");
                    return CommandOutcome::Continue;
                }
                return self.guarded_quit();
            }
            ["wq!" | "x!"] => {
                if let Err(e) = self.store.commit() {
                    self.status = format!("Error: {e}");
                    return CommandOutcome::Continue;
                }
                return CommandOutcome::Quit { force: true };
            }
            ["q!" | "quit!"] => {
                if let Err(e) = self.store.rollback() {
                    self.status = format!("Error: {e}");
                    return CommandOutcome::Continue;
                }
                return CommandOutcome::Quit { force: true };
            }
            // V8 `:goto <cell>` / `:cell <cell>` for the long form.
            ["goto", target] | ["cell", target] => {
                if !self.jump_to_target(target) {
                    self.status = format!("Bad target: {target}");
                }
                return CommandOutcome::Continue;
            }
            ["help"] => {
                self.status = HELP_TEXT.into();
                return CommandOutcome::Continue;
            }
            // Bare numeric (`:42`) jumps to row N (1-indexed).
            [token] if token.parse::<u32>().is_ok() => {
                let n: u32 = token.parse().unwrap();
                self.goto_row(n.saturating_sub(1));
                return CommandOutcome::Continue;
            }
            // Bare cell-id (`:A1`, `:Z9`) jumps to that cell.
            [token] if from_cell_id(&token.to_uppercase()).is_some() => {
                if let Some((r, c)) = from_cell_id(&token.to_uppercase()) {
                    self.jump_cursor_to(r, c);
                }
                return CommandOutcome::Continue;
            }
            // Patch session lifecycle. P1 of att coisl46s.
            ["patch", "new", path] => {
                self.run_patch_new(path);
                return CommandOutcome::Continue;
            }
            ["patch", "save"] => {
                self.run_patch_save(crate::store::PatchSaveMode::KeepDirty);
                return CommandOutcome::Continue;
            }
            ["patch", "save", "--commit"] => {
                self.run_patch_save(crate::store::PatchSaveMode::Commit);
                return CommandOutcome::Continue;
            }
            ["patch", "save", "--rollback"] => {
                self.run_patch_save(crate::store::PatchSaveMode::Rollback);
                return CommandOutcome::Continue;
            }
            ["patch", "close"] => {
                self.run_patch_close();
                return CommandOutcome::Continue;
            }
            ["patch", "detach"] => {
                self.run_patch_detach();
                return CommandOutcome::Continue;
            }
            ["patch", "invert"] => {
                self.run_patch_invert();
                return CommandOutcome::Continue;
            }
            ["patch", "pause"] => {
                self.store.patch_set_enabled(false);
                self.status = if self.store.patch_status().is_some() {
                    "patch: paused".into()
                } else {
                    "no active patch".into()
                };
                return CommandOutcome::Continue;
            }
            ["patch", "resume"] => {
                self.store.patch_set_enabled(true);
                self.status = if self.store.patch_status().is_some() {
                    "patch: resumed".into()
                } else {
                    "no active patch".into()
                };
                return CommandOutcome::Continue;
            }
            ["patch", "status"] => {
                self.status = match self.store.patch_status() {
                    Some(s) => format!(
                        "patch: {}{}",
                        s.path.display(),
                        if s.paused { " (paused)" } else { "" }
                    ),
                    None => "no active patch".into(),
                };
                return CommandOutcome::Continue;
            }
            ["patch", "apply", path] => {
                self.run_patch_apply(path);
                return CommandOutcome::Continue;
            }
            ["patch", "break", path] => {
                self.run_patch_break(path);
                return CommandOutcome::Continue;
            }
            ["patch", "show"] => {
                self.run_patch_show();
                return CommandOutcome::Continue;
            }
            _ => {}
        }

        match execute_command(&cmd) {
            Ok(CommandOutcome::Quit { force: false }) => self.guarded_quit(),
            Ok(o) => o,
            Err(msg) => {
                self.status = msg;
                CommandOutcome::Continue
            }
        }
    }

    // ── Patch sessions (att coisl46s P1) ─────────────────────────────

    fn run_patch_new(&mut self, path: &str) {
        let path = std::path::PathBuf::from(path);
        match self.store.patch_open(path.clone()) {
            Ok(()) => self.status = format!("patch: recording → {}", path.display()),
            Err(e) => self.status = format!("Error: {e}"),
        }
    }

    fn run_patch_save(&mut self, mode: crate::store::PatchSaveMode) {
        match self.store.patch_save(mode) {
            Ok(p) => {
                let suffix = match mode {
                    crate::store::PatchSaveMode::KeepDirty => "",
                    crate::store::PatchSaveMode::Commit => " (committed)",
                    crate::store::PatchSaveMode::Rollback => " (rolled back)",
                };
                self.status = format!("saved patch → {}{}", p.display(), suffix);
            }
            Err(e) => self.status = format!("Error: {e}"),
        }
    }

    fn run_patch_close(&mut self) {
        match self.store.patch_close() {
            Ok(p) => self.status = format!("patch closed → {}", p.display()),
            Err(e) => self.status = format!("Error: {e}"),
        }
    }

    fn run_patch_detach(&mut self) {
        match self.store.patch_detach() {
            Ok(()) => self.status = "patch detached (changes not saved)".into(),
            Err(e) => self.status = format!("Error: {e}"),
        }
    }

    fn run_patch_invert(&mut self) {
        match self.store.patch_invert() {
            Ok(()) => {
                self.refresh_cells();
                self.refresh_columns();
                self.status = "patch inverted (workbook reverted to recording start)".into();
            }
            Err(e) => self.status = format!("Error: {e}"),
        }
    }

    fn run_patch_apply(&mut self, path: &str) {
        let bytes = match crate::store::patch::read_patch_file(std::path::Path::new(path)) {
            Ok(b) => b,
            Err(e) => {
                self.status = format!("Error: read {path}: {e}");
                return;
            }
        };
        // `Store::apply_changeset` already wraps the apply in
        // `with_session_disabled`, so an external patch we apply
        // mid-recording won't fold into our own session.
        let conflicts = match self
            .store
            .apply_changeset(&bytes, crate::store::ConflictPolicy::Omit)
        {
            Ok(n) => n,
            Err(e) => {
                self.status = format!("Error: apply: {e}");
                return;
            }
        };
        let sheet_names: Vec<String> = self
            .store
            .list_sheets()
            .map(|sheets| sheets.into_iter().map(|m| m.name).collect())
            .unwrap_or_default();
        for name in &sheet_names {
            if let Err(e) = self.store.recalculate(name) {
                self.status = format!("Error: recalculate {name}: {e}");
                return;
            }
        }
        // Reload the active sheet so refresh_cells picks up the new
        // cells/values.
        self.refresh_cells();
        self.refresh_columns();
        self.mark_touched();
        self.status = if conflicts == 0 {
            format!("applied {path} (review then :w or :q!)")
        } else {
            format!(
                "applied {path}, {conflicts} conflict(s) omitted (review then :w or :q!)"
            )
        };
    }

    fn run_patch_break(&mut self, path: &str) {
        // Save the current patch in keep-dirty mode (preserves the
        // workbook txn, just flushes the changeset), close it, and
        // open a fresh patch at `path`. Useful for chunking a long
        // edit session into bite-sized patches.
        if let Err(e) = self.store.patch_save(crate::store::PatchSaveMode::KeepDirty) {
            self.status = format!("Error: save current patch: {e}");
            return;
        }
        if let Err(e) = self.store.patch_close() {
            self.status = format!("Error: close current patch: {e}");
            return;
        }
        let new_path = std::path::PathBuf::from(path);
        match self.store.patch_open(new_path.clone()) {
            Ok(()) => self.status = format!("patch break → {}", new_path.display()),
            Err(e) => self.status = format!("Error: open new patch: {e}"),
        }
    }

    fn run_patch_show(&mut self) {
        // Materialise the session's current changeset into bytes by
        // saving to a tmp file, reading back, and rendering through
        // patch_cli's renderer. Cheaper than building a separate
        // "render in-place" path; the bytes go away with the tmp
        // file. This is also why `:patch show` requires an active
        // patch — without one there's no in-progress changeset to
        // render.
        let tmp_path = match self.store.patch_status() {
            Some(s) => s.path.with_extension("lpatch.show"),
            None => {
                self.status = "no active patch".into();
                return;
            }
        };
        if let Err(e) = self
            .store
            .patch_save(crate::store::PatchSaveMode::KeepDirty)
        {
            self.status = format!("Error: snapshot for show: {e}");
            return;
        }
        // After patch_save the live patch path holds the snapshot;
        // reuse that. (We don't write to a separate path because
        // patch_save targets `patch.path`.)
        let live_path = self
            .store
            .patch_status()
            .map(|s| s.path)
            .unwrap_or_else(|| tmp_path.clone());
        let bytes = match crate::store::patch::read_patch_file(&live_path) {
            Ok(b) => b,
            Err(e) => {
                self.status = format!("Error: read patch: {e}");
                return;
            }
        };
        let mut out: Vec<u8> = Vec::new();
        if let Err(e) = crate::patch_cli::render_to(&bytes, &mut out) {
            self.status = format!("Error: render: {e}");
            return;
        }
        let text = String::from_utf8_lossy(&out).into_owned();
        let lines: Vec<String> = text
            .lines()
            .map(|s| s.to_string())
            .collect();
        if lines.is_empty() {
            self.status = "patch is empty".into();
        } else {
            self.patch_show = Some(PatchShowState {
                lines,
                scroll: 0,
            });
            self.mode = Mode::PatchShow;
        }
    }

    /// Apply the unsaved-changes / in-memory-session guards before
    /// returning `Quit`. Refuses with a status-bar warning if either
    /// guard fires; the user retries with the bang-suffix variant
    /// (`:q!` / `:wq!` / `:x!`) to force.
    ///
    /// The unsaved-changes guard only fires for file-backed DBs.
    /// Ephemeral sessions (`:memory:` / `tutor`) have no meaningful
    /// `:w` target, so the in-memory warning carries the message and
    /// "no write since last change" would be misleading.
    fn guarded_quit(&mut self) -> CommandOutcome {
        let ephemeral = self.db_label == ":memory:" || self.db_label == "tutor";
        if !ephemeral && self.has_unsaved_changes() {
            self.status = "No write since last change (add ! to override)".into();
            return CommandOutcome::Continue;
        }
        if self.should_warn_in_memory() {
            self.status = "In-memory session — changes will be lost (add ! to override)".into();
            return CommandOutcome::Continue;
        }
        CommandOutcome::Quit { force: false }
    }

    // ── Edit buffer cursor helpers ──────────────────────────────────

    pub fn edit_move_left(&mut self) {
        if self.edit_cursor > 0 {
            // Step back one char (handle multi-byte)
            self.edit_cursor = self.edit_buf[..self.edit_cursor]
                .char_indices()
                .next_back()
                .map(|(i, _)| i)
                .unwrap_or(0);
        }
        self.refresh_completions();
    }

    pub fn edit_move_right(&mut self) {
        if self.edit_cursor < self.edit_buf.len() {
            self.edit_cursor = self.edit_buf[self.edit_cursor..]
                .char_indices()
                .nth(1)
                .map(|(i, _)| self.edit_cursor + i)
                .unwrap_or(self.edit_buf.len());
        }
        self.refresh_completions();
    }

    pub fn edit_home(&mut self) {
        self.edit_cursor = 0;
        self.refresh_completions();
    }

    pub fn edit_end(&mut self) {
        self.edit_cursor = self.edit_buf.len();
        self.refresh_completions();
    }

    pub fn edit_insert(&mut self, c: char) {
        self.edit_buf.insert(self.edit_cursor, c);
        self.edit_cursor += c.len_utf8();
        self.refresh_completions();
    }

    pub fn edit_backspace(&mut self) {
        if self.edit_cursor > 0 {
            let prev = self.edit_buf[..self.edit_cursor]
                .char_indices()
                .next_back()
                .map(|(i, _)| i)
                .unwrap_or(0);
            self.edit_buf.drain(prev..self.edit_cursor);
            self.edit_cursor = prev;
        }
        self.refresh_completions();
    }

    pub fn edit_delete(&mut self) {
        if self.edit_cursor < self.edit_buf.len() {
            let next = self.edit_buf[self.edit_cursor..]
                .char_indices()
                .nth(1)
                .map(|(i, _)| self.edit_cursor + i)
                .unwrap_or(self.edit_buf.len());
            self.edit_buf.drain(self.edit_cursor..next);
        }
        self.refresh_completions();
    }

    // [sheet.clipboard.paste]
    // [sheet.clipboard.paste-formula-shift]
    // [sheet.clipboard.paste-fill-selection]
    /// Apply a parsed clipboard grid to the sheet, optionally clearing a
    /// cut source range in the same atomic operation (one undo entry —
    /// see [sheet.undo.scope]).
    ///
    /// The actual paste target depends on selection state:
    /// - 1×1 source pasted into a multi-cell selection: tile across every
    ///   selected cell, shifting formulas per-target.
    /// - Otherwise: rect-paste anchored at the cursor; one delta applies
    ///   to every cell in the rect.
    ///
    /// `cut_source` is a rect to clear *after* the paste lands, skipping
    /// cells that fall inside the paste target (the "non-overlapping
    /// sources" clause of [sheet.clipboard.paste]).
    pub fn apply_pasted_grid(
        &mut self,
        grid: &PastedGrid,
        cut_source: Option<(u32, u32, u32, u32)>,
    ) -> Result<(u32, u32), String> {
        if grid.cells.is_empty() {
            return Ok((0, 0));
        }

        let (rows, cols, mut changes, target_rect) = self.build_paste_changes(grid);

        if let Some((r1, c1, r2, c2)) = cut_source {
            let in_target = |r: u32, c: u32| {
                let (tr1, tc1, tr2, tc2) = target_rect;
                r >= tr1 && r <= tr2 && c >= tc1 && c <= tc2
            };
            for r in r1..=r2 {
                for c in c1..=c2 {
                    if in_target(r, c) {
                        continue;
                    }
                    changes.push(CellChange {
                        row_idx: r,
                        col_idx: c,
                        raw_value: String::new(),
                        format_json: None,
                    });
                }
            }
        }

        self.apply_changes_recorded(&changes)?;
        Ok((rows, cols))
    }

    /// Pure helper: build the cell changes for a paste, plus the target
    /// rect those changes occupy (used by cut-clear to skip overlap).
    /// Returns `(rows_written, cols_written, changes, target_rect)`.
    fn build_paste_changes(
        &self,
        grid: &PastedGrid,
    ) -> (u32, u32, Vec<CellChange>, (u32, u32, u32, u32)) {
        let cursor = (self.cursor_row, self.cursor_col);
        let sel = self.selection_range();

        // [sheet.clipboard.paste-fill-selection]
        if grid.is_single_cell() {
            if let Some((sr1, sc1, sr2, sc2)) = sel {
                if sr1 != sr2 || sc1 != sc2 {
                    let cell = &grid.cells[0][0];
                    let mut changes = Vec::new();
                    for tr in sr1..=sr2 {
                        for tc in sc1..=sc2 {
                            if tr > MAX_ROW || tc > MAX_COL {
                                continue;
                            }
                            let raw = shift_pasted_raw(cell, grid.source_anchor, (tr, tc));
                            changes.push(CellChange {
                                row_idx: tr,
                                col_idx: tc,
                                raw_value: raw,
                                format_json: None,
                            });
                        }
                    }
                    return (sr2 - sr1 + 1, sc2 - sc1 + 1, changes, (sr1, sc1, sr2, sc2));
                }
            }
        }

        // Standard rect paste anchored at the cursor.
        let mut changes = Vec::new();
        let rows = grid.rows();
        let cols = grid.cols();
        for (ri, row_cells) in grid.cells.iter().enumerate() {
            for (ci, cell) in row_cells.iter().enumerate() {
                let tr = cursor.0 + ri as u32;
                let tc = cursor.1 + ci as u32;
                if tr > MAX_ROW || tc > MAX_COL {
                    continue;
                }
                let source_pos = grid
                    .source_anchor
                    .map(|(ar, ac)| (ar + ri as u32, ac + ci as u32));
                let raw = shift_pasted_raw(cell, source_pos, (tr, tc));
                changes.push(CellChange {
                    row_idx: tr,
                    col_idx: tc,
                    raw_value: raw,
                    format_json: None,
                });
            }
        }
        let target_rect = (
            cursor.0,
            cursor.1,
            (cursor.0 + rows.saturating_sub(1)).min(MAX_ROW),
            (cursor.1 + cols.saturating_sub(1)).min(MAX_COL),
        );
        (rows, cols, changes, target_rect)
    }

    /// Insert today's date as an ISO literal at the cursor cell.
    /// Bound to `:today` (ex command) and `ctrl+;` (Excel/Sheets
    /// convention). The datetime extension's auto-detect path lands
    /// the value as a `jdate`. Pinned in time — re-opening the
    /// workbook tomorrow keeps yesterday's date (use `=TODAY()` for
    /// the recomputing version).
    #[cfg(feature = "datetime")]
    pub fn insert_today_literal(&mut self) {
        let iso = lotus_datetime::today_iso();
        if let Err(e) = self.apply_changes_recorded(&[CellChange {
            row_idx: self.cursor_row,
            col_idx: self.cursor_col,
            raw_value: iso.clone(),
            format_json: None,
        }]) {
            self.status = format!("Error: {e}");
        } else {
            self.status = format!("Inserted {iso}");
        }
    }

    /// Insert the current local datetime as an ISO literal at the
    /// cursor cell. Bound to `:now` and `ctrl+shift+;`. Lands as a
    /// `jdatetime` via auto-detect; second-precision (no subseconds).
    #[cfg(feature = "datetime")]
    pub fn insert_now_literal(&mut self) {
        let iso = lotus_datetime::now_iso();
        if let Err(e) = self.apply_changes_recorded(&[CellChange {
            row_idx: self.cursor_row,
            col_idx: self.cursor_col,
            raw_value: iso.clone(),
            format_json: None,
        }]) {
            self.status = format!("Error: {e}");
        } else {
            self.status = format!("Inserted {iso}");
        }
    }

    // [sheet.delete.delete-key-clears]
    pub fn delete_cell(&mut self) {
        match self.apply_changes_recorded(&[CellChange {
            row_idx: self.cursor_row,
            col_idx: self.cursor_col,
            raw_value: String::new(),
            format_json: None,
        }]) {
            Ok(()) => {
                self.status = "Deleted".into();
                self.last_edit = Some(EditAction {
                    kind: EditKind::Delete,
                    anchor_dr: 0,
                    anchor_dc: 0,
                    rect_rows: 1,
                    rect_cols: 1,
                    text: None,
                });
            }
            Err(e) => self.status = format!("Error: {e}"),
        }
    }

    pub fn refresh_cells(&mut self) {
        self.cells = self
            .store
            .load_sheet(self.active_sheet_name())
            .map(|s| s.cells)
            .unwrap_or_default();
    }

    pub fn refresh_columns(&mut self) {
        self.columns = self
            .store
            .load_columns(self.active_sheet_name())
            .unwrap_or_default();
    }

    /// Width (in characters) of `col_idx`. Looks up the schema cache;
    /// defaults to `DEFAULT_COL_WIDTH` for columns without a row.
    pub fn column_width(&self, col_idx: u32) -> u16 {
        self.columns
            .iter()
            .find(|c| c.col_idx == col_idx)
            .map(|c| (c.width as u16).clamp(MIN_COL_WIDTH, MAX_COL_WIDTH))
            .unwrap_or(DEFAULT_COL_WIDTH)
    }

    /// Persist a new width for `col_idx` and refresh the cache.
    pub fn set_column_width(&mut self, col_idx: u32, width: u16) -> Result<(), String> {
        let clamped = width.clamp(MIN_COL_WIDTH, MAX_COL_WIDTH);
        let name = self.active_sheet_name().to_string();
        self.store
            .set_column_width(&name, col_idx, clamped as u32)
            .map_err(|e| e.to_string())?;
        self.refresh_columns();
        self.mark_touched();
        Ok(())
    }

    /// Vim `:colwidth auto`: fit `col_idx` to the longest displayed
    /// value in the column, plus a 1-char breathing-room pad. Empty
    /// columns fall back to `DEFAULT_COL_WIDTH`.
    pub fn autofit_column(&mut self, col_idx: u32) -> Result<u16, String> {
        let longest = self
            .cells
            .iter()
            .filter(|c| c.col_idx == col_idx)
            .map(|c| self.displayed_for(c.row_idx, c.col_idx).chars().count())
            .max()
            .unwrap_or(0);
        let width = if longest == 0 {
            DEFAULT_COL_WIDTH
        } else {
            ((longest + 1) as u16).clamp(MIN_COL_WIDTH, MAX_COL_WIDTH)
        };
        self.set_column_width(col_idx, width)?;
        Ok(width)
    }

    /// Get computed display value for a cell (no format applied — this
    /// is the engine's stringified result, used for stats, clipboard,
    /// and search). For the grid-rendered string see [`displayed_for`].
    pub fn get_display(&self, row: u32, col: u32) -> String {
        self.cells
            .iter()
            .find(|c| c.row_idx == row && c.col_idx == col)
            .and_then(|c| c.computed_value.clone())
            .unwrap_or_default()
    }

    /// Get the cell's parsed format, if any.
    pub fn get_format(&self, row: u32, col: u32) -> Option<format::CellFormat> {
        self.get_format_json_raw(row, col)
            .as_deref()
            .and_then(format::parse_format_json)
    }

    /// Render a cell for grid display, applying its format if one is
    /// set. Used by both `ui::draw_grid` and `autofit_column` so the
    /// rendered width and the laid-out width agree.
    ///
    /// Hyperlink-typed cells store a `url + SEP + label` payload in
    /// `cell.computed`; this method strips the URL portion so the
    /// grid shows the user-visible label. Plain-text URL cells pass
    /// through unchanged.
    pub fn displayed_for(&self, row: u32, col: u32) -> String {
        let raw = self.get_display(row, col);
        let display = match hyperlink::split_payload(&raw) {
            Some((_url, label)) => label.to_string(),
            None => raw,
        };
        let fmt = self.get_format(row, col);
        // Per-cell strftime override (`:fmt date <pattern>`): apply
        // before format::render so style/align/color still layer on the
        // strftime output. Silently falls back to the friendly display
        // if the cell isn't a datetime or strftime fails — the user
        // sees the colour-coded peach styling either way.
        #[cfg(feature = "datetime")]
        let display = self.apply_date_format(row, col, display, fmt.as_ref());
        match fmt {
            Some(fmt) => format::render(&display, &fmt),
            None => display,
        }
    }

    /// Apply `fmt.date`'s strftime pattern to `default` when the cell
    /// at `(row, col)` is a `lotus-datetime` value. Pure passthrough
    /// when the cell isn't a datetime, when there's no format, or when
    /// `fmt.date` is unset.
    #[cfg(feature = "datetime")]
    fn apply_date_format(
        &self,
        row: u32,
        col: u32,
        default: String,
        fmt: Option<&format::CellFormat>,
    ) -> String {
        let Some(pattern) = fmt.and_then(|f| f.date.as_deref()) else {
            return default;
        };
        let Some(cell) = self
            .store
            .cell_custom(self.active_sheet_name(), row, col)
        else {
            return default;
        };
        let cv = lotus_core::CustomValue {
            type_tag: cell.type_tag.clone(),
            data: cell.data.clone(),
        };
        lotus_datetime::format_custom_value(&cv, pattern).unwrap_or(default)
    }

    /// URL associated with the cell at `(row, col)`, if any. Returns
    /// the URL portion of a hyperlink custom value when present; falls
    /// back to the cell's displayed string when that string itself
    /// looks like a URL. Used by `go` and ctrl-click to know what to
    /// open and by `ui::draw_grid` to decide which cells get the
    /// hyperlink underline + cyan styling.
    pub fn url_for_cell(&self, row: u32, col: u32) -> Option<String> {
        let raw = self.get_display(row, col);
        if let Some((url, _label)) = hyperlink::split_payload(&raw) {
            return Some(url.to_string());
        }
        if hyperlink::looks_like_url(&raw) {
            return Some(raw);
        }
        None
    }

    /// URL under the cursor, if any. Convenience wrapper over
    /// [`Self::url_for_cell`] used by the `go` keybinding.
    pub fn url_under_cursor(&self) -> Option<String> {
        self.url_for_cell(self.cursor_row, self.cursor_col)
    }

    /// Datetime type tag (`"jdate"`, `"jspan"`, …) of the cell at
    /// `(row, col)`, when its computed value is one of the six types
    /// shipped by `lotus-datetime`. Returns `None` for empty cells,
    /// built-in scalar types, hyperlinks, and any other custom type
    /// that isn't a datetime. Drives the peach auto-style branch in
    /// `ui::draw_grid`. Backed by the in-memory cache
    /// `Store::custom_cells` populated during `recalculate`.
    #[cfg(feature = "datetime")]
    pub fn datetime_tag_for_cell(&self, row: u32, col: u32) -> Option<&str> {
        let cell = self
            .store
            .cell_custom(self.active_sheet_name(), row, col)?;
        crate::datetime::is_datetime_tag(&cell.type_tag).then_some(cell.type_tag.as_str())
    }

    /// Apply a per-axis update to every cell in the current selection
    /// (or just the cursor cell when nothing is range-selected). Each
    /// cell's existing format is loaded, mutated in place by `update`,
    /// and persisted; raw values are preserved. Empty cells are
    /// skipped — `set_cells_and_recalculate` deletes empty-raw rows,
    /// so attaching a format to a blank cell would silently disappear.
    ///
    /// This is the merge-aware path: `:fmt usd` on a bold cell keeps
    /// it bold; only the number axis changes. For full-clear semantics
    /// see [`clear_format_in_selection`].
    pub fn apply_format_update<F: FnMut(&mut format::CellFormat)>(
        &mut self,
        mut update: F,
    ) {
        let (r1, c1, r2, c2) = self.effective_rect();
        let mut changes = Vec::new();
        for r in r1..=r2 {
            for c in c1..=c2 {
                let raw = self.get_raw(r, c);
                if raw.is_empty() {
                    continue;
                }
                let mut fmt = self.get_format(r, c).unwrap_or_default();
                update(&mut fmt);
                changes.push(CellChange {
                    row_idx: r,
                    col_idx: c,
                    raw_value: raw,
                    format_json: Some(format::to_format_json(&fmt)),
                });
            }
        }
        if changes.is_empty() {
            self.status = "No cells to format".into();
            return;
        }
        if let Err(e) = self.apply_changes_recorded(&changes) {
            self.status = format!("Error: {e}");
        }
    }

    /// Wipe formatting on every cell in the selection — sends the
    /// `"null"` clear-sentinel which `set_cells_and_recalculate`
    /// interprets as "drop the format_json column to NULL". Raw
    /// values are preserved.
    pub fn clear_format_in_selection(&mut self) {
        let (r1, c1, r2, c2) = self.effective_rect();
        let mut changes = Vec::new();
        for r in r1..=r2 {
            for c in c1..=c2 {
                let raw = self.get_raw(r, c);
                if raw.is_empty() {
                    continue;
                }
                changes.push(CellChange {
                    row_idx: r,
                    col_idx: c,
                    raw_value: raw,
                    format_json: Some("null".to_string()),
                });
            }
        }
        if changes.is_empty() {
            self.status = "No cells to format".into();
            return;
        }
        match self.apply_changes_recorded(&changes) {
            Ok(()) => self.status = "Format cleared".into(),
            Err(e) => self.status = format!("Error: {e}"),
        }
    }

    /// Bump (or set) the decimal count for each cell in the selection.
    /// On cells without a number format yet, applies USD/2 first then
    /// bumps — matches gsheets's toolbar where pressing the +decimals
    /// button on an unformatted number shifts it to USD with one
    /// extra digit. Other format axes (bold, align, color) are
    /// preserved.
    pub fn bump_format_decimals(&mut self, delta: i8) {
        self.apply_format_update(|fmt| {
            let base = fmt.number.unwrap_or(format::NumberFormat::Usd { decimals: 2 });
            fmt.number = Some(format::bump_number_decimals(&base, delta));
        });
        if !self.status.starts_with("Error") && self.status != "No cells to format" {
            self.status = format!("Decimals {:+}", delta);
        }
    }

    // ── Color picker ───────────────────────────────────────────────

    /// Open the modal color picker for the given axis. Captures the
    /// current selection rect so `apply_color_picker` always targets
    /// the same cells even if the user moved the cursor elsewhere
    /// while the picker was open. No-op outside Nav / Visual.
    pub fn open_color_picker(&mut self, kind: ColorPickerKind) {
        if !matches!(self.mode, Mode::Nav | Mode::Visual(_)) {
            return;
        }
        let target_rect = self.effective_rect();
        // Preselect the cell's current color if any, so the highlight
        // lands on the active swatch.
        let cur_color = match kind {
            ColorPickerKind::Fg => self.get_format(target_rect.0, target_rect.1).and_then(|f| f.fg),
            ColorPickerKind::Bg => self.get_format(target_rect.0, target_rect.1).and_then(|f| f.bg),
        };
        let cursor = cur_color
            .and_then(|c| {
                format::COLOR_PRESETS
                    .iter()
                    .position(|(_, preset)| *preset == c)
            })
            .unwrap_or(0);
        self.color_picker = Some(ColorPickerState {
            kind,
            cursor,
            hex_input: None,
            target_rect,
        });
        self.mode = Mode::ColorPicker;
    }

    /// Cancel the picker without applying.
    pub fn close_color_picker(&mut self) {
        self.color_picker = None;
        self.mode = Mode::Nav;
        self.status.clear();
    }

    /// Apply the picker's currently-selected color and close.
    pub fn apply_color_picker(&mut self) {
        let Some(state) = self.color_picker.clone() else {
            return;
        };
        let color = if let Some(ref hex) = state.hex_input {
            match format::parse_hex_color(hex) {
                Some(c) => c,
                None => {
                    self.status = format!("Bad hex: {hex}");
                    return;
                }
            }
        } else {
            format::COLOR_PRESETS[state.cursor].1
        };
        let (r1, c1, r2, c2) = state.target_rect;
        let mut changes = Vec::new();
        for r in r1..=r2 {
            for c in c1..=c2 {
                let raw = self.get_raw(r, c);
                if raw.is_empty() {
                    continue;
                }
                let mut fmt = self.get_format(r, c).unwrap_or_default();
                match state.kind {
                    ColorPickerKind::Fg => fmt.fg = Some(color),
                    ColorPickerKind::Bg => fmt.bg = Some(color),
                }
                changes.push(CellChange {
                    row_idx: r,
                    col_idx: c,
                    raw_value: raw,
                    format_json: Some(format::to_format_json(&fmt)),
                });
            }
        }
        if !changes.is_empty() {
            if let Err(e) = self.apply_changes_recorded(&changes) {
                self.status = format!("Error: {e}");
                return;
            }
        }
        self.close_color_picker();
    }

    /// Move the swatch cursor by `delta` (positive = forward, negative =
    /// back). Wraps within `COLOR_PRESETS`. No-op while in hex-input mode.
    pub fn color_picker_step(&mut self, delta: i32) {
        let Some(state) = self.color_picker.as_mut() else {
            return;
        };
        if state.hex_input.is_some() {
            return;
        }
        let n = format::COLOR_PRESETS.len() as i32;
        let next = ((state.cursor as i32 + delta).rem_euclid(n)) as usize;
        state.cursor = next;
    }

    /// Toggle the picker into / out of hex-entry mode.
    pub fn color_picker_toggle_hex(&mut self) {
        let Some(state) = self.color_picker.as_mut() else {
            return;
        };
        state.hex_input = match state.hex_input.take() {
            Some(_) => None,
            None => Some("#".to_string()),
        };
    }

    /// Append a character to the hex-input buffer (no-op in swatch mode).
    pub fn color_picker_hex_input(&mut self, ch: char) {
        let Some(state) = self.color_picker.as_mut() else {
            return;
        };
        let Some(buf) = state.hex_input.as_mut() else {
            return;
        };
        // Only allow `#`-prefix + hex digits, capped at 7 chars (`#rrggbb`).
        if buf.len() >= 7 {
            return;
        }
        if ch.is_ascii_hexdigit() {
            buf.push(ch);
        }
    }

    pub fn color_picker_hex_backspace(&mut self) {
        let Some(state) = self.color_picker.as_mut() else {
            return;
        };
        let Some(buf) = state.hex_input.as_mut() else {
            return;
        };
        if buf.len() > 1 {
            buf.pop();
        }
    }

    /// Get raw value for a cell.
    pub fn get_raw(&self, row: u32, col: u32) -> String {
        self.cells
            .iter()
            .find(|c| c.row_idx == row && c.col_idx == col)
            .map(|c| c.raw_value.clone())
            .unwrap_or_default()
    }

    /// Get the raw JSON payload of a cell's format, if any. The string is
    /// the value persisted in `datasette_sheets_cell.format_json`; pair
    /// with `format::parse_format_json` to decode.
    pub fn get_format_json_raw(&self, row: u32, col: u32) -> Option<String> {
        self.cells
            .iter()
            .find(|c| c.row_idx == row && c.col_idx == col)
            .and_then(|c| c.format_json.clone())
    }
}

/// Parse an HTML table clipboard payload into a [`PastedGrid`].
///
/// Handles tables from Google Sheets, our own [`App::yank_html`], and
/// similar sources. Captures two app-private markers when present (see
/// [sheet.clipboard.copy] / [sheet.clipboard.paste-formula-shift]):
/// - `data-sheets-source-anchor="B2"` on `<table>` → grid `source_anchor`.
/// - `data-sheets-formula="=…"` on `<td>` / `<th>` → cell `formula`.
///
/// External-app payloads without these markers parse to a plain
/// values-only grid (formula is `None`, source_anchor is `None`).
pub fn parse_pasted_grid(html: &str) -> Option<PastedGrid> {
    let table_start = html.find("<table")?;
    let table_end = html.find("</table>").map(|i| i + 8)?;
    let table = &html[table_start..table_end];

    // Capture the source anchor from the opening <table ...> tag if any.
    let table_tag_end = table.find('>').map(|i| i + 1)?;
    let table_tag = &table[..table_tag_end];
    let source_anchor = extract_attr(table_tag, "data-sheets-source-anchor")
        .as_deref()
        .map(decode_entities)
        .and_then(|id| from_cell_id(&id));

    let mut rows: Vec<Vec<PastedCell>> = Vec::new();
    let mut pos = table_tag_end;

    while pos < table.len() {
        let tr_start = match table[pos..].find("<tr") {
            Some(i) => pos + i,
            None => break,
        };
        let tr_end = match table[tr_start..].find("</tr>") {
            Some(i) => tr_start + i + 5,
            None => break,
        };
        let tr = &table[tr_start..tr_end];

        let mut cells: Vec<PastedCell> = Vec::new();
        let mut cpos = 0;
        while cpos < tr.len() {
            let td_tag_start = match (tr[cpos..].find("<td"), tr[cpos..].find("<th")) {
                (Some(a), Some(b)) => cpos + a.min(b),
                (Some(a), None) => cpos + a,
                (None, Some(b)) => cpos + b,
                (None, None) => break,
            };
            let content_start = match tr[td_tag_start..].find('>') {
                Some(i) => td_tag_start + i + 1,
                None => break,
            };
            let content_end = match (
                tr[content_start..].find("</td>"),
                tr[content_start..].find("</th>"),
            ) {
                (Some(a), Some(b)) => content_start + a.min(b),
                (Some(a), None) => content_start + a,
                (None, Some(b)) => content_start + b,
                (None, None) => break,
            };

            let opening_tag = &tr[td_tag_start..content_start];
            let formula = extract_attr(opening_tag, "data-sheets-formula")
                .as_deref()
                .map(decode_entities);
            let value = strip_tags_and_decode(&tr[content_start..content_end]);

            cells.push(PastedCell { value, formula });

            cpos = content_end + 5; // skip past </td> or </th>
        }

        if !cells.is_empty() {
            rows.push(cells);
        }
        pos = tr_end;
    }

    if rows.is_empty() {
        None
    } else {
        Some(PastedGrid {
            source_anchor,
            cells: rows,
        })
    }
}

/// Extract the value of a `name="..."` attribute from a single tag.
/// Returns the raw (still-escaped) attribute body — call `decode_entities`
/// on it before using.
fn extract_attr(tag: &str, name: &str) -> Option<String> {
    let pat = format!("{name}=\"");
    let start = tag.find(&pat)? + pat.len();
    let end = tag[start..].find('"')? + start;
    Some(tag[start..end].to_string())
}

fn decode_entities(s: &str) -> String {
    s.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
        .replace("&nbsp;", " ")
}

/// Compute the raw cell text for a target cell after applying
/// [sheet.clipboard.paste-formula-shift]: when the source carried a
/// formula AND we know its source position, shift relative refs by the
/// `target - source` delta. Otherwise use the cell's value verbatim.
fn shift_pasted_raw(
    cell: &PastedCell,
    source_pos: Option<(u32, u32)>,
    target: (u32, u32),
) -> String {
    match (&cell.formula, source_pos) {
        (Some(formula), Some((sr, sc))) => {
            let dr = target.0 as i32 - sr as i32;
            let dc = target.1 as i32 - sc as i32;
            if dr == 0 && dc == 0 {
                formula.clone()
            } else {
                // Engine wants 1-based bounds.
                shift_formula_refs(formula, dr, dc, MAX_ROW + 1, MAX_COL + 1)
            }
        }
        (Some(formula), None) => formula.clone(),
        (None, _) => cell.value.clone(),
    }
}

/// Strip HTML tags and decode common entities.
fn strip_tags_and_decode(s: &str) -> String {
    // Remove tags
    let mut result = String::with_capacity(s.len());
    let mut in_tag = false;
    for ch in s.chars() {
        match ch {
            '<' => in_tag = true,
            '>' => in_tag = false,
            _ if !in_tag => result.push(ch),
            _ => {}
        }
    }
    // Decode entities
    result
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&#39;", "'")
        .replace("&nbsp;", " ")
}

fn html_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

// [sheet.editing.formula-ref-pointing]
/// Whether `caret` sits at an "insertable" position in the formula
/// `buf` — a spot where pressing an arrow key should insert a cell
/// reference instead of moving the editor caret.
///
/// Spec heuristic (per the spec's note about deferring a full
/// grammar-aware check): the caret is insertable when
/// 1. the buffer starts with `=`, and
/// 2. the last non-whitespace char before the caret is one of
///    `=`, `+`, `-`, `*`, `/`, `^`, `&`, `<`, `>`, `(`, `,`, and
/// 3. the next non-whitespace char is end-of-buffer, `)` or `,` —
///    so `=SUM(1,| 2)` does not insert (it would produce
///    `=SUM(1,A2 2)` which isn't valid).
pub fn is_insertable_at(buf: &str, caret: usize) -> bool {
    if !buf.starts_with('=') {
        return false;
    }
    if caret == 0 || caret > buf.len() {
        return false;
    }
    let before = buf[..caret].trim_end_matches(|c: char| c.is_whitespace());
    let last_byte = before.bytes().last();
    let left_ok = matches!(
        last_byte,
        Some(b'=' | b'+' | b'-' | b'*' | b'/' | b'^' | b'&' | b'(' | b',' | b'<' | b'>')
    );
    if !left_ok {
        return false;
    }
    let after = buf[caret..].trim_start_matches(|c: char| c.is_whitespace());
    let next_byte = after.bytes().next();
    matches!(next_byte, None | Some(b')' | b','))
}

/// Map the second word of `:row ins …` / `:col ins …` onto an
/// `InsertSide`. Accepts both row-style (`above`/`below`) and col-style
/// (`left`/`right`) — they're symmetric.
pub fn parse_insert_side(s: &str) -> Option<InsertSide> {
    match s {
        "above" | "left" => Some(InsertSide::AboveOrLeft),
        "below" | "right" => Some(InsertSide::BelowOrRight),
        _ => None,
    }
}

/// V8: short status-line summary of the keymap. The full table lives
/// in the README; this is what `:help` shows.
pub const HELP_TEXT: &str =
    "hjkl move | i/a insert | v Visual | d/c/y operators | / search | u undo | :q quit | :w file save";

/// V8: CSV/TSV escaping. CSV needs quoting when the value contains the
/// separator, a quote, or a newline. TSV escapes embedded tabs/newlines
/// by replacing them with spaces — keeps the output one-row-per-line.
fn quote_for_format(value: &str, separator: char) -> String {
    if separator == '\t' {
        return value.replace(['\t', '\n'], " ");
    }
    let needs_quoting = value.contains(separator) || value.contains('"') || value.contains('\n');
    if !needs_quoting {
        return value.to_string();
    }
    let escaped = value.replace('"', "\"\"");
    format!("\"{escaped}\"")
}

/// Pure command parser. Returns `Ok(Quit)` on `:q`/`:quit` (and bang
/// variants), `Ok(Continue)` on no-op commands, `Err(msg)` for an unknown
/// command (the message is shown in the status bar).
///
/// Bang suffix (`:q!` / `:wq!` / `:x!`) sets `Quit { force: true }`,
/// which `App::run_command` uses to bypass the unsaved-changes /
/// in-memory-session guards.
pub fn execute_command(cmd: &str) -> Result<CommandOutcome, String> {
    match cmd {
        "" => Ok(CommandOutcome::Continue),
        "q" | "quit" => Ok(CommandOutcome::Quit { force: false }),
        "q!" | "quit!" => Ok(CommandOutcome::Quit { force: true }),
        // `:w` is a no-op since writes are immediate. Accept it for muscle
        // memory and report success silently.
        "w" | "write" => Ok(CommandOutcome::Continue),
        "wq" | "x" => Ok(CommandOutcome::Quit { force: false }),
        "wq!" | "x!" => Ok(CommandOutcome::Quit { force: true }),
        // V5 vim `:noh` / `:nohlsearch`. Pure-command parsing can't
        // mutate App, so the run-loop watches for these and calls
        // App::clear_search instead.
        "noh" | "nohlsearch" => Ok(CommandOutcome::Continue),
        other => Err(format!("Unknown command: :{other}")),
    }
}

/// Build the same registry `Store::recalculate` builds — hyperlink
/// always, datetime when the cargo feature is on. Returned as an
/// `Arc<Registry>` so the autocomplete popup can hold a long-lived
/// reference cheaply. Failures here would mean the registration code
/// drifted into a self-collision; not actionable at the call site, so
/// we panic with the offending error rather than silently shipping a
/// popup with no functions in it.
fn build_registry() -> Arc<Registry> {
    let mut sheet = Sheet::new();
    hyperlink::register(&mut sheet).expect("hyperlink registry construction");
    #[cfg(feature = "datetime")]
    crate::datetime::register(&mut sheet).expect("datetime registry construction");
    Arc::clone(sheet.registry())
}

/// Load every sheet in the workbook, creating a default `Sheet1` if the
/// database is empty. Always returns a non-empty Vec.
fn ensure_at_least_one_sheet(store: &mut Store) -> Vec<SheetMeta> {
    let existing = store.list_sheets().unwrap_or_default();
    if !existing.is_empty() {
        return existing;
    }
    store
        .create_sheet("Sheet1")
        .expect("create default sheet");
    store
        .list_sheets()
        .expect("list sheets after creating default")
}

/// Compute the target cell for a content-aware jump (Ctrl/Cmd+Arrow).
///
/// Pure function so it can be tested without a `Connection`. Implements the
/// three rules from `sheet.navigation.arrow-jump`:
/// 1. Already at the grid edge in the direction of travel → stay put.
/// 2. Current filled AND neighbour filled → walk forward to the last filled
///    cell of the contiguous run.
/// 3. Otherwise → skip past empties to the first filled cell, or snap to the
///    grid edge if none.
///
/// `dr`/`dc` must be exactly one of `(-1,0)`, `(1,0)`, `(0,-1)`, `(0,1)`.
pub fn compute_jump_target(
    start_row: u32,
    start_col: u32,
    dr: i32,
    dc: i32,
    max_row: u32,
    max_col: u32,
    is_filled: impl Fn(u32, u32) -> bool,
) -> (u32, u32) {
    debug_assert!(
        (dr.abs() + dc.abs()) == 1,
        "compute_jump_target expects a unit direction, got ({dr}, {dc})"
    );

    // Rule 1: at grid edge in the direction of travel → no-op.
    if (dr < 0 && start_row == 0)
        || (dr > 0 && start_row == max_row)
        || (dc < 0 && start_col == 0)
        || (dc > 0 && start_col == max_col)
    {
        return (start_row, start_col);
    }

    // Step the cursor by (dr, dc), clamped to the grid. Returns None at the
    // edge in the direction of travel.
    let step = |r: u32, c: u32| -> Option<(u32, u32)> {
        let nr = r as i32 + dr;
        let nc = c as i32 + dc;
        if nr < 0 || nc < 0 || nr as u32 > max_row || nc as u32 > max_col {
            None
        } else {
            Some((nr as u32, nc as u32))
        }
    };

    let (nr, nc) = step(start_row, start_col).expect("rule 1 already handled");

    if is_filled(start_row, start_col) && is_filled(nr, nc) {
        // Rule 2: walk to the last filled cell in the contiguous run.
        let (mut r, mut c) = (nr, nc);
        while let Some((next_r, next_c)) = step(r, c) {
            if is_filled(next_r, next_c) {
                r = next_r;
                c = next_c;
            } else {
                break;
            }
        }
        (r, c)
    } else {
        // Rule 3: skip empties to the first filled cell, or snap to edge.
        let (mut r, mut c) = (nr, nc);
        loop {
            if is_filled(r, c) {
                return (r, c);
            }
            match step(r, c) {
                Some((next_r, next_c)) => {
                    r = next_r;
                    c = next_c;
                }
                None => return (r, c), // hit the edge with no filled cell
            }
        }
    }
}

/// Vim `w` analog. Forward to the start of the next contiguous filled run.
/// If currently inside a run, walks past the rest of the run; then skips
/// any empties and lands on the first filled cell. Snaps to grid edge if
/// no filled cell exists in the direction of travel.
///
/// Pure function so it's testable without a `Connection`.
pub fn compute_word_forward(
    start_row: u32,
    start_col: u32,
    max_col: u32,
    is_filled: impl Fn(u32, u32) -> bool,
) -> (u32, u32) {
    let step = |c: u32| -> Option<u32> {
        if c >= max_col { None } else { Some(c + 1) }
    };
    let row = start_row;
    let mut c = start_col;

    // Phase 1: if inside a filled run, walk past it to the first non-filled
    // cell beyond the run.
    if is_filled(row, c) {
        loop {
            match step(c) {
                Some(nc) if is_filled(row, nc) => c = nc,
                Some(nc) => {
                    c = nc;
                    break;
                }
                None => return (row, c),
            }
        }
    }
    // Phase 2: skip empties to the first filled cell, or snap to edge.
    loop {
        if is_filled(row, c) {
            return (row, c);
        }
        match step(c) {
            Some(nc) => c = nc,
            None => return (row, c),
        }
    }
}

/// Vim `b` analog. Backward to the start of the current or previous word.
/// If currently inside a run after its first cell, lands on the run's
/// first cell; otherwise steps back through empties to the previous run
/// and lands on its first cell. Stays put when already at column 0.
pub fn compute_word_backward(
    start_row: u32,
    start_col: u32,
    is_filled: impl Fn(u32, u32) -> bool,
) -> (u32, u32) {
    let row = start_row;
    if start_col == 0 {
        return (row, 0);
    }
    let mut c = start_col - 1;

    // Phase 1: if we landed on empty, walk back through empties to the
    // first filled cell. Snap to col 0 if none found.
    if !is_filled(row, c) {
        loop {
            if c == 0 {
                return (row, 0);
            }
            c -= 1;
            if is_filled(row, c) {
                break;
            }
        }
    }
    // Phase 2: walk back to the start of the contiguous run.
    while c > 0 && is_filled(row, c - 1) {
        c -= 1;
    }
    (row, c)
}

/// Ctrl+A "data region": BFS from the seed through 8-connected populated
/// cells, return the bounding box of the connected component. Returns
/// `None` when the seed itself isn't in `populated` — caller decides the
/// fallback (vlotus selects the whole sheet).
///
/// 8-connected (orthogonal + diagonal) matches Google Sheets' Ctrl+A on a
/// data region: two cells touching at a corner are part of the same blob.
/// The bounding box may include cells that are themselves empty — that's
/// intentional, the result is always a rectangle.
pub fn compute_data_region(
    seed_row: u32,
    seed_col: u32,
    populated: &std::collections::HashSet<(u32, u32)>,
) -> Option<(u32, u32, u32, u32)> {
    if !populated.contains(&(seed_row, seed_col)) {
        return None;
    }
    let mut queue: std::collections::VecDeque<(u32, u32)> = std::collections::VecDeque::new();
    let mut visited: std::collections::HashSet<(u32, u32)> = std::collections::HashSet::new();
    queue.push_back((seed_row, seed_col));
    visited.insert((seed_row, seed_col));
    let (mut r1, mut c1) = (seed_row, seed_col);
    let (mut r2, mut c2) = (seed_row, seed_col);

    while let Some((r, c)) = queue.pop_front() {
        r1 = r1.min(r);
        r2 = r2.max(r);
        c1 = c1.min(c);
        c2 = c2.max(c);
        for dr in -1..=1i32 {
            for dc in -1..=1i32 {
                if dr == 0 && dc == 0 {
                    continue;
                }
                let nr = r as i32 + dr;
                let nc = c as i32 + dc;
                if nr < 0 || nc < 0 {
                    continue;
                }
                let n = (nr as u32, nc as u32);
                if populated.contains(&n) && visited.insert(n) {
                    queue.push_back(n);
                }
            }
        }
    }
    Some((r1, c1, r2, c2))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashSet;

    /// Build an in-memory `App` ready for write/undo/redo tests.
    fn make_test_app() -> App {
        let store = Store::open_in_memory().unwrap();
        App::new(store, ":memory:")
    }

    /// Move the cursor (no clamping needed for these tests) and stage an
    /// edit, then commit it. Mirrors what the run-loop does on a key press.
    fn type_into(app: &mut App, row: u32, col: u32, value: &str) {
        app.cursor_row = row;
        app.cursor_col = col;
        app.start_edit_blank();
        for c in value.chars() {
            app.edit_insert(c);
        }
        app.confirm_edit();
    }

    /// Build an `is_filled` predicate from an explicit list of filled cells.
    fn filled_set(cells: &[(u32, u32)]) -> impl Fn(u32, u32) -> bool + '_ {
        let set: HashSet<(u32, u32)> = cells.iter().copied().collect();
        move |r, c| set.contains(&(r, c))
    }

    // Use a small grid for tests so edge cases are easy to reason about.
    const R: u32 = 10;
    const C: u32 = 10;

    #[test]
    fn jump_at_edge_is_noop() {
        let f = filled_set(&[]);
        assert_eq!(compute_jump_target(0, 0, -1, 0, R, C, &f), (0, 0));
        assert_eq!(compute_jump_target(0, 0, 0, -1, R, C, &f), (0, 0));
        assert_eq!(compute_jump_target(R, C, 1, 0, R, C, &f), (R, C));
        assert_eq!(compute_jump_target(R, C, 0, 1, R, C, &f), (R, C));
    }

    #[test]
    fn jump_filled_to_filled_walks_to_end_of_run() {
        // Row 0 cols 0..=4 filled, then a gap, then col 7 filled.
        let f = filled_set(&[(0, 0), (0, 1), (0, 2), (0, 3), (0, 4), (0, 7)]);
        // From (0,0) going right, expect to stop at (0,4) — last of the run.
        assert_eq!(compute_jump_target(0, 0, 0, 1, R, C, &f), (0, 4));
        // From (0,1) going right, same answer (walks the run).
        assert_eq!(compute_jump_target(0, 1, 0, 1, R, C, &f), (0, 4));
        // From (0,4) going left, expect (0,0).
        assert_eq!(compute_jump_target(0, 4, 0, -1, R, C, &f), (0, 0));
    }

    #[test]
    fn jump_empty_skips_to_first_filled() {
        // Filled island at col 5 only.
        let f = filled_set(&[(0, 5)]);
        // From (0,0) on an empty cell, going right, lands on (0,5).
        assert_eq!(compute_jump_target(0, 0, 0, 1, R, C, &f), (0, 5));
        // From (0,9) on an empty cell, going left, lands on (0,5).
        assert_eq!(compute_jump_target(0, 9, 0, -1, R, C, &f), (0, 5));
    }

    #[test]
    fn jump_filled_with_empty_neighbour_skips_to_next_island() {
        // Two islands: col 2 and col 7.
        let f = filled_set(&[(0, 2), (0, 7)]);
        // From (0,2), neighbour (0,3) is empty → rule 3: skip to first filled
        // → (0,7).
        assert_eq!(compute_jump_target(0, 2, 0, 1, R, C, &f), (0, 7));
    }

    #[test]
    fn jump_no_content_in_direction_snaps_to_edge() {
        let f = filled_set(&[(0, 0)]);
        // Row 0 going right past the filled origin → no other filled → edge.
        assert_eq!(compute_jump_target(0, 0, 0, 1, R, C, &f), (0, C));
        // From an empty col with no content below → snap to bottom edge.
        let f = filled_set(&[]);
        assert_eq!(compute_jump_target(0, 0, 1, 0, R, C, &f), (R, 0));
    }

    // ── V2 vim motions ──────────────────────────────────────────────

    #[test]
    fn word_forward_walks_past_run_then_skips_empties() {
        // Filled at cols 0,1,2 then gap then col 6.
        let f = filled_set(&[(0, 0), (0, 1), (0, 2), (0, 6)]);
        // From (0,0) — inside a run — `w` jumps past the run to (0,6).
        assert_eq!(compute_word_forward(0, 0, C, &f), (0, 6));
        // From (0,1) inside the run — same.
        assert_eq!(compute_word_forward(0, 1, C, &f), (0, 6));
        // From (0,3) (empty) — skips empties to (0,6).
        assert_eq!(compute_word_forward(0, 3, C, &f), (0, 6));
        // From (0,6) inside trailing run — walks past to edge.
        assert_eq!(compute_word_forward(0, 6, C, &f), (0, C));
    }

    #[test]
    fn word_forward_stays_at_edge_with_no_more_content() {
        let f = filled_set(&[(0, 0)]);
        // From the only filled cell — walks past, finds no more, stops at C.
        assert_eq!(compute_word_forward(0, 0, C, &f), (0, C));
    }

    #[test]
    fn word_backward_lands_on_run_start() {
        let f = filled_set(&[(0, 0), (0, 1), (0, 2), (0, 6)]);
        // From (0,2), inside a run after its start → walk back to (0,0).
        assert_eq!(compute_word_backward(0, 2, &f), (0, 0));
        // From (0,6), inside a run at its start → walk back through empties
        // and land on start of the previous run (0,0).
        assert_eq!(compute_word_backward(0, 6, &f), (0, 0));
        // From (0,4), empty → walk back through empties to (0,2), then to start (0,0).
        assert_eq!(compute_word_backward(0, 4, &f), (0, 0));
    }

    #[test]
    fn word_backward_at_col_zero_is_noop() {
        let f = filled_set(&[]);
        assert_eq!(compute_word_backward(3, 0, &f), (3, 0));
    }

    #[test]
    fn first_filled_in_row_finds_left_edge_of_content() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 3, "hello");
        type_into(&mut app, 0, 5, "world");
        app.cursor_row = 0;
        app.cursor_col = 7;
        app.move_to_first_filled_in_row();
        assert_eq!((app.cursor_row, app.cursor_col), (0, 3));
    }

    #[test]
    fn first_filled_in_empty_row_falls_back_to_col_zero() {
        let mut app = make_test_app();
        app.cursor_row = 4;
        app.cursor_col = 5;
        app.move_to_first_filled_in_row();
        assert_eq!((app.cursor_row, app.cursor_col), (4, 0));
    }

    #[test]
    fn last_filled_in_row_finds_right_edge_of_content() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 1, "a");
        type_into(&mut app, 0, 4, "b");
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.move_to_last_filled_in_row();
        assert_eq!((app.cursor_row, app.cursor_col), (0, 4));
    }

    #[test]
    fn goto_first_row_and_last_filled_row() {
        let mut app = make_test_app();
        type_into(&mut app, 2, 0, "x");
        type_into(&mut app, 5, 0, "y");
        app.cursor_row = 3;
        app.cursor_col = 0;
        app.goto_first_row();
        assert_eq!(app.cursor_row, 0);
        app.goto_last_filled_row();
        assert_eq!(app.cursor_row, 5);
    }

    #[test]
    fn goto_last_filled_row_on_empty_column_uses_global_max() {
        // Data only in column B; cursor in column A (which is empty).
        // `G` should land on the global last data row (5), not MAX_ROW.
        let mut app = make_test_app();
        type_into(&mut app, 5, 1, "x");
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.goto_last_filled_row();
        assert_eq!(app.cursor_row, 5);
    }

    #[test]
    fn goto_last_filled_row_on_empty_sheet_lands_at_row_0() {
        let mut app = make_test_app();
        app.cursor_row = 50;
        app.goto_last_filled_row();
        assert_eq!(app.cursor_row, 0);
    }

    #[test]
    fn goto_row_clamps_to_max_row() {
        let mut app = make_test_app();
        app.goto_row(MAX_ROW + 100);
        assert_eq!(app.cursor_row, MAX_ROW);
    }

    #[test]
    fn paragraph_motions_skip_to_next_run_start() {
        let mut app = make_test_app();
        // Column 0 has runs at rows 1–3 and rows 6–8.
        type_into(&mut app, 1, 0, "a");
        type_into(&mut app, 2, 0, "b");
        type_into(&mut app, 3, 0, "c");
        type_into(&mut app, 6, 0, "d");
        type_into(&mut app, 7, 0, "e");
        type_into(&mut app, 8, 0, "f");

        app.cursor_row = 1;
        app.cursor_col = 0;
        app.paragraph_forward();
        // From start of first run, `}` walks to end of run (row 3).
        assert_eq!(app.cursor_row, 3);
        app.paragraph_forward();
        // From end of first run, `}` skips empties and walks to end of next.
        assert_eq!(app.cursor_row, 8);

        app.cursor_row = 8;
        app.paragraph_backward();
        assert_eq!(app.cursor_row, 6);
        app.paragraph_backward();
        assert_eq!(app.cursor_row, 1);
    }

    #[test]
    fn viewport_motions_use_scroll_window() {
        let mut app = make_test_app();
        app.scroll_row = 10;
        app.visible_rows = 20;

        app.move_to_viewport_top();
        assert_eq!(app.cursor_row, 10);
        app.move_to_viewport_middle();
        assert_eq!(app.cursor_row, 20);
        app.move_to_viewport_bottom();
        assert_eq!(app.cursor_row, 29);
    }

    #[test]
    fn scroll_half_page_moves_cursor_by_half_visible_rows() {
        let mut app = make_test_app();
        app.visible_rows = 20;
        app.cursor_row = 0;
        app.scroll_half_page(1);
        assert_eq!(app.cursor_row, 10);
        app.scroll_half_page(-1);
        assert_eq!(app.cursor_row, 0);
    }

    #[test]
    fn scroll_full_page_clamps_at_grid_top() {
        let mut app = make_test_app();
        app.visible_rows = 20;
        app.cursor_row = 5;
        app.scroll_full_page(-1);
        assert_eq!(app.cursor_row, 0);
    }

    #[test]
    fn z_prefix_scrolls_keep_cursor_put() {
        let mut app = make_test_app();
        app.visible_rows = 10;
        app.cursor_row = 50;

        app.scroll_cursor_to_top();
        assert_eq!(app.scroll_row, 50);
        assert_eq!(app.cursor_row, 50);

        app.scroll_cursor_to_middle();
        assert_eq!(app.scroll_row, 45);
        assert_eq!(app.cursor_row, 50);

        app.scroll_cursor_to_bottom();
        assert_eq!(app.scroll_row, 41);
        assert_eq!(app.cursor_row, 50);
    }

    #[test]
    fn consume_count_takes_and_resets() {
        let mut app = make_test_app();
        app.pending_count = Some(5);
        assert_eq!(app.consume_count(), 5);
        assert!(app.pending_count.is_none());
        // Default when empty.
        assert_eq!(app.consume_count(), 1);
    }

    #[test]
    fn clear_pending_motion_state_clears_all_prefix_flags() {
        let mut app = make_test_app();
        app.pending_count = Some(7);
        app.pending_g = true;
        app.pending_z = true;
        app.pending_f = true;
        app.clear_pending_motion_state();
        assert!(app.pending_count.is_none());
        assert!(!app.pending_g);
        assert!(!app.pending_z);
        assert!(!app.pending_f);
    }

    // ── V3 visual mode + yank register ──────────────────────────────

    #[test]
    fn enter_visual_anchors_at_cursor() {
        let mut app = make_test_app();
        app.cursor_row = 3;
        app.cursor_col = 4;
        app.enter_visual(VisualKind::Cell);
        assert_eq!(app.mode, Mode::Visual(VisualKind::Cell));
        assert_eq!(app.selection_anchor, Some((3, 4)));
        // Single-cell rect at entry.
        assert_eq!(app.selection_range(), Some((3, 4, 3, 4)));
    }

    #[test]
    fn motion_in_visual_extends_rectangle() {
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.enter_visual(VisualKind::Cell);
        app.move_cursor(2, 0);
        app.move_cursor(0, 3);
        assert_eq!(app.selection_range(), Some((0, 0, 2, 3)));
        // Anchor unchanged.
        assert_eq!(app.selection_anchor, Some((0, 0)));
    }

    #[test]
    fn motion_in_normal_resets_anchor() {
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.enter_visual(VisualKind::Cell);
        app.move_cursor(2, 0);
        app.exit_visual();
        // Now in Nav, motions clear the anchor.
        app.move_cursor(0, 1);
        assert!(app.selection_anchor.is_none());
    }

    #[test]
    fn vline_selection_pins_columns_to_full_grid() {
        let mut app = make_test_app();
        app.cursor_row = 4;
        app.cursor_col = 7;
        app.enter_visual(VisualKind::Row);
        app.move_cursor(3, 5); // moves cursor to (7, 12) but cols are pinned
        let (r1, c1, r2, c2) = app.selection_range().unwrap();
        assert_eq!((r1, c1, r2, c2), (4, 0, 7, MAX_COL));
    }

    #[test]
    fn vcolumn_selection_pins_rows_to_full_grid() {
        let mut app = make_test_app();
        app.cursor_row = 4;
        app.cursor_col = 7;
        app.enter_visual(VisualKind::Column);
        // Move cursor down + right; rows should still be pinned 0..=MAX_ROW.
        app.move_cursor(3, 5);
        let (r1, c1, r2, c2) = app.selection_range().unwrap();
        assert_eq!((r1, c1, r2, c2), (0, 7, MAX_ROW, 12));
    }

    #[test]
    fn vcolumn_status_label() {
        let mut app = make_test_app();
        app.enter_visual(VisualKind::Column);
        assert!(app.status.starts_with("V-COLUMN"));
    }

    #[test]
    fn swap_visual_corners_swaps_anchor_and_cursor() {
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.enter_visual(VisualKind::Cell);
        app.move_cursor(3, 4);
        app.swap_visual_corners();
        assert_eq!((app.cursor_row, app.cursor_col), (0, 0));
        assert_eq!(app.selection_anchor, Some((3, 4)));
        // Rect unchanged (just the anchor swapped).
        assert_eq!(app.selection_range(), Some((0, 0, 3, 4)));
    }

    #[test]
    fn exit_visual_saves_last_visual() {
        let mut app = make_test_app();
        app.cursor_row = 1;
        app.cursor_col = 1;
        app.enter_visual(VisualKind::Cell);
        app.move_cursor(2, 3);
        app.exit_visual();
        assert_eq!(app.mode, Mode::Nav);
        let last = app.last_visual.expect("last_visual saved");
        assert_eq!(last.anchor, (1, 1));
        assert_eq!(last.cursor, (3, 4));
        assert_eq!(last.kind, VisualKind::Cell);
    }

    #[test]
    fn reselect_last_visual_restores_rect() {
        let mut app = make_test_app();
        app.cursor_row = 1;
        app.cursor_col = 1;
        app.enter_visual(VisualKind::Cell);
        app.move_cursor(2, 3);
        app.exit_visual();
        // Wander.
        app.cursor_row = 8;
        app.cursor_col = 8;
        assert!(app.reselect_last_visual());
        assert_eq!(app.mode, Mode::Visual(VisualKind::Cell));
        assert_eq!((app.cursor_row, app.cursor_col), (3, 4));
        assert_eq!(app.selection_range(), Some((1, 1, 3, 4)));
    }

    #[test]
    fn reselect_last_visual_returns_false_when_none_saved() {
        let mut app = make_test_app();
        assert!(!app.reselect_last_visual());
    }

    #[test]
    fn yank_selection_round_trips_through_paste() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "alpha");
        type_into(&mut app, 0, 1, "beta");
        type_into(&mut app, 1, 0, "=A1");
        type_into(&mut app, 1, 1, "=B1");

        app.cursor_row = 0;
        app.cursor_col = 0;
        app.enter_visual(VisualKind::Cell);
        app.move_cursor(1, 1);
        let _ = app.yank_selection(false);
        app.exit_visual();

        // Paste at (3, 0) — formula refs should shift by +3 rows.
        app.cursor_row = 3;
        app.cursor_col = 0;
        let (rows, cols) = app.paste_from_register().unwrap();
        assert_eq!((rows, cols), (2, 2));
        assert_eq!(app.get_raw(3, 0), "alpha");
        assert_eq!(app.get_raw(3, 1), "beta");
        // Formula `=A1` at source (1,0), pasted at (4,0) → shift dr=3 → `=A4`.
        assert_eq!(app.get_raw(4, 0), "=A4");
        assert_eq!(app.get_raw(4, 1), "=B4");
    }

    #[test]
    fn yank_row_captures_full_a_through_max_col() {
        let mut app = make_test_app();
        type_into(&mut app, 5, 2, "x");
        type_into(&mut app, 5, 7, "y");
        app.cursor_row = 5;
        app.yank_row();
        let reg = app.yank_register.as_ref().expect("register populated");
        assert!(reg.linewise);
        assert_eq!(reg.grid.cells.len(), 1);
        assert_eq!(reg.grid.cells[0].len(), (MAX_COL + 1) as usize);
        assert_eq!(reg.grid.cells[0][2].value, "x");
        assert_eq!(reg.grid.cells[0][7].value, "y");
    }

    #[test]
    fn paste_from_empty_register_is_noop() {
        let mut app = make_test_app();
        let (rows, cols) = app.paste_from_register().unwrap();
        assert_eq!((rows, cols), (0, 0));
    }

    #[test]
    fn clear_rect_clears_every_cell_in_one_undo_step() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "a");
        type_into(&mut app, 0, 1, "b");
        type_into(&mut app, 1, 0, "c");
        type_into(&mut app, 1, 1, "d");

        let undo_groups_before = undo_group_count(&app);
        app.clear_rect(0, 0, 1, 1).unwrap();
        // Single undo group covers the whole rect.
        assert_eq!(undo_group_count(&app), undo_groups_before + 1);
        assert_eq!(app.get_raw(0, 0), "");
        assert_eq!(app.get_raw(0, 1), "");
        assert_eq!(app.get_raw(1, 0), "");
        assert_eq!(app.get_raw(1, 1), "");

        // One undo restores all four.
        assert!(app.undo());
        assert_eq!(app.get_raw(0, 0), "a");
        assert_eq!(app.get_raw(1, 1), "d");
    }

    // ── V4 operator state ───────────────────────────────────────────

    // ── V6 dot-repeat ───────────────────────────────────────────────

    // ── V5 search + marks ───────────────────────────────────────────

    // ── Variable column widths ──────────────────────────────────────

    #[test]
    fn column_width_defaults_when_unset() {
        let app = make_test_app();
        // App::new seeded A–H at DEFAULT_COL_WIDTH; columns past H fall
        // back to DEFAULT_COL_WIDTH via the unset branch.
        assert_eq!(app.column_width(0), DEFAULT_COL_WIDTH);
        assert_eq!(app.column_width(20), DEFAULT_COL_WIDTH);
    }

    #[test]
    fn set_column_width_persists_and_clamps() {
        let mut app = make_test_app();
        app.set_column_width(2, 25).unwrap();
        assert_eq!(app.column_width(2), 25);
        // Below MIN clamps up.
        app.set_column_width(2, 0).unwrap();
        assert_eq!(app.column_width(2), MIN_COL_WIDTH);
        // Above MAX clamps down.
        app.set_column_width(2, 200).unwrap();
        assert_eq!(app.column_width(2), MAX_COL_WIDTH);
    }

    #[test]
    fn autofit_column_uses_longest_displayed_value() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 1, "ab");
        type_into(&mut app, 1, 1, "longer value here");
        type_into(&mut app, 2, 1, "x");
        let w = app.autofit_column(1).unwrap();
        // Longest is "longer value here" = 17 chars; +1 padding = 18.
        assert_eq!(w, 18);
        assert_eq!(app.column_width(1), 18);
    }

    #[test]
    fn autofit_column_uses_formatted_width_for_currency() {
        // `1234.56` is 7 chars raw, but with USD/2 it renders as
        // `$1,234.56` = 9 chars. Autofit must charge for the rendered
        // width, not the raw, or the column clips the leading `$`.
        let mut app = make_test_app();
        type_into(&mut app, 0, 1, "$1,234.56");
        let w = app.autofit_column(1).unwrap();
        // Longest formatted = "$1,234.56" = 9 chars; +1 padding = 10.
        assert_eq!(w, 10);
    }

    #[test]
    fn displayed_for_applies_format() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        assert_eq!(app.displayed_for(0, 0), "$1.25");
    }

    #[test]
    fn displayed_for_passes_through_when_unformatted() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1.25");
        assert_eq!(app.displayed_for(0, 0), "1.25");
    }

    #[test]
    fn displayed_for_formats_formula_result_when_format_set() {
        // Formula cells inherit their cell's format. Set up a formula
        // cell, attach USD format directly via CellChange (T4 will
        // give us a nicer command-level path), and confirm the
        // computed result renders formatted.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "10");
        type_into(&mut app, 1, 0, "20");
        app.apply_changes_recorded(&[CellChange {
            row_idx: 2,
            col_idx: 0,
            raw_value: "=A1+A2".to_string(),
            format_json: Some(r#"{"n":{"k":"usd","d":2}}"#.to_string()),
        }])
        .unwrap();
        assert_eq!(app.displayed_for(2, 0), "$30.00");
    }

    #[test]
    fn autofit_empty_column_uses_default() {
        let mut app = make_test_app();
        let w = app.autofit_column(7).unwrap();
        assert_eq!(w, DEFAULT_COL_WIDTH);
    }

    #[test]
    fn handle_colwidth_parses_numeric_and_auto() {
        let mut app = make_test_app();
        app.handle_colwidth(0, "20");
        assert_eq!(app.column_width(0), 20);
        app.handle_colwidth(0, "auto");
        // Auto on empty column → DEFAULT_COL_WIDTH.
        assert_eq!(app.column_width(0), DEFAULT_COL_WIDTH);
        app.handle_colwidth(0, "AUTO");
        assert_eq!(app.column_width(0), DEFAULT_COL_WIDTH);
        // Out-of-range bumps status, doesn't update.
        app.handle_colwidth(0, "500");
        assert!(app.status.contains("between"));
        assert_eq!(app.column_width(0), DEFAULT_COL_WIDTH);
        app.handle_colwidth(0, "garbage");
        assert!(app.status.contains("Bad width"));
    }

    // ── V8 ex polish ─────────────────────────────────────────────────

    #[test]
    fn jump_to_target_handles_row_and_cell() {
        let mut app = make_test_app();
        assert!(app.jump_to_target("42"));
        assert_eq!(app.cursor_row, 41);
        assert!(app.jump_to_target("a1"));
        assert_eq!((app.cursor_row, app.cursor_col), (0, 0));
        assert!(app.jump_to_target("Z9"));
        assert_eq!((app.cursor_row, app.cursor_col), (8, 25));
        assert!(!app.jump_to_target("nope"));
    }

    #[test]
    fn quote_for_format_csv_escapes_quotes_and_separator() {
        assert_eq!(quote_for_format("plain", ','), "plain");
        assert_eq!(quote_for_format("a,b", ','), "\"a,b\"");
        assert_eq!(quote_for_format("she said \"hi\"", ','), "\"she said \"\"hi\"\"\"");
        assert_eq!(quote_for_format("x\ny", ','), "\"x\ny\"");
    }

    #[test]
    fn quote_for_format_tsv_replaces_tab_and_newline() {
        assert_eq!(quote_for_format("a\tb", '\t'), "a b");
        assert_eq!(quote_for_format("a\nb", '\t'), "a b");
        assert_eq!(quote_for_format("plain", '\t'), "plain");
    }

    #[test]
    fn export_sheet_writes_csv() {
        use std::io::Read;
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "alpha");
        type_into(&mut app, 0, 1, "beta");
        type_into(&mut app, 1, 0, "1,2");
        let dir = std::env::temp_dir();
        let path = dir.join(format!("vlotus-csv-{}.csv", std::process::id()));
        let path_str = path.to_str().unwrap();
        app.export_sheet(path_str).unwrap();
        let mut file = std::fs::File::open(path_str).unwrap();
        let mut buf = String::new();
        file.read_to_string(&mut buf).unwrap();
        let _ = std::fs::remove_file(path_str);
        assert!(buf.starts_with("alpha,beta\n"));
        assert!(buf.contains("\"1,2\","));
    }

    #[test]
    fn export_sheet_rejects_unknown_extension() {
        let app = make_test_app();
        assert!(app.export_sheet("/tmp/foo.xlsx").is_err());
    }

    #[test]
    fn search_finds_first_match_after_cursor_and_wraps() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "alpha");
        type_into(&mut app, 1, 1, "beta");
        type_into(&mut app, 2, 2, "alpha-bet");
        app.cursor_row = 0;
        app.cursor_col = 1;
        // Forward from (0,1): first match is (1,1)? No — looking for "alpha"
        // matches (0,0) and (2,2). From (0,1), next forward is (2,2).
        app.search = Some(SearchState {
            pattern: "alpha".into(),
            direction: SearchDir::Forward,
            case_insensitive: true,
        });
        app.search_step(SearchDir::Forward);
        assert_eq!((app.cursor_row, app.cursor_col), (2, 2));
        // Step again — should wrap to (0,0).
        app.search_step(SearchDir::Forward);
        assert_eq!((app.cursor_row, app.cursor_col), (0, 0));
    }

    #[test]
    fn search_step_backward_with_capital_n_reverses() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        type_into(&mut app, 5, 0, "x");
        app.cursor_row = 5;
        app.cursor_col = 0;
        app.search = Some(SearchState {
            pattern: "x".into(),
            direction: SearchDir::Forward,
            case_insensitive: true,
        });
        // N reverses: from (5,0), prev match is (0,0).
        app.search_step(SearchDir::Backward);
        assert_eq!((app.cursor_row, app.cursor_col), (0, 0));
    }

    #[test]
    fn search_state_matches_is_case_insensitive_when_flagged() {
        let s = SearchState {
            pattern: "Foo".into(),
            direction: SearchDir::Forward,
            case_insensitive: true,
        };
        assert!(s.matches("foobar"));
        assert!(s.matches("FOO"));
        assert!(s.matches("xFooy"));
        assert!(!s.matches("bar"));
    }

    #[test]
    fn marks_set_and_jump() {
        let mut app = make_test_app();
        app.cursor_row = 7;
        app.cursor_col = 3;
        app.set_mark('a');
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.jump_to_mark('a', false);
        assert_eq!((app.cursor_row, app.cursor_col), (7, 3));
        // 'a (row-only) lands at col 0.
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.jump_to_mark('a', true);
        assert_eq!((app.cursor_row, app.cursor_col), (7, 0));
    }

    #[test]
    fn jump_to_unset_mark_reports_status() {
        let mut app = make_test_app();
        app.jump_to_mark('z', false);
        assert!(app.status.contains("not set"));
    }

    #[test]
    fn search_current_cell_uses_displayed_value() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "needle");
        type_into(&mut app, 4, 0, "needle");
        type_into(&mut app, 8, 0, "haystack");
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.search_current_cell(SearchDir::Forward);
        assert_eq!((app.cursor_row, app.cursor_col), (4, 0));
    }

    #[test]
    fn clear_search_drops_state() {
        let mut app = make_test_app();
        app.search = Some(SearchState {
            pattern: "x".into(),
            direction: SearchDir::Forward,
            case_insensitive: true,
        });
        app.clear_search();
        assert!(app.search.is_none());
    }

    #[test]
    fn delete_cell_records_repeatable_action() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        type_into(&mut app, 1, 0, "y");
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.delete_cell();
        let action = app.last_edit.as_ref().expect("last_edit recorded");
        assert_eq!(action.kind, EditKind::Delete);
        assert_eq!((action.rect_rows, action.rect_cols), (1, 1));
        // Move and dot-repeat.
        app.cursor_row = 1;
        assert!(app.repeat_last_edit());
        assert_eq!(app.get_raw(1, 0), "");
    }

    #[test]
    fn currency_input_parses_and_formats_round_trip() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1,234.56");
        assert_eq!(app.get_raw(0, 0), "1234.56");
        assert_eq!(
            app.get_format_json_raw(0, 0).as_deref(),
            Some(r#"{"n":{"k":"usd","d":2}}"#)
        );
    }

    #[test]
    fn currency_negative_input_round_trip() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "-$3.50");
        assert_eq!(app.get_raw(0, 0), "-3.50");
        assert_eq!(
            app.get_format_json_raw(0, 0).as_deref(),
            Some(r#"{"n":{"k":"usd","d":2}}"#)
        );
    }

    #[test]
    fn currency_input_with_dollar_text_stays_string() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$abc");
        assert_eq!(app.get_raw(0, 0), "$abc");
        assert!(app.get_format_json_raw(0, 0).is_none());
    }

    #[test]
    fn formula_with_dollar_unchanged() {
        // Absolute references like `=$A$1` must NOT be peeled by the
        // currency auto-detect. The `=` guard fires first.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "10");
        type_into(&mut app, 1, 1, "=$A$1+1");
        assert_eq!(app.get_raw(1, 1), "=$A$1+1");
        assert!(app.get_format_json_raw(1, 1).is_none());
    }

    #[test]
    fn currency_dot_repeat_re_runs_auto_detect() {
        // `.` should re-run currency detection at the new cell — the
        // original buffer is what's in last_edit, so the format gets
        // freshly computed each time.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        app.cursor_row = 5;
        app.cursor_col = 5;
        assert!(app.repeat_last_edit());
        assert_eq!(app.get_raw(5, 5), "1.25");
        assert_eq!(
            app.get_format_json_raw(5, 5).as_deref(),
            Some(r#"{"n":{"k":"usd","d":2}}"#)
        );
    }

    /// Move the cursor onto a cell, then run a `:cmd`. Mirrors the
    /// keyboard sequence: navigate, then `:`. Without resetting the
    /// cursor, `type_into` leaves it advanced one row past the cell.
    fn run_ex_at(app: &mut App, row: u32, col: u32, cmd: &str) {
        app.cursor_row = row;
        app.cursor_col = col;
        run_ex(app, cmd);
    }

    #[test]
    fn fmt_usd_attaches_format_to_active_cell() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1.25");
        run_ex_at(&mut app, 0, 0, "fmt usd");
        assert_eq!(
            app.get_format_json_raw(0, 0).as_deref(),
            Some(r#"{"n":{"k":"usd","d":2}}"#)
        );
        assert_eq!(app.displayed_for(0, 0), "$1.25");
        // Raw value untouched.
        assert_eq!(app.get_raw(0, 0), "1.25");
    }

    #[test]
    fn fmt_usd_with_explicit_decimals_parses() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1");
        run_ex_at(&mut app, 0, 0, "fmt usd 4");
        assert_eq!(
            app.get_format_json_raw(0, 0).as_deref(),
            Some(r#"{"n":{"k":"usd","d":4}}"#)
        );
        assert_eq!(app.displayed_for(0, 0), "$1.0000");
    }

    #[test]
    fn fmt_usd_rejects_bad_decimals() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1");
        run_ex_at(&mut app, 0, 0, "fmt usd 99");
        assert!(app.status.contains("Bad decimals"));
        assert!(app.get_format_json_raw(0, 0).is_none());
    }

    #[test]
    fn fmt_plus_minus_bumps_decimals() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25"); // auto-detect: USD/2
        run_ex_at(&mut app, 0, 0, "fmt+");
        assert_eq!(app.displayed_for(0, 0), "$1.250");
        run_ex_at(&mut app, 0, 0, "fmt+");
        assert_eq!(app.displayed_for(0, 0), "$1.2500");
        run_ex_at(&mut app, 0, 0, "fmt-");
        assert_eq!(app.displayed_for(0, 0), "$1.250");
    }

    #[test]
    fn fmt_plus_on_unformatted_applies_usd_then_bumps() {
        // gsheets toolbar parity: pressing +decimals on a plain number
        // shifts it to USD with the bumped count.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "5");
        run_ex_at(&mut app, 0, 0, "fmt+");
        // Default 2 + 1 = 3.
        assert_eq!(
            app.get_format_json_raw(0, 0).as_deref(),
            Some(r#"{"n":{"k":"usd","d":3}}"#)
        );
        assert_eq!(app.displayed_for(0, 0), "$5.000");
    }

    #[test]
    fn fmt_clear_drops_format() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        assert!(app.get_format_json_raw(0, 0).is_some());
        run_ex_at(&mut app, 0, 0, "fmt clear");
        assert!(app.get_format(0, 0).is_none());
        assert_eq!(app.displayed_for(0, 0), "1.25");
    }

    #[test]
    fn fmt_bold_attaches_flag() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt bold");
        assert!(app.get_format(0, 0).unwrap().bold);
    }

    #[test]
    fn fmt_nobold_clears_flag() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt bold");
        run_ex_at(&mut app, 0, 0, "fmt nobold");
        assert!(!app.get_format(0, 0).unwrap().bold);
    }

    #[test]
    fn fmt_style_flags_compose() {
        // bold + italic + underline + strike all set on the same cell.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt bold");
        run_ex_at(&mut app, 0, 0, "fmt italic");
        run_ex_at(&mut app, 0, 0, "fmt underline");
        run_ex_at(&mut app, 0, 0, "fmt strike");
        let fmt = app.get_format(0, 0).unwrap();
        assert!(fmt.bold && fmt.italic && fmt.underline && fmt.strike);
    }

    #[test]
    fn fmt_style_preserves_number_format() {
        // Setting bold on a USD cell must not drop the USD format —
        // apply_format_update merges, doesn't replace.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        run_ex_at(&mut app, 0, 0, "fmt bold");
        let fmt = app.get_format(0, 0).unwrap();
        assert!(fmt.bold);
        assert_eq!(
            fmt.number,
            Some(format::NumberFormat::Usd { decimals: 2 })
        );
        assert_eq!(app.displayed_for(0, 0), "$1.25");
    }

    #[test]
    fn fmt_usd_preserves_style_flags() {
        // Reverse direction: setting USD on a bold cell keeps bold.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1.25");
        run_ex_at(&mut app, 0, 0, "fmt bold");
        run_ex_at(&mut app, 0, 0, "fmt usd");
        let fmt = app.get_format(0, 0).unwrap();
        assert!(fmt.bold);
        assert_eq!(
            fmt.number,
            Some(format::NumberFormat::Usd { decimals: 2 })
        );
    }

    #[test]
    fn fmt_fg_named_color() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt fg red");
        assert_eq!(
            app.get_format(0, 0).unwrap().fg,
            Some(format::Color::rgb(0xf3, 0x8b, 0xa8))
        );
    }

    #[test]
    fn fmt_bg_hex_color() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt bg #1e1e2e");
        assert_eq!(
            app.get_format(0, 0).unwrap().bg,
            Some(format::Color::rgb(0x1e, 0x1e, 0x2e))
        );
    }

    #[test]
    fn fmt_fg_bad_color_reports_status() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt fg notacolor");
        assert!(app.status.contains("Bad color"));
        assert!(app.get_format(0, 0).is_none());
    }

    #[test]
    fn fmt_nofg_clears_only_fg() {
        // Setting fg + bg, then nofg, leaves bg intact — :fmt nofg
        // is field-scoped, not full-clear.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt fg red");
        run_ex_at(&mut app, 0, 0, "fmt bg blue");
        run_ex_at(&mut app, 0, 0, "fmt nofg");
        let fmt = app.get_format(0, 0).unwrap();
        assert!(fmt.fg.is_none());
        assert!(fmt.bg.is_some());
    }

    #[test]
    fn fmt_color_preserves_other_axes() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        run_ex_at(&mut app, 0, 0, "fmt bold");
        run_ex_at(&mut app, 0, 0, "fmt fg red");
        let fmt = app.get_format(0, 0).unwrap();
        assert!(fmt.bold);
        assert!(fmt.fg.is_some());
        assert_eq!(
            fmt.number,
            Some(format::NumberFormat::Usd { decimals: 2 })
        );
    }

    #[test]
    fn fmt_percent_attaches_format() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "0.05");
        run_ex_at(&mut app, 0, 0, "fmt percent");
        assert_eq!(
            app.get_format(0, 0).unwrap().number,
            Some(format::NumberFormat::Percent { decimals: 0 })
        );
        assert_eq!(app.displayed_for(0, 0), "5%");
    }

    #[test]
    fn fmt_percent_with_decimals_renders_correctly() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "0.12345");
        run_ex_at(&mut app, 0, 0, "fmt percent 2");
        assert_eq!(app.displayed_for(0, 0), "12.35%"); // rounded
    }

    #[test]
    fn percent_input_auto_detects_at_edit_commit() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "4.5%");
        assert_eq!(app.get_raw(0, 0), "0.045");
        let fmt = app.get_format(0, 0).unwrap();
        assert_eq!(fmt.number, Some(format::NumberFormat::Percent { decimals: 1 }));
        assert_eq!(app.displayed_for(0, 0), "4.5%");
    }

    #[test]
    fn percent_input_with_text_stays_string() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "abc%");
        // Bare "abc%" doesn't auto-detect (parse_percent_suffix sees
        // a non-digit before %); gets stored as text.
        assert_eq!(app.get_raw(0, 0), "abc%");
        assert!(app.get_format(0, 0).is_none());
    }

    #[test]
    fn fmt_left_overrides_classify_right_align() {
        // A USD cell normally classifies as Number → right-aligned.
        // `:fmt left` should override that.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        run_ex_at(&mut app, 0, 0, "fmt left");
        assert_eq!(app.get_format(0, 0).unwrap().align, Some(format::Align::Left));
    }

    #[test]
    fn fmt_center_and_right_round_trip() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt center");
        assert_eq!(app.get_format(0, 0).unwrap().align, Some(format::Align::Center));
        run_ex_at(&mut app, 0, 0, "fmt right");
        assert_eq!(app.get_format(0, 0).unwrap().align, Some(format::Align::Right));
    }

    #[test]
    fn fmt_auto_clears_explicit_alignment() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        run_ex_at(&mut app, 0, 0, "fmt center");
        run_ex_at(&mut app, 0, 0, "fmt auto");
        assert!(app.get_format(0, 0).unwrap().align.is_none());
    }

    #[test]
    fn fmt_alignment_preserves_other_axes() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        run_ex_at(&mut app, 0, 0, "fmt bold");
        run_ex_at(&mut app, 0, 0, "fmt left");
        let fmt = app.get_format(0, 0).unwrap();
        assert!(fmt.bold);
        assert_eq!(fmt.align, Some(format::Align::Left));
        assert_eq!(fmt.number, Some(format::NumberFormat::Usd { decimals: 2 }));
    }

    #[test]
    fn fmt_clear_drops_style_flags_too() {
        // :fmt clear should wipe every axis, not just the number format.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        run_ex_at(&mut app, 0, 0, "fmt bold");
        run_ex_at(&mut app, 0, 0, "fmt italic");
        run_ex_at(&mut app, 0, 0, "fmt clear");
        assert!(app.get_format(0, 0).is_none());
    }

    #[test]
    fn fmt_applies_across_visual_selection() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1");
        type_into(&mut app, 0, 1, "2");
        type_into(&mut app, 1, 0, "3");
        type_into(&mut app, 1, 1, "4");
        // Select 2x2.
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.enter_visual(VisualKind::Cell);
        app.move_cursor(1, 1);
        run_ex(&mut app, "fmt usd");
        for (r, c) in [(0, 0), (0, 1), (1, 0), (1, 1)] {
            assert_eq!(
                app.get_format_json_raw(r, c).as_deref(),
                Some(r#"{"n":{"k":"usd","d":2}}"#),
                "cell ({r},{c})",
            );
        }
    }

    #[test]
    fn fmt_undo_restores_unformatted_state() {
        // Apply USD to a previously-unformatted cell, undo, and confirm
        // the format is actually gone (not just preserved by COALESCE).
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1.25");
        assert!(app.get_format(0, 0).is_none());
        run_ex_at(&mut app, 0, 0, "fmt usd");
        assert!(app.get_format(0, 0).is_some());
        assert!(app.undo());
        assert!(
            app.get_format(0, 0).is_none(),
            "undo must drop the format applied to a previously-unformatted cell"
        );
        assert_eq!(app.get_raw(0, 0), "1.25");
    }

    #[test]
    fn fmt_undo_restores_prior_format() {
        // Apply USD/2, then USD/4, then undo — should land back at USD/2.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "$1.25");
        run_ex_at(&mut app, 0, 0, "fmt usd 4");
        assert!(app.displayed_for(0, 0).starts_with("$1.2500"));
        assert!(app.undo());
        assert_eq!(app.displayed_for(0, 0), "$1.25");
    }

    #[test]
    fn fmt_skips_empty_cells_in_selection() {
        // Format applied across a partial selection where some cells
        // are blank: only the non-empty cells get formatted (an empty
        // raw row gets DELETEd by set_cells, so attaching a format to
        // it would silently disappear).
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1");
        // (1, 0) and (0, 1) intentionally empty.
        type_into(&mut app, 1, 1, "2");
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.enter_visual(VisualKind::Cell);
        app.move_cursor(1, 1);
        run_ex(&mut app, "fmt usd");
        assert!(app.get_format(0, 0).is_some());
        assert!(app.get_format(1, 1).is_some());
        assert!(app.get_format(0, 1).is_none());
        assert!(app.get_format(1, 0).is_none());
    }

    #[test]
    fn confirm_edit_records_insert_for_dot_repeat() {
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('h');
        app.edit_insert('i');
        app.confirm_edit();
        let action = app.last_edit.as_ref().expect("Insert recorded");
        assert_eq!(action.kind, EditKind::Insert);
        assert_eq!(action.text.as_deref(), Some("hi"));
        // Repeat at a different cell.
        app.cursor_row = 5;
        app.cursor_col = 5;
        assert!(app.repeat_last_edit());
        assert_eq!(app.get_raw(5, 5), "hi");
    }

    #[test]
    fn cancel_edit_downgrades_change_to_delete() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        // Simulate apply_operator(Change) having recorded a half-built action.
        app.last_edit = Some(EditAction {
            kind: EditKind::Change,
            anchor_dr: 0,
            anchor_dc: 0,
            rect_rows: 1,
            rect_cols: 1,
            text: None,
        });
        app.start_edit_blank();
        app.cancel_edit();
        let action = app.last_edit.as_ref().expect("preserved");
        assert_eq!(action.kind, EditKind::Delete);
    }

    #[test]
    fn repeat_last_edit_returns_false_when_none() {
        let mut app = make_test_app();
        assert!(!app.repeat_last_edit());
    }

    #[test]
    fn pending_operator_stores_count_and_clears_via_motion_state() {
        let mut app = make_test_app();
        app.pending_operator = Some(Operator::Delete);
        app.pending_op_count = Some(5);
        app.clear_pending_motion_state();
        assert!(app.pending_operator.is_none());
        assert!(app.pending_op_count.is_none());
    }

    #[test]
    fn build_grid_over_includes_formulas() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "alpha");
        type_into(&mut app, 0, 1, "=A1*2");
        let grid = app.build_grid_over(0, 0, 0, 1);
        assert_eq!(grid.cells[0][0].value, "alpha");
        assert_eq!(grid.cells[0][0].formula, None);
        assert_eq!(grid.cells[0][1].formula.as_deref(), Some("=A1*2"));
        assert_eq!(grid.source_anchor, Some((0, 0)));
    }

    #[test]
    fn clipboard_mark_contains_inclusive() {
        let m = ClipboardMark {
            r1: 2,
            c1: 3,
            r2: 5,
            c2: 6,
            mode: ClipMarkMode::Copy,
        };
        assert!(m.contains(2, 3));
        assert!(m.contains(5, 6));
        assert!(m.contains(3, 4));
        assert!(!m.contains(1, 3));
        assert!(!m.contains(2, 2));
        assert!(!m.contains(6, 4));
    }

    #[test]
    fn clipboard_mark_perimeter_excludes_strict_interior() {
        let m = ClipboardMark {
            r1: 2,
            c1: 3,
            r2: 5,
            c2: 6,
            mode: ClipMarkMode::Cut,
        };
        // corners
        assert!(m.on_perimeter(2, 3));
        assert!(m.on_perimeter(5, 6));
        // top / bottom / left / right edges
        assert!(m.on_perimeter(2, 5));
        assert!(m.on_perimeter(5, 4));
        assert!(m.on_perimeter(3, 3));
        assert!(m.on_perimeter(4, 6));
        // strict interior
        assert!(!m.on_perimeter(3, 4));
        assert!(!m.on_perimeter(4, 5));
        // outside
        assert!(!m.on_perimeter(0, 0));
    }

    #[test]
    fn clipboard_mark_single_cell_is_all_perimeter() {
        // A 1×1 mark — every cell in it is on the perimeter (there's no
        // interior to be strictly-inside).
        let m = ClipboardMark {
            r1: 7,
            c1: 7,
            r2: 7,
            c2: 7,
            mode: ClipMarkMode::Copy,
        };
        assert!(m.on_perimeter(7, 7));
        assert!(!m.on_perimeter(7, 8));
    }

    #[test]
    fn command_quit_aliases() {
        for cmd in ["q", "quit", "wq", "x"] {
            assert_eq!(
                execute_command(cmd),
                Ok(CommandOutcome::Quit { force: false }),
                "{cmd}"
            );
        }
    }

    #[test]
    fn command_quit_bang_aliases_force() {
        for cmd in ["q!", "quit!", "wq!", "x!"] {
            assert_eq!(
                execute_command(cmd),
                Ok(CommandOutcome::Quit { force: true }),
                "{cmd}"
            );
        }
    }

    #[test]
    fn command_write_is_noop_continue() {
        assert_eq!(execute_command("w"), Ok(CommandOutcome::Continue));
        assert_eq!(execute_command("write"), Ok(CommandOutcome::Continue));
    }

    #[test]
    fn command_empty_is_noop() {
        assert_eq!(execute_command(""), Ok(CommandOutcome::Continue));
    }

    #[test]
    fn command_unknown_is_error() {
        let err = execute_command("nope").unwrap_err();
        assert!(err.contains("nope"), "error mentions the unknown command: {err}");
    }

    #[test]
    fn nav_bang_opens_shell_prompt() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        app.start_shell();
        assert_eq!(app.mode, Mode::Shell);
        assert_eq!(app.edit_buf, "");
        assert_eq!(app.edit_cursor, 0);
    }

    #[test]
    fn shell_esc_cancels_back_to_nav() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        app.start_shell();
        app.edit_buf = "echo hi".into();
        app.edit_cursor = 7;
        app.cancel_shell();
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.edit_buf, "");
        assert_eq!(app.edit_cursor, 0);
    }

    #[test]
    fn shell_enter_with_empty_buffer_is_silent_noop() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        app.start_shell();
        let out = app.run_shell();
        assert_eq!(out, CommandOutcome::Continue);
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.status, "");
    }

    /// Build a 3-row payload (header + 2 data rows, 2 cols).
    fn make_csv_payload() -> PastedGrid {
        let cell = |s: &str| PastedCell {
            value: s.to_string(),
            formula: None,
        };
        PastedGrid {
            source_anchor: None,
            cells: vec![
                vec![cell("name"), cell("age")],
                vec![cell("alice"), cell("30")],
                vec![cell("bob"), cell("25")],
            ],
        }
    }

    #[test]
    fn shell_paste_lands_headers_at_cursor_and_data_below() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        app.cursor_row = 1;
        app.cursor_col = 1;
        let grid = make_csv_payload();
        app.insert_shell_payload(&grid);
        assert_eq!(app.get_display(1, 1), "name");
        assert_eq!(app.get_display(1, 2), "age");
        assert_eq!(app.get_display(2, 1), "alice");
        assert_eq!(app.get_display(2, 2), "30");
        assert_eq!(app.get_display(3, 1), "bob");
        assert_eq!(app.get_display(3, 2), "25");
        assert!(
            app.status.contains("3×2"),
            "status reports rows×cols: {}",
            app.status
        );
    }

    #[test]
    fn shell_paste_is_one_undo_group() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        let grid = make_csv_payload();
        app.insert_shell_payload(&grid);
        // One undo step should clear every pasted cell.
        app.undo();
        for r in 0..3 {
            for c in 0..2 {
                assert_eq!(
                    app.get_display(r, c),
                    "",
                    "cell ({r},{c}) cleared by single undo"
                );
            }
        }
    }

    #[test]
    fn shell_paste_clamps_at_grid_bottom() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        // Drop the cursor one row from the bottom so a 3-row paste
        // overflows by 2 rows.
        app.cursor_row = MAX_ROW;
        app.cursor_col = 0;
        let grid = make_csv_payload();
        app.insert_shell_payload(&grid);
        // Header row lands; the two data rows would be at MAX_ROW+1
        // and MAX_ROW+2 which are outside the grid and get silently
        // dropped (existing apply_pasted_grid contract).
        assert_eq!(app.get_display(MAX_ROW, 0), "name");
        assert_eq!(app.get_display(MAX_ROW, 1), "age");
        // Out-of-bounds rows didn't write anything past MAX_ROW.
        // The status reports the requested paste size, not the clamped
        // count — same contract as the clipboard paste path.
        assert!(
            app.status.contains("3×2"),
            "status reports the requested paste size: {}",
            app.status
        );
    }

    #[test]
    fn shell_paste_with_empty_grid_status_message() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        let grid = PastedGrid {
            source_anchor: None,
            cells: vec![],
        };
        app.insert_shell_payload(&grid);
        assert!(
            app.status.contains("empty output"),
            "status reports empty: {}",
            app.status
        );
    }

    #[cfg(unix)]
    #[test]
    fn run_shell_pipeline_with_csv_input() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        app.start_shell();
        app.edit_buf = "printf 'a,b\\nx,y\\n'".into();
        app.edit_cursor = app.edit_buf.len();
        let out = app.run_shell();
        assert_eq!(out, CommandOutcome::Continue);
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.get_display(0, 0), "a");
        assert_eq!(app.get_display(0, 1), "b");
        assert_eq!(app.get_display(1, 0), "x");
        assert_eq!(app.get_display(1, 1), "y");
    }

    #[cfg(unix)]
    #[test]
    fn run_shell_surfaces_nonzero_exit_in_status() {
        let mut app = App::new(Store::open_in_memory().unwrap(), ":memory:");
        app.start_shell();
        app.edit_buf = "exit 3".into();
        app.edit_cursor = app.edit_buf.len();
        app.run_shell();
        assert!(
            app.status.contains("!:") && app.status.contains("3"),
            "status surfaces error: {}",
            app.status
        );
        // No paste happened.
        assert_eq!(app.get_display(0, 0), "");
    }

    /// Buffer `cmd` and run it, returning the outcome. Mirrors what
    /// `handle_command_key` does when the user hits Enter on `:cmd`.
    fn run_ex_outcome(app: &mut App, cmd: &str) -> CommandOutcome {
        app.start_command();
        app.edit_buf = cmd.to_string();
        app.edit_cursor = cmd.len();
        app.run_command()
    }

    fn make_file_backed_app() -> App {
        App::new(Store::open_in_memory().unwrap(), "/tmp/some-file.db")
    }

    fn make_tutor_app() -> App {
        App::new(Store::open_in_memory().unwrap(), "tutor")
    }

    #[test]
    fn quit_in_memory_untouched_session_succeeds() {
        let mut app = make_test_app();
        assert!(!app.touched, "fresh App is untouched");
        assert_eq!(
            run_ex_outcome(&mut app, "q"),
            CommandOutcome::Quit { force: false }
        );
    }

    #[test]
    fn quit_in_memory_touched_session_is_refused() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1");
        assert!(app.touched, "cell write marks touched");
        assert_eq!(run_ex_outcome(&mut app, "q"), CommandOutcome::Continue);
        assert!(
            app.status.contains("In-memory session"),
            "status warns about in-memory: {}",
            app.status
        );
    }

    #[test]
    fn quit_bang_overrides_in_memory_warning() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1");
        assert_eq!(
            run_ex_outcome(&mut app, "q!"),
            CommandOutcome::Quit { force: true }
        );
    }

    #[test]
    fn quit_in_tutor_session_skips_in_memory_warning() {
        let mut app = make_tutor_app();
        type_into(&mut app, 0, 0, "1");
        // tutor has db_label "tutor", not ":memory:" — so the in-memory
        // guard does not fire even though the connection is in-memory.
        assert_eq!(
            run_ex_outcome(&mut app, "q"),
            CommandOutcome::Quit { force: false }
        );
    }

    #[test]
    fn patch_apply_command_lands_changes_into_dirty_buffer() {
        // Build a patch on store A; apply it via :patch apply on a
        // fresh App and verify the cells materialised AND the dirty
        // buffer is set (so :q would refuse).
        use std::sync::atomic::{AtomicUsize, Ordering};
        static COUNTER: AtomicUsize = AtomicUsize::new(0);
        let nanos = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let n = COUNTER.fetch_add(1, Ordering::Relaxed);
        let dir = std::env::temp_dir().join(format!(
            "vlotus-app-patch-{}-{nanos}-{n}",
            std::process::id()
        ));
        std::fs::create_dir_all(&dir).unwrap();
        let patch_path = dir.join("p.lpatch");

        // Generate a patch via a separate Store.
        {
            let mut s = crate::store::Store::open_in_memory().unwrap();
            s.create_sheet("Sheet1").unwrap();
            s.commit().unwrap();
            s.patch_open(patch_path.clone()).unwrap();
            s.apply(
                "Sheet1",
                &[CellChange {
                    row_idx: 0,
                    col_idx: 0,
                    raw_value: "from-patch".into(),
                    format_json: None,
                }],
            )
            .unwrap();
            s.patch_close().unwrap();
        }

        let mut app = make_test_app();
        let cmd = format!("patch apply {}", patch_path.display());
        let outcome = run_ex_outcome(&mut app, &cmd);
        assert_eq!(outcome, CommandOutcome::Continue);
        assert!(
            app.status.contains("applied"),
            "status reports apply: {}",
            app.status
        );
        // Cell from the patch should be present in App's view.
        assert_eq!(app.get_raw(0, 0), "from-patch");
        // Dirty buffer is on — user reviews then :w / :q!.
        assert!(app.has_unsaved_changes());
    }

    #[test]
    fn quit_file_backed_dirty_buffer_is_refused_then_allowed_after_w() {
        // T6: edits accumulate in the dirty-buffer txn. `:q` is
        // refused until `:w` commits. `:q!` would discard.
        let mut app = make_file_backed_app();
        type_into(&mut app, 0, 0, "1");
        assert!(app.has_unsaved_changes());
        assert_eq!(run_ex_outcome(&mut app, "q"), CommandOutcome::Continue);
        assert!(
            app.status.contains("No write since last change"),
            "status warns: {}",
            app.status
        );
        // After `:w`, the dirty buffer is committed and `:q` succeeds.
        assert_eq!(run_ex_outcome(&mut app, "w"), CommandOutcome::Continue);
        assert!(!app.has_unsaved_changes());
        assert_eq!(
            run_ex_outcome(&mut app, "q"),
            CommandOutcome::Quit { force: false }
        );
    }

    #[test]
    fn quit_with_dirty_buffer_is_refused() {
        // Forward-looking for CSV mode (T2 `1vdapj3i`) which flips
        // `dirty = true` on edits and back to false on `:w`.
        let mut app = make_file_backed_app();
        app.dirty = true;
        assert_eq!(run_ex_outcome(&mut app, "q"), CommandOutcome::Continue);
        assert!(
            app.status.contains("No write since last change"),
            "status warns about unsaved changes: {}",
            app.status
        );
    }

    #[test]
    fn quit_bang_overrides_dirty_buffer_warning() {
        let mut app = make_file_backed_app();
        app.dirty = true;
        assert_eq!(
            run_ex_outcome(&mut app, "q!"),
            CommandOutcome::Quit { force: true }
        );
    }

    #[test]
    fn wq_bang_forces_quit_in_memory() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1");
        assert_eq!(
            run_ex_outcome(&mut app, "wq!"),
            CommandOutcome::Quit { force: true }
        );
    }

    #[test]
    fn wq_in_memory_touched_is_refused() {
        // `:w` is a no-op in sqlite mode (writes immediate); for
        // in-memory, it doesn't make changes persist either, so the
        // warning still fires.
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "1");
        assert_eq!(run_ex_outcome(&mut app, "wq"), CommandOutcome::Continue);
        assert!(app.status.contains("In-memory session"));
    }

    #[test]
    fn parse_pasted_grid_extracts_anchor_and_per_cell_formula() {
        let html = "<table data-sheets-source-anchor=\"B2\">\n\
            <tbody>\n\
            <tr><td>1</td><td data-sheets-formula=\"=A1+1\">2</td></tr>\n\
            <tr><td data-sheets-formula=\"=$A$1*2\">14</td><td>foo</td></tr>\n\
            </tbody></table>";
        let grid = parse_pasted_grid(html).expect("parses");
        assert_eq!(grid.source_anchor, Some((1, 1))); // B2 → row 1, col 1 (0-based)
        assert_eq!(grid.cells.len(), 2);
        assert_eq!(grid.cells[0][0].value, "1");
        assert_eq!(grid.cells[0][0].formula, None);
        assert_eq!(grid.cells[0][1].value, "2");
        assert_eq!(grid.cells[0][1].formula.as_deref(), Some("=A1+1"));
        assert_eq!(grid.cells[1][0].formula.as_deref(), Some("=$A$1*2"));
        assert_eq!(grid.cells[1][1].value, "foo");
        assert_eq!(grid.cells[1][1].formula, None);
    }

    #[test]
    fn parse_pasted_grid_external_clipboard_has_no_anchor() {
        // A Sheets-/Excel-shaped HTML payload with no app-private markers.
        let html = "<table>\n<tbody>\n<tr><td>1</td><td>2</td></tr></tbody></table>";
        let grid = parse_pasted_grid(html).unwrap();
        assert_eq!(grid.source_anchor, None);
        assert!(grid.cells[0].iter().all(|c| c.formula.is_none()));
    }

    #[test]
    fn parse_pasted_grid_decodes_attribute_entities() {
        // Quotes inside formulas (`"hi"`) are escaped as `&quot;` in the HTML.
        let html = "<table data-sheets-source-anchor=\"A1\">\n\
            <tr><td data-sheets-formula=\"=CONCAT(&quot;a&quot;,&quot;b&quot;)\">ab</td></tr>\n\
            </table>";
        let grid = parse_pasted_grid(html).unwrap();
        assert_eq!(
            grid.cells[0][0].formula.as_deref(),
            Some("=CONCAT(\"a\",\"b\")")
        );
    }

    #[test]
    fn shift_pasted_raw_no_formula_returns_value() {
        let cell = PastedCell {
            value: "hello".into(),
            formula: None,
        };
        assert_eq!(shift_pasted_raw(&cell, Some((0, 0)), (5, 5)), "hello");
    }

    #[test]
    fn shift_pasted_raw_zero_delta_returns_formula_unchanged() {
        let cell = PastedCell {
            value: "3".into(),
            formula: Some("=A1+B2".into()),
        };
        assert_eq!(shift_pasted_raw(&cell, Some((1, 1)), (1, 1)), "=A1+B2");
    }

    #[test]
    fn shift_pasted_raw_shifts_relative_refs() {
        // A cell at B2 holds =A2*2. Pasting it at C5 (delta = +3 rows, +1 col)
        // should shift the formula to =B5*2.
        let cell = PastedCell {
            value: "0".into(),
            formula: Some("=A2*2".into()),
        };
        // source_pos = B2 = (1, 1); target = C5 = (4, 2).
        let shifted = shift_pasted_raw(&cell, Some((1, 1)), (4, 2));
        assert_eq!(shifted, "=B5*2");
    }

    #[test]
    fn shift_pasted_raw_pins_absolute_refs() {
        let cell = PastedCell {
            value: "0".into(),
            formula: Some("=$A$1*2".into()),
        };
        // Same delta as above; absolute ref must NOT shift.
        let shifted = shift_pasted_raw(&cell, Some((1, 1)), (4, 2));
        assert_eq!(shifted, "=$A$1*2");
    }

    #[test]
    fn shift_pasted_raw_off_grid_becomes_ref_error() {
        // =A1 pasted somewhere that would shift A1 above row 0 → #REF!.
        let cell = PastedCell {
            value: "0".into(),
            formula: Some("=A1".into()),
        };
        // source_pos = (5, 5); target = (0, 0); delta = (-5, -5). A1 would
        // shift to (-4, -4) → off-grid → #REF!.
        let shifted = shift_pasted_raw(&cell, Some((5, 5)), (0, 0));
        assert!(shifted.contains("#REF!"), "expected #REF! in {shifted}");
    }

    #[test]
    fn undo_round_trip_single_cell() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "5");
        assert_eq!(app.get_raw(0, 0), "5");

        assert!(app.undo());
        assert_eq!(app.get_raw(0, 0), "");

        assert!(app.redo());
        assert_eq!(app.get_raw(0, 0), "5");
    }

    #[test]
    fn undo_chain_of_edits_replays_in_reverse() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "5");
        type_into(&mut app, 0, 0, "10"); // overwrite

        assert!(app.undo()); // back to 5
        assert_eq!(app.get_raw(0, 0), "5");
        assert!(app.undo()); // back to empty
        assert_eq!(app.get_raw(0, 0), "");
        // Stack now empty.
        assert!(!app.undo());

        // Redo chain replays forward in order.
        assert!(app.redo());
        assert_eq!(app.get_raw(0, 0), "5");
        assert!(app.redo());
        assert_eq!(app.get_raw(0, 0), "10");
        assert!(!app.redo());
    }

    #[test]
    fn new_action_clears_redo_branch() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "5");
        type_into(&mut app, 0, 0, "10");
        assert!(app.undo()); // back to 5; redo stack now has one entry.
        assert_eq!(app.redo_stack.len(), 1);

        type_into(&mut app, 0, 1, "99"); // a fresh, unrelated edit
        // [sheet.undo.redo] new mutation must clear the redo branch.
        assert!(app.redo_stack.is_empty());
        assert!(!app.redo());
    }

    #[test]
    fn delete_cell_is_undoable() {
        let mut app = make_test_app();
        type_into(&mut app, 2, 3, "hello");
        app.cursor_row = 2;
        app.cursor_col = 3;
        app.delete_cell();
        assert_eq!(app.get_raw(2, 3), "");
        assert!(app.undo());
        assert_eq!(app.get_raw(2, 3), "hello");
    }

    #[test]
    fn paste_with_cut_clear_is_one_undo_step() {
        // Cut B2 (raw "9"), paste at C5; undo should restore B2 AND clear C5.
        let mut app = make_test_app();
        type_into(&mut app, 1, 1, "9"); // B2 = 9

        let grid = PastedGrid {
            source_anchor: Some((1, 1)),
            cells: vec![vec![PastedCell {
                value: "9".into(),
                formula: None,
            }]],
        };
        app.cursor_row = 4; // row 5 (0-based 4)
        app.cursor_col = 2; // col C
        app.apply_pasted_grid(&grid, Some((1, 1, 1, 1))).unwrap();

        assert_eq!(app.get_raw(4, 2), "9");
        assert_eq!(app.get_raw(1, 1), "");

        // Single undo restores both sides.
        assert!(app.undo());
        assert_eq!(app.get_raw(4, 2), "");
        assert_eq!(app.get_raw(1, 1), "9");
    }

    #[test]
    fn undo_log_capped_at_max_depth() {
        // The on-disk undo log uses `crate::store::undo::UNDO_DEPTH`
        // (currently 50, matching the prior in-memory cap). Push past
        // that and verify older groups get pruned.
        let mut app = make_test_app();
        let cap = crate::store::undo::UNDO_DEPTH;
        for i in 0..(cap + 5) {
            type_into(&mut app, 0, 0, &format!("{i}"));
        }
        assert_eq!(undo_group_count(&app), cap);
    }

    /// Count distinct undo groups currently on disk. Used by tests to
    /// observe `record_undo_group` / `pop_undo` / pruning behaviour
    /// without poking at SQL.
    fn undo_group_count(app: &App) -> usize {
        let count: i64 = app
            .store
            .conn()
            .query_row(
                "SELECT COUNT(DISTINCT group_id) FROM undo_entry",
                [],
                |r| r.get(0),
            )
            .unwrap();
        count as usize
    }

    #[test]
    fn autocomplete_appears_for_function_prefix() {
        let mut app = make_test_app();
        app.start_edit();
        app.edit_insert('=');
        app.edit_insert('S');
        app.edit_insert('U');
        let ac = app.autocomplete.as_ref().expect("autocomplete is open");
        assert!(
            ac.list.items.iter().any(|it| it.label == "SUM"),
            "items: {:?}",
            ac.list.items.iter().map(|it| &it.label).collect::<Vec<_>>()
        );
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn autocomplete_surfaces_datetime_extension_functions() {
        // Regression: registered custom functions (SPAN_TO_SECONDS,
        // DATE, NOW, …) used to be invisible to the popup because the
        // autocomplete path called the registry-blind `complete()`. With
        // the registry-aware variant they should appear alongside builtins.
        let mut app = make_test_app();
        app.start_edit();
        app.edit_insert('=');
        app.edit_insert('S');
        app.edit_insert('P');
        app.edit_insert('A');
        app.edit_insert('N');
        let ac = app.autocomplete.as_ref().expect("autocomplete is open");
        let labels: Vec<&str> = ac.list.items.iter().map(|i| i.label.as_str()).collect();
        assert!(
            labels.contains(&"SPAN_TO_SECONDS"),
            "expected SPAN_TO_SECONDS in popup, got {labels:?}"
        );
        assert!(
            labels.contains(&"SPAN_TO_DAYS"),
            "expected SPAN_TO_DAYS in popup, got {labels:?}"
        );
    }

    #[test]
    fn autocomplete_surfaces_hyperlink_function() {
        // Hyperlink is the always-on vlotus-local custom function. Even
        // without the datetime feature, the popup should see it via the
        // registry-aware completion path.
        let mut app = make_test_app();
        app.start_edit();
        app.edit_insert('=');
        app.edit_insert('H');
        app.edit_insert('Y');
        let ac = app.autocomplete.as_ref().expect("autocomplete is open");
        let labels: Vec<&str> = ac.list.items.iter().map(|i| i.label.as_str()).collect();
        assert!(
            labels.contains(&"HYPERLINK"),
            "expected HYPERLINK in popup, got {labels:?}"
        );
    }

    #[test]
    fn autocomplete_suppressed_when_no_prefix_typed() {
        // Buffer is just `=` — no function-name prefix yet, so the popup
        // must stay closed. Otherwise it would intercept Up/Down before
        // they reach pointing mode (cell above/below as ref).
        let mut app = make_test_app();
        app.start_edit();
        app.edit_insert('=');
        assert!(app.autocomplete.is_none());

        // `=SUM(` — caret right after `(`, again no prefix, popup stays closed
        // so Up/Down can insert a reference for the first argument.
        app.edit_insert('S');
        app.edit_insert('U');
        app.edit_insert('M');
        app.edit_insert('(');
        assert!(app.autocomplete.is_none());

        // Typing one matching character flips the popup back on.
        app.edit_insert('A');
        assert!(app.autocomplete.is_some());
    }

    #[test]
    fn autocomplete_dismissed_for_non_formula_input() {
        let mut app = make_test_app();
        app.start_edit_blank();
        app.edit_insert('h'); // raw value, no leading '='
        app.edit_insert('e');
        app.edit_insert('l');
        assert!(app.autocomplete.is_none());
    }

    #[test]
    fn autocomplete_select_next_wraps_at_end() {
        let mut app = make_test_app();
        app.start_edit();
        app.edit_insert('=');
        app.edit_insert('S');
        let n = app.autocomplete.as_ref().unwrap().list.items.len();
        assert!(n > 1, "need >1 SU* items for this test");
        for _ in 0..n {
            app.autocomplete_select_next();
        }
        // Wrapped back to 0.
        assert_eq!(app.autocomplete.as_ref().unwrap().selected, 0);
    }

    #[test]
    fn autocomplete_accept_inserts_function_with_open_paren() {
        let mut app = make_test_app();
        app.start_edit();
        app.edit_insert('=');
        app.edit_insert('S');
        app.edit_insert('U');
        // Force-select SUM (it should be the only or first SU* item).
        let idx = app
            .autocomplete
            .as_ref()
            .unwrap()
            .list
            .items
            .iter()
            .position(|it| it.label == "SUM")
            .expect("SUM in items");
        app.autocomplete.as_mut().unwrap().selected = idx;
        assert!(app.autocomplete_accept());
        assert_eq!(app.edit_buf, "=SUM(");
        assert_eq!(app.edit_cursor, "=SUM(".len());
    }

    #[test]
    fn autocomplete_user_selected_starts_false_and_flips_on_navigation() {
        let mut app = make_test_app();
        app.start_edit();
        for c in "=S".chars() {
            app.edit_insert(c);
        }
        // Auto-shown popup: selection is just the default top item, not a
        // deliberate user choice — Enter should still commit the cell.
        let ac = app.autocomplete.as_ref().expect("popup should be open");
        assert!(!ac.user_selected);
        // Navigating the list flips the flag.
        app.autocomplete_select_next();
        assert!(app.autocomplete.as_ref().unwrap().user_selected);
        app.autocomplete_select_prev();
        assert!(app.autocomplete.as_ref().unwrap().user_selected);
    }

    #[test]
    fn autocomplete_typing_more_resets_selection_when_not_user_picked() {
        let mut app = make_test_app();
        app.start_edit();
        for c in "=S".chars() {
            app.edit_insert(c);
        }
        // Default selection at index 0 since the user hasn't navigated.
        assert_eq!(app.autocomplete.as_ref().unwrap().selected, 0);
        app.edit_insert('U');
        // After narrowing the list selection should still be the top
        // item, since the user never made a deliberate pick.
        let ac = app.autocomplete.as_ref().unwrap();
        assert_eq!(ac.selected, 0);
        assert!(!ac.user_selected);
    }

    #[test]
    fn signature_help_appears_inside_function_call() {
        let mut app = make_test_app();
        app.start_edit();
        for c in "=SUM(".chars() {
            app.edit_insert(c);
        }
        let sig = app.signature.as_ref().expect("inside SUM() should give sig");
        assert_eq!(sig.function.name, "SUM");
        assert_eq!(sig.active_param, 0);
    }

    #[test]
    fn signature_active_param_advances_past_commas() {
        let mut app = make_test_app();
        app.start_edit();
        for c in "=ROUND(1,".chars() {
            app.edit_insert(c);
        }
        let sig = app.signature.as_ref().unwrap();
        assert_eq!(sig.function.name, "ROUND");
        assert_eq!(sig.active_param, 1);
    }

    #[test]
    fn cancel_edit_clears_autocomplete_and_signature() {
        let mut app = make_test_app();
        app.start_edit();
        for c in "=SU".chars() {
            app.edit_insert(c);
        }
        assert!(app.autocomplete.is_some());
        app.cancel_edit();
        assert!(app.autocomplete.is_none());
        assert!(app.signature.is_none());
    }

    #[test]
    fn insertable_after_equals_only_when_buffer_has_no_other_content() {
        assert!(is_insertable_at("=", 1));
        assert!(is_insertable_at("=  ", 3));
        // Caret before '=' → never insertable.
        assert!(!is_insertable_at("=A1", 0));
        // "hello" doesn't start with '='.
        assert!(!is_insertable_at("hello", 1));
    }

    #[test]
    fn insertable_after_operator_or_paren_or_comma() {
        assert!(is_insertable_at("=1+", 3));
        assert!(is_insertable_at("=SUM(", 5));
        assert!(is_insertable_at("=SUM(1,", 7));
        assert!(is_insertable_at("=2*  ", 5));
    }

    #[test]
    fn not_insertable_when_next_meaningful_char_is_a_value() {
        // Caret between ',' and ' 2)' — would produce =SUM(1,A2 2) which
        // is invalid. Spec calls this out explicitly.
        assert!(!is_insertable_at("=SUM(1, 2)", 7));
    }

    #[test]
    fn not_insertable_in_middle_of_identifier() {
        // Caret between 'A' and '1' — a partial cell ref. Should NOT
        // start pointing.
        assert!(!is_insertable_at("=A1", 2));
        // Caret in middle of a number.
        assert!(!is_insertable_at("=42+1", 2));
    }

    #[test]
    fn pointing_arrow_inserts_ref_at_insertable_caret() {
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit();
        app.edit_insert('=');
        // Caret is now at position 1, right after '='. Right-arrow should
        // start pointing and insert a ref to (0, 1) → "B1".
        assert!(app.try_pointing_arrow(0, 1, false));
        assert_eq!(app.edit_buf, "=B1");
        let p = app.pointing.expect("pointing active");
        assert_eq!(p.target_col, 1);
        assert_eq!(p.target_row, 0);
    }

    #[test]
    fn pointing_arrow_moves_existing_ref() {
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit();
        app.edit_insert('=');
        // Start pointing on (0, 1) → B1.
        app.try_pointing_arrow(0, 1, false);
        // Right again → C1, replacing the inserted token.
        app.try_pointing_arrow(0, 1, false);
        assert_eq!(app.edit_buf, "=C1");
        let p = app.pointing.unwrap();
        assert_eq!(p.target_col, 2);
    }

    #[test]
    fn pointing_arrow_clamps_at_grid_edge() {
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit();
        app.edit_insert('=');
        // Left from origin should clamp to col 0 → A1.
        assert!(app.try_pointing_arrow(0, -1, false));
        assert_eq!(app.edit_buf, "=A1");
    }

    #[test]
    fn non_arrow_key_exits_pointing() {
        let mut app = make_test_app();
        app.start_edit();
        app.edit_insert('=');
        app.try_pointing_arrow(0, 1, false);
        assert!(app.pointing.is_some());
        // Simulate the run-loop's "non-arrow key → exit pointing" hook.
        app.exit_pointing();
        app.edit_insert('+');
        assert!(app.pointing.is_none());
        assert_eq!(app.edit_buf, "=B1+");
    }

    #[test]
    fn pointing_does_not_start_in_middle_of_ref() {
        let mut app = make_test_app();
        app.start_edit();
        for c in "=A1".chars() {
            app.edit_insert(c);
        }
        // Caret is at end of "A1" — right after the digit, in the middle
        // of an identifier from is_insertable_at's perspective.
        assert!(!app.try_pointing_arrow(0, 1, false));
        assert!(app.pointing.is_none());
    }

    #[test]
    fn shift_arrow_with_no_pointing_starts_pointing_like_plain_arrow() {
        // Shift on entry has no anchor to extend from — first ref is
        // single-cell with anchor==target, ready for follow-up extends.
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit();
        app.edit_insert('=');
        assert!(app.try_pointing_arrow(0, 1, true));
        assert_eq!(app.edit_buf, "=B1");
        let p = app.pointing.expect("pointing active");
        assert_eq!((p.anchor_row, p.anchor_col), (0, 1));
        assert_eq!((p.target_row, p.target_col), (0, 1));
    }

    #[test]
    fn shift_arrow_during_pointing_extends_range() {
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit();
        app.edit_insert('=');
        // Enter pointing on B1 via plain arrow.
        app.try_pointing_arrow(0, 1, false);
        assert_eq!(app.edit_buf, "=B1");
        // Shift+Down keeps anchor at B1, advances target to B2 → range B1:B2.
        assert!(app.try_pointing_arrow(1, 0, true));
        assert_eq!(app.edit_buf, "=B1:B2");
        let p = app.pointing.unwrap();
        assert_eq!((p.anchor_row, p.anchor_col), (0, 1));
        assert_eq!((p.target_row, p.target_col), (1, 1));
    }

    #[test]
    fn shift_arrow_back_through_anchor_collapses_then_extends_other_side() {
        // Anchor pinned at B2 (the entry cell). Shift+Up steps target
        // upward through the anchor and out the other side; the rendered
        // text follows: B2:B3 → B2 → B1:B2.
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 1;
        app.start_edit();
        app.edit_insert('=');
        // Down → enter on B2 (cursor_col 1 + Down).
        app.try_pointing_arrow(1, 0, false);
        assert_eq!(app.edit_buf, "=B2");
        // Shift+Down → B2:B3.
        app.try_pointing_arrow(1, 0, true);
        assert_eq!(app.edit_buf, "=B2:B3");
        // Shift+Up → target back to B2 → collapsed single-cell B2.
        app.try_pointing_arrow(-1, 0, true);
        assert_eq!(app.edit_buf, "=B2");
        // Shift+Up again → target B1, anchor still B2 → range B1:B2
        // (rewrite_pointing_text normalises top-left → bottom-right).
        app.try_pointing_arrow(-1, 0, true);
        assert_eq!(app.edit_buf, "=B1:B2");
        let p = app.pointing.unwrap();
        assert_eq!((p.anchor_row, p.anchor_col), (1, 1));
        assert_eq!((p.target_row, p.target_col), (0, 1));
    }

    #[test]
    fn plain_arrow_during_pointing_still_collapses_after_shift_extension() {
        // After extending to a range, plain Arrow must still collapse
        // back to a single cell (the user's documented mental model).
        let mut app = make_test_app();
        app.cursor_row = 0;
        app.cursor_col = 1;
        app.start_edit();
        app.edit_insert('=');
        app.try_pointing_arrow(1, 0, false); // =B2
        app.try_pointing_arrow(1, 0, true); // =B2:B3
        assert_eq!(app.edit_buf, "=B2:B3");
        // Plain Down → B4 (single cell, anchor reset).
        app.try_pointing_arrow(1, 0, false);
        assert_eq!(app.edit_buf, "=B4");
        let p = app.pointing.unwrap();
        assert_eq!((p.anchor_row, p.anchor_col), (3, 1));
        assert_eq!((p.target_row, p.target_col), (3, 1));
    }

    #[test]
    fn shift_arrow_at_non_insertable_caret_falls_through() {
        // Caret right after a closing paren — not an insertable position.
        // Shift+Arrow must not start pointing.
        let mut app = make_test_app();
        app.start_edit();
        for c in "=SUM(A1)".chars() {
            app.edit_insert(c);
        }
        assert!(!app.try_pointing_arrow(0, 1, true));
        assert!(app.pointing.is_none());
    }

    #[test]
    fn add_sheet_appends_and_switches() {
        let mut app = make_test_app();
        assert_eq!(app.sheets.len(), 1);
        let initial_id = app.active_sheet_name().to_string();

        app.add_sheet("Notes").unwrap();
        assert_eq!(app.sheets.len(), 2);
        assert_eq!(app.active_sheet, 1);
        assert_eq!(app.active_sheet_name(), "Notes");
        assert_ne!(app.active_sheet_name(), initial_id);
    }

    #[test]
    fn rename_active_sheet_updates_name_and_persists_cells() {
        let mut app = make_test_app();
        // Seed a cell on the active sheet so we can confirm the FK
        // cascade rewrote it in place.
        app.start_edit();
        for c in "hi".chars() {
            app.edit_insert(c);
        }
        app.confirm_edit();
        let old = app.active_sheet_name().to_string();

        app.rename_active_sheet("Renamed").unwrap();

        assert_eq!(app.active_sheet_name(), "Renamed");
        assert_ne!(app.active_sheet_name(), old);
        // Cell survived the cascade — load_sheet under the new name
        // still finds it.
        let loaded = app.store.load_sheet("Renamed").unwrap();
        assert_eq!(loaded.cells.len(), 1);
    }

    #[test]
    fn rename_active_sheet_rejects_collision() {
        let mut app = make_test_app();
        app.add_sheet("B").unwrap();
        app.switch_sheet(0); // back to Sheet1
        let err = app.rename_active_sheet("B").unwrap_err();
        assert!(err.contains("already exists"));
    }

    #[test]
    fn rename_active_sheet_rejects_empty() {
        let mut app = make_test_app();
        let err = app.rename_active_sheet("   ").unwrap_err();
        assert!(err.contains("empty"));
    }

    #[test]
    fn rename_active_sheet_to_self_is_noop() {
        let mut app = make_test_app();
        let name = app.active_sheet_name().to_string();
        app.rename_active_sheet(&name).unwrap();
        assert_eq!(app.active_sheet_name(), name);
    }

    #[test]
    fn sheet_rename_command_runs_through_run_command() {
        let mut app = make_test_app();
        run_ex(&mut app, "sheet rename My Notes");
        assert_eq!(app.active_sheet_name(), "My Notes");
        // Alias :sheet ren works too.
        run_ex(&mut app, "sheet ren Final");
        assert_eq!(app.active_sheet_name(), "Final");
        // Missing argument keeps name + reports usage.
        run_ex(&mut app, "sheet rename");
        assert_eq!(app.active_sheet_name(), "Final");
        assert!(app.status.contains("rename"));
    }

    #[test]
    fn next_and_prev_sheet_wrap() {
        let mut app = make_test_app();
        app.add_sheet("S2").unwrap();
        app.add_sheet("S3").unwrap();
        assert_eq!(app.active_sheet, 2);

        app.next_sheet();
        assert_eq!(app.active_sheet, 0);
        app.prev_sheet();
        assert_eq!(app.active_sheet, 2);
    }

    fn run_ex(app: &mut App, cmd: &str) {
        app.start_command();
        app.edit_buf = cmd.to_string();
        app.edit_cursor = cmd.len();
        let _ = app.run_command();
    }

    #[test]
    fn tabfirst_jumps_to_first_sheet() {
        let mut app = make_test_app();
        app.add_sheet("S2").unwrap();
        app.add_sheet("S3").unwrap();
        assert_eq!(app.active_sheet, 2);
        run_ex(&mut app, "tabfirst");
        assert_eq!(app.active_sheet, 0);
        // Aliases land on the same sheet.
        app.switch_sheet(2);
        run_ex(&mut app, "tabr");
        assert_eq!(app.active_sheet, 0);
    }

    #[test]
    fn tablast_jumps_to_last_sheet() {
        let mut app = make_test_app();
        app.add_sheet("S2").unwrap();
        app.add_sheet("S3").unwrap();
        app.switch_sheet(0);
        assert_eq!(app.active_sheet, 0);
        run_ex(&mut app, "tablast");
        assert_eq!(app.active_sheet, 2);
        // Alias.
        app.switch_sheet(0);
        run_ex(&mut app, "tabl");
        assert_eq!(app.active_sheet, 2);
    }

    #[test]
    fn switch_sheet_clears_clipboard_mark() {
        let mut app = make_test_app();
        app.add_sheet("S2").unwrap();
        app.set_clipboard_mark(ClipMarkMode::Copy);
        assert!(app.clipboard_mark.is_some());

        app.prev_sheet();
        // [sheet.clipboard.sheet-switch-clears-mark]
        assert!(app.clipboard_mark.is_none());
    }

    #[test]
    fn delete_active_sheet_refuses_when_only_one() {
        let mut app = make_test_app();
        let err = app.delete_active_sheet().unwrap_err();
        assert!(err.contains("Cannot delete"));
        assert_eq!(app.sheets.len(), 1);
    }

    #[test]
    fn delete_active_sheet_drops_and_falls_back() {
        let mut app = make_test_app();
        app.add_sheet("S2").unwrap();
        app.add_sheet("S3").unwrap();
        // active = 2 (S3). Delete it; active should fall back to 1 (S2).
        app.delete_active_sheet().unwrap();
        assert_eq!(app.sheets.len(), 2);
        assert_eq!(app.active_sheet, 1);
        assert_eq!(app.active_sheet_name(), "S2");
    }

    #[test]
    fn switch_sheet_isolates_cell_data() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "first");
        app.add_sheet("S2").unwrap();
        // Now on S2 — should be empty.
        assert_eq!(app.get_raw(0, 0), "");
        type_into(&mut app, 0, 0, "second");
        assert_eq!(app.get_raw(0, 0), "second");

        app.prev_sheet(); // back to Sheet1
        assert_eq!(app.get_raw(0, 0), "first");
    }

    #[test]
    fn insert_row_above_cursor_shifts_content_down() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "first");
        type_into(&mut app, 1, 0, "second");
        app.cursor_row = 1; // we'll insert above this
        app.insert_rows_at_cursor(InsertSide::AboveOrLeft, 1).unwrap();
        assert_eq!(app.get_raw(0, 0), "first");
        assert_eq!(app.get_raw(1, 0), ""); // newly blank
        assert_eq!(app.get_raw(2, 0), "second");
    }

    #[test]
    fn insert_col_below_right_keeps_cursor_col_intact() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "A");
        type_into(&mut app, 0, 1, "B");
        app.cursor_col = 0;
        app.insert_cols_at_cursor(InsertSide::BelowOrRight, 1).unwrap();
        // A stays at col 0; new blank lands at col 1; B shifts to col 2.
        assert_eq!(app.get_raw(0, 0), "A");
        assert_eq!(app.get_raw(0, 1), "");
        assert_eq!(app.get_raw(0, 2), "B");
    }

    #[test]
    fn delete_row_requires_confirm_before_executing() {
        let mut app = make_test_app();
        type_into(&mut app, 1, 0, "victim");
        app.cursor_row = 1;
        app.cursor_col = 0;

        app.request_delete_row();
        assert!(app.pending_confirm.is_some());
        // Cell is still there until confirm.
        assert_eq!(app.get_raw(1, 0), "victim");

        app.confirm_pending();
        assert!(app.pending_confirm.is_none());
        assert_eq!(app.get_raw(1, 0), "");
    }

    #[test]
    fn cancel_pending_confirm_leaves_data_intact() {
        let mut app = make_test_app();
        type_into(&mut app, 1, 0, "keep");
        app.cursor_row = 1;
        app.cursor_col = 0;
        app.request_delete_row();
        app.cancel_pending_confirm();
        assert!(app.pending_confirm.is_none());
        assert_eq!(app.get_raw(1, 0), "keep");
    }

    #[test]
    fn delete_col_with_formula_dependent_rewrites_to_ref_error() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "10"); // A1
        type_into(&mut app, 0, 1, "=A1*2"); // B1
        app.cursor_row = 0;
        app.cursor_col = 0; // delete column A
        app.request_delete_col();
        app.confirm_pending();

        // The =A1*2 formula was at (0, 1). After deleting column 0, it
        // shifts to (0, 0) and its A1 ref becomes #REF!.
        assert!(app.get_raw(0, 0).contains("#REF!"));
    }

    #[test]
    fn parse_insert_side_aliases() {
        assert_eq!(parse_insert_side("above"), Some(InsertSide::AboveOrLeft));
        assert_eq!(parse_insert_side("left"), Some(InsertSide::AboveOrLeft));
        assert_eq!(parse_insert_side("below"), Some(InsertSide::BelowOrRight));
        assert_eq!(parse_insert_side("right"), Some(InsertSide::BelowOrRight));
        assert_eq!(parse_insert_side("nope"), None);
    }

    #[test]
    fn jump_works_vertically() {
        // Column 3 filled at rows 1,2,3 then a gap then row 8.
        let f = filled_set(&[(1, 3), (2, 3), (3, 3), (8, 3)]);
        assert_eq!(compute_jump_target(1, 3, 1, 0, R, C, &f), (3, 3));
        assert_eq!(compute_jump_target(3, 3, 1, 0, R, C, &f), (8, 3));
        assert_eq!(compute_jump_target(8, 3, -1, 0, R, C, &f), (3, 3));
    }

    // ── :set mouse / :set nomouse / :set mouse? ─────────────────────

    #[test]
    fn mouse_capture_is_on_by_default() {
        // App::new defaults `mouse_enabled = true` so click/drag/scroll
        // "just work" without the user having to discover `:set mouse`.
        let app = make_test_app();
        assert!(app.mouse_enabled);
    }

    #[test]
    fn set_nomouse_disables_capture() {
        let mut app = make_test_app();
        run_ex(&mut app, "set nomouse");
        assert!(!app.mouse_enabled);
        assert_eq!(app.status, "mouse capture disabled");
    }

    #[test]
    fn set_mouse_re_enables_capture() {
        let mut app = make_test_app();
        run_ex(&mut app, "set nomouse");
        assert!(!app.mouse_enabled);
        run_ex(&mut app, "set mouse");
        assert!(app.mouse_enabled);
        assert_eq!(app.status, "mouse capture enabled");
    }

    #[test]
    fn set_mouse_query_reports_current_state() {
        let mut app = make_test_app();
        run_ex(&mut app, "set mouse?");
        assert_eq!(app.status, "mouse=on"); // default
        run_ex(&mut app, "set nomouse");
        run_ex(&mut app, "set mouse?");
        assert_eq!(app.status, "mouse=off");
    }

    #[test]
    fn set_mouse_is_idempotent() {
        // Running `:set mouse` on an already-mouse-on app stays on.
        let mut app = make_test_app();
        run_ex(&mut app, "set mouse");
        run_ex(&mut app, "set mouse");
        assert!(app.mouse_enabled);
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn datetime_tag_for_cell_returns_jdate_for_iso_literal() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "2025-04-27");
        assert_eq!(app.datetime_tag_for_cell(0, 0), Some("jdate"));
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn datetime_tag_for_cell_is_none_for_plain_text_and_numbers() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "hello");
        type_into(&mut app, 1, 0, "42");
        assert!(app.datetime_tag_for_cell(0, 0).is_none());
        assert!(app.datetime_tag_for_cell(1, 0).is_none());
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn datetime_tag_for_cell_is_none_for_hyperlink() {
        // Hyperlinks are Custom values too — confirm the datetime probe
        // declines them so we don't paint links peach.
        let mut app = make_test_app();
        type_into(
            &mut app,
            0,
            0,
            "=HYPERLINK(\"https://example.com\", \"click\")",
        );
        assert!(app.datetime_tag_for_cell(0, 0).is_none());
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn today_command_inserts_current_date_literal() {
        let mut app = make_test_app();
        run_ex(&mut app, "today");
        let raw = app.cells.iter().find(|c| c.row_idx == 0 && c.col_idx == 0);
        let raw = raw.expect("cell A1 should exist after :today").raw_value.clone();
        // ISO 8601 shape: YYYY-MM-DD, ten chars, two dashes, all digits
        // around them. We don't pin the exact date — that drifts across
        // the date boundary.
        assert_eq!(raw.len(), 10, "expected ISO date, got {raw:?}");
        let bytes = raw.as_bytes();
        assert_eq!(bytes[4], b'-', "{raw}");
        assert_eq!(bytes[7], b'-', "{raw}");
        assert_eq!(app.datetime_tag_for_cell(0, 0), Some("jdate"));
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn now_command_inserts_current_datetime_literal() {
        let mut app = make_test_app();
        run_ex(&mut app, "now");
        let cell = app
            .cells
            .iter()
            .find(|c| c.row_idx == 0 && c.col_idx == 0)
            .expect("cell A1 should exist after :now");
        // ISO datetime shape: 19 chars, 'T' at index 10.
        assert_eq!(cell.raw_value.len(), 19, "{:?}", cell.raw_value);
        assert_eq!(cell.raw_value.as_bytes()[10], b'T');
        assert_eq!(app.datetime_tag_for_cell(0, 0), Some("jdatetime"));
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn fmt_date_applies_strftime_to_jdate_cell() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "2025-04-27");
        // type_into advances the cursor; rewind so :fmt targets (0,0).
        app.cursor_row = 0;
        app.cursor_col = 0;
        run_ex(&mut app, "fmt date %a %b %d");
        // 2025-04-27 is a Sunday → "Sun Apr 27"
        assert_eq!(app.displayed_for(0, 0), "Sun Apr 27");
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn fmt_date_no_op_for_non_datetime_cell() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "hello");
        app.cursor_row = 0;
        app.cursor_col = 0;
        run_ex(&mut app, "fmt date %Y-%m-%d");
        // Format is attached but harmless — display stays as the text.
        assert_eq!(app.displayed_for(0, 0), "hello");
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn fmt_nodate_clears_strftime_override() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "2025-04-27");
        app.cursor_row = 0;
        app.cursor_col = 0;
        run_ex(&mut app, "fmt date %a");
        assert_eq!(app.displayed_for(0, 0), "Sun");
        run_ex(&mut app, "fmt nodate");
        assert_eq!(app.displayed_for(0, 0), "2025-04-27");
    }

    #[cfg(feature = "datetime")]
    #[test]
    fn fmt_date_persists_through_format_json_round_trip() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "2025-04-27");
        app.cursor_row = 0;
        app.cursor_col = 0;
        run_ex(&mut app, "fmt date %Y/%m/%d");
        // Read the cell's format JSON and confirm the date axis is in it.
        let json = app.get_format_json_raw(0, 0).unwrap();
        assert!(
            json.contains(r#""df":"%Y/%m/%d""#),
            "json missing df field: {json}"
        );
        // Re-parse and confirm the round-trip is faithful.
        let fmt = format::parse_format_json(&json).unwrap();
        assert_eq!(fmt.date.as_deref(), Some("%Y/%m/%d"));
    }

    // ── Ctrl+A: select_data_region ──────────────────────────────────

    fn populated(cells: &[(u32, u32)]) -> HashSet<(u32, u32)> {
        cells.iter().copied().collect()
    }

    #[test]
    fn data_region_single_populated_cell_returns_one_by_one_rect() {
        let p = populated(&[(5, 5)]);
        assert_eq!(compute_data_region(5, 5, &p), Some((5, 5, 5, 5)));
    }

    #[test]
    fn data_region_orthogonal_block_yields_full_bbox_from_interior() {
        // A1:D12 fully populated.
        let mut cells = Vec::new();
        for r in 0..=11 {
            for c in 0..=3 {
                cells.push((r, c));
            }
        }
        let p = populated(&cells);
        // Seed from interior cell C5 (row 4, col 2).
        assert_eq!(compute_data_region(4, 2, &p), Some((0, 0, 11, 3)));
    }

    #[test]
    fn data_region_diagonal_neighbours_connect() {
        // Three cells touching only at corners.
        let p = populated(&[(0, 0), (1, 1), (2, 2)]);
        assert_eq!(compute_data_region(0, 0, &p), Some((0, 0, 2, 2)));
    }

    #[test]
    fn data_region_internal_empty_cells_included_in_bbox() {
        // Perimeter of a 4×4 block — center cells are empty but still
        // inside the bbox of the connected ring.
        let p = populated(&[
            (0, 0), (0, 1), (0, 2), (0, 3),
            (1, 0),                  (1, 3),
            (2, 0),                  (2, 3),
            (3, 0), (3, 1), (3, 2), (3, 3),
        ]);
        assert_eq!(compute_data_region(0, 0, &p), Some((0, 0, 3, 3)));
    }

    #[test]
    fn data_region_isolated_island_does_not_grab_disjoint_island() {
        // Two blobs separated by >1 empty cell in every direction.
        let p = populated(&[(0, 0), (0, 1), (5, 5), (5, 6)]);
        assert_eq!(compute_data_region(0, 0, &p), Some((0, 0, 0, 1)));
        assert_eq!(compute_data_region(5, 6, &p), Some((5, 5, 5, 6)));
    }

    #[test]
    fn data_region_seed_on_empty_cell_returns_none() {
        let p = populated(&[(0, 0), (1, 1)]);
        assert_eq!(compute_data_region(5, 5, &p), None);
    }

    #[test]
    fn select_data_region_on_block_enters_visual_with_correct_rect() {
        let mut app = make_test_app();
        // Populate A1:D12.
        for r in 0..=11u32 {
            for c in 0..=3u32 {
                type_into(&mut app, r, c, "x");
            }
        }
        // Cursor on C5 (row 4, col 2).
        app.cursor_row = 4;
        app.cursor_col = 2;
        app.select_data_region();
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
        assert_eq!(app.selection_range(), Some((0, 0, 11, 3)));
        assert_eq!(app.selection_anchor, Some((0, 0)));
        assert_eq!((app.cursor_row, app.cursor_col), (11, 3));
    }

    #[test]
    fn select_data_region_on_empty_cell_selects_entire_sheet() {
        let mut app = make_test_app();
        type_into(&mut app, 0, 0, "x");
        // Cursor on a known-empty cell.
        app.cursor_row = 5;
        app.cursor_col = 5;
        app.select_data_region();
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
        assert_eq!(app.selection_range(), Some((0, 0, MAX_ROW, MAX_COL)));
    }

    #[test]
    fn select_data_region_on_isolated_cell_is_one_by_one() {
        let mut app = make_test_app();
        type_into(&mut app, 3, 3, "lonely");
        app.cursor_row = 3;
        app.cursor_col = 3;
        app.select_data_region();
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
        assert_eq!(app.selection_range(), Some((3, 3, 3, 3)));
        assert_eq!(app.selection_anchor, Some((3, 3)));
    }

    #[test]
    fn select_data_region_from_v_line_re_anchors_as_visual_cell() {
        let mut app = make_test_app();
        for r in 0..=2u32 {
            for c in 0..=2u32 {
                type_into(&mut app, r, c, "x");
            }
        }
        app.cursor_row = 1;
        app.cursor_col = 1;
        // Enter V-LINE first — Ctrl+A should overwrite to a cell-rect, not
        // keep the row-extents-forced behavior.
        app.enter_visual(VisualKind::Row);
        app.select_data_region();
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
        assert_eq!(app.selection_range(), Some((0, 0, 2, 2)));
    }

    #[test]
    fn select_data_region_treats_whitespace_only_as_empty() {
        let mut app = make_test_app();
        // The seed cell holds only whitespace — `is_filled`/the populated
        // filter trims it, so Ctrl+A falls through to whole-sheet.
        type_into(&mut app, 0, 0, "   ");
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.select_data_region();
        assert_eq!(app.selection_range(), Some((0, 0, MAX_ROW, MAX_COL)));
    }
}
