mod app;
#[cfg(feature = "datetime")]
mod datetime;
mod format;
mod hyperlink;
mod patch_cli;
mod shell;
mod store;
mod theme;
mod tutor;
mod ui;

#[cfg(test)]
mod snapshots;

use std::io;
use std::process;

use app::{
    compute_jump_target, compute_word_backward, compute_word_forward, App, ClipMarkMode,
    CommandOutcome, EditAction, EditKind, MarkAction, Mode, Operator, PastedCell, PastedGrid,
    SearchDir, VisualKind, YankRegister, MAX_COL, MAX_ROW,
};
use clap::{Parser, Subcommand, ValueEnum};
use arboard::Clipboard;
use crossterm::{
    event::{
        self, DisableMouseCapture, EnableMouseCapture, Event, KeyCode, KeyEventKind, KeyModifiers,
        MouseButton, MouseEvent, MouseEventKind,
    },
    execute,
    terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::prelude::*;
use store::{
    coords::{from_cell_id, to_cell_id},
    Store,
};

/// Terminal spreadsheet backed by SQLite
#[derive(Parser)]
#[command(name = "vlotus", version, about)]
struct Cli {
    #[command(subcommand)]
    command: Option<Command>,

    /// SQLite database file (in-memory if omitted)
    #[arg(global = true)]
    db: Option<String>,
}

#[derive(Subcommand)]
enum Command {
    /// Evaluate a cell and print its computed value to stdout
    Eval {
        /// SQLite database file
        db: String,
        /// Cell reference (e.g. A1, B34, AA10)
        cell: String,
    },
    /// Open the bundled vimtutor-style lesson workbook (read-only,
    /// in-memory). Each lesson is one sheet; `gt`/`gT` walks them.
    Tutor,
    /// Work with `.lpatch` files: apply, show, invert, combine.
    Patch {
        #[command(subcommand)]
        op: PatchOp,
    },
}

#[derive(Subcommand)]
enum PatchOp {
    /// Apply a patch to a workbook, recalculate, and commit.
    Apply {
        db: String,
        patch: String,
        /// Apply the inverse of the patch instead.
        #[arg(long)]
        invert: bool,
        /// What to do when an incoming change conflicts with a row
        /// already present in the destination.
        #[arg(long, value_enum, default_value_t = ConflictPolicy::Omit)]
        on_conflict: ConflictPolicy,
    },
    /// Print a human-readable diff (Sheet1!A1: foo → bar).
    Show { patch: String },
    /// Write the inverse of `<input>` to `<output>`.
    Invert { input: String, output: String },
    /// Concatenate patches in order via SQLite's changegroup module.
    Combine {
        output: String,
        #[arg(required = true)]
        patches: Vec<String>,
    },
    /// Derive a patch from the diff between two workbooks.
    Diff {
        from: String,
        to: String,
        output: String,
    },
}

#[derive(Copy, Clone, Debug, ValueEnum)]
enum ConflictPolicy {
    Omit,
    Replace,
    Abort,
}

fn main() {
    let cli = Cli::parse();

    match cli.command {
        Some(Command::Eval { db, cell }) => {
            process::exit(run_eval(&db, &cell));
        }
        Some(Command::Tutor) => {
            if let Err(e) = run_tutor() {
                eprintln!("Error: {e}");
                process::exit(1);
            }
        }
        Some(Command::Patch { op }) => {
            process::exit(patch_cli::dispatch(op));
        }
        None => {
            let db_path = cli.db.as_deref();
            if let Err(e) = run_tui(db_path) {
                eprintln!("Error: {e}");
                process::exit(1);
            }
        }
    }
}

fn run_eval(db_path: &str, cell_ref: &str) -> i32 {
    let cell_ref_upper = cell_ref.to_uppercase();
    let (row_idx, col_idx) = match from_cell_id(&cell_ref_upper) {
        Some(coords) => coords,
        None => {
            eprintln!("Invalid cell reference: {cell_ref}");
            return 1;
        }
    };

    let mut store = match Store::open(std::path::Path::new(db_path)) {
        Ok(s) => s,
        Err(e) => {
            eprintln!("Failed to open database {db_path}: {e}");
            return 1;
        }
    };

    let sheet_name: Option<String> = store
        .list_sheets()
        .ok()
        .and_then(|sheets| sheets.into_iter().next().map(|s| s.name));
    let sheet_name = match sheet_name {
        Some(n) => n,
        None => {
            eprintln!("No sheets found in {db_path}");
            return 1;
        }
    };

    if let Err(e) = store.recalculate(&sheet_name) {
        eprintln!("Recalculation error: {e}");
        return 1;
    }

    match store.get_computed(&sheet_name, row_idx, col_idx) {
        Ok(Some(val)) => {
            println!("{val}");
            0
        }
        Ok(None) => 0,
        Err(e) => {
            eprintln!("get_computed: {e}");
            1
        }
    }
}

fn run_tui(db_path: Option<&str>) -> io::Result<()> {
    let store = match db_path {
        Some(path) => Store::open(std::path::Path::new(path))
            .expect("Failed to open database"),
        None => Store::open_in_memory().expect("Failed to create in-memory database"),
    };
    let mut app = App::new(store, db_path.unwrap_or(":memory:"));
    run_terminal(&mut app)
}

/// Open the bundled tutorial workbook. Always in-memory so user
/// mutations evaporate at exit — the curriculum stays pristine.
fn run_tutor() -> io::Result<()> {
    let mut store = Store::open_in_memory().expect("Failed to create in-memory database");
    if let Err(e) = tutor::seed_tutor_db(&mut store) {
        return Err(io::Error::other(format!("tutor seed: {e}")));
    }
    // Seed work is part of the baseline; commit it so user edits on
    // top of the curriculum show up as "modified" relative to seed.
    let _ = store.commit();
    let mut app = App::new(store, "tutor");
    app.status = "Welcome to vlotus tutor — read column A and follow along.".into();
    run_terminal(&mut app)
}

/// Shared boilerplate: enter raw mode + alternate screen, run the
/// event loop, restore the terminal on exit.
fn run_terminal(app: &mut App) -> io::Result<()> {
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    // Mouse capture is opt-in via `:set mouse` (T7). `run_loop` diffs
    // `app.mouse_enabled` each tick and emits enable/disable as needed,
    // so a session that never calls `:set mouse` never captures the
    // mouse — preserving native terminal text-selection.
    let result = run_loop(&mut terminal, app);

    disable_raw_mode()?;
    if app.mouse_enabled {
        let _ = execute!(terminal.backend_mut(), DisableMouseCapture);
    }
    execute!(terminal.backend_mut(), LeaveAlternateScreen)?;
    terminal.show_cursor()?;

    result
}

fn run_loop(
    terminal: &mut Terminal<CrosstermBackend<io::Stdout>>,
    app: &mut App,
) -> io::Result<()> {
    // Live state of crossterm's mouse capture, kept in sync with
    // `app.mouse_enabled`. Diffed after every event dispatch so `:set
    // mouse` / `:set nomouse` toggles take effect on the next tick.
    let mut capture_on = false;
    loop {
        terminal.draw(|f| ui::draw(f, app))?;

        match event::read()? {
            Event::Key(key) => {
                if key.kind != KeyEventKind::Press {
                    continue;
                }

                match app.mode {
                    Mode::Edit => handle_edit_key(app, key),
                    Mode::Command => {
                        if let CommandOutcome::Quit { .. } = handle_command_key(app, key) {
                            return Ok(());
                        }
                    }
                    Mode::Shell => {
                        handle_shell_key(app, key);
                    }
                    Mode::Search(_) => handle_search_key(app, key),
                    Mode::ColorPicker => handle_color_picker_key(app, key),
                    Mode::PatchShow => handle_patch_show_key(app, key),
                    Mode::Nav | Mode::Visual(_) => handle_nav_key(app, key),
                }
            }
            Event::Mouse(me) => {
                let size = terminal.size()?;
                let area = Rect::new(0, 0, size.width, size.height);
                handle_mouse(app, me, area);
            }
            _ => {}
        }

        // Apply any pending mouse-capture toggle. Crossterm's enable/disable
        // are no-ops when the state already matches, but we track our own
        // bool so the on-shutdown disable in `run_terminal` knows whether
        // capture was ever turned on.
        if app.mouse_enabled != capture_on {
            if app.mouse_enabled {
                execute!(terminal.backend_mut(), EnableMouseCapture)?;
            } else {
                execute!(terminal.backend_mut(), DisableMouseCapture)?;
            }
            capture_on = app.mouse_enabled;
        }
    }
}

/// Maximum gap between two clicks on the same cell that still counts as
/// a double-click. 400ms matches macOS / Windows defaults — feels right
/// to most users; bump if anyone complains.
const DOUBLE_CLICK_MS: u64 = 400;

/// Pure double-click predicate: `prev` is the previous click (if any),
/// `now` is the new click's timestamp, `cell` is the cell it landed on.
/// Returns true iff the previous click is within `DOUBLE_CLICK_MS` and
/// hit the same cell.
fn is_double_click(
    prev: Option<(std::time::Instant, (u32, u32))>,
    now: std::time::Instant,
    cell: (u32, u32),
) -> bool {
    let Some((t, prev_cell)) = prev else {
        return false;
    };
    if prev_cell != cell {
        return false;
    }
    now.duration_since(t).as_millis() < DOUBLE_CLICK_MS as u128
}

/// Dispatch a mouse event to the appropriate App-level handler. Currently
/// covers click-to-select-cell (T1), click-and-drag-to-visual (T2),
/// scroll-wheel (T3), and double-click-to-Edit (T4); row/column-header
/// and tabline clicks are owned by later tickets in the mouse epic
/// (`att tree i38d0opr`).
fn handle_mouse(app: &mut App, me: MouseEvent, area: Rect) {
    match me.kind {
        MouseEventKind::Down(MouseButton::Left) => {
            match ui::cell_at(app, area, me.column, me.row) {
                Some(ui::HitTarget::Cell { row, col }) => {
                    let now = std::time::Instant::now();
                    // T4: second click on the same cell within the
                    // threshold → enter Edit mode. Pre-fill the buffer
                    // with the cell's raw value (formula text, not
                    // computed display) so the user can refine in place.
                    //
                    // Suppress in Edit mode: a fast second click while
                    // editing should keep doing T8's ref-insert path
                    // (or commit-then-move) instead of throwing away the
                    // in-flight buffer via start_edit.
                    if !matches!(app.mode, Mode::Edit)
                        && is_double_click(app.last_click, now, (row, col))
                    {
                        app.jump_cursor_to(row, col);
                        app.start_edit();
                        app.last_click = None;
                        app.drag_anchor = None;
                        return;
                    }
                    // Ctrl-click on a URL cell follows the link instead
                    // of selecting. Move the cursor first so the user
                    // sees which cell they hit, then open. If the cell
                    // has no URL, fall through to the normal click
                    // path so plain ctrl-clicks still feel responsive.
                    if me.modifiers.contains(KeyModifiers::CONTROL)
                        && !matches!(app.mode, Mode::Edit)
                    {
                        app.jump_cursor_to(row, col);
                        if let Some(url) = app.url_under_cursor() {
                            match hyperlink::open_url(&url) {
                                Ok(()) => app.status = format!("Opened {url}"),
                                Err(e) => app.status = format!("Open failed: {e}"),
                            }
                            app.last_click = Some((now, (row, col)));
                            app.drag_anchor = None;
                            return;
                        }
                    }
                    handle_left_click_cell(app, row, col);
                    // T2: remember where the press happened so a follow-up
                    // Drag can anchor the new selection at the original
                    // click point.
                    app.drag_anchor = Some((row, col));
                    app.last_click = Some((now, (row, col)));
                }
                // T5: tab click switches the active sheet. `switch_sheet`
                // handles the data reload + status line + clipboard-mark
                // clear; it's a no-op when the click hits the already-
                // active tab.
                Some(ui::HitTarget::Tab(idx)) => {
                    app.switch_sheet(idx);
                }
                // T11: row-header click during formula edit inserts a
                // whole-row ref (`1:1`) at the caret. Falls through to
                // T6's V-LINE entry when not in formula context.
                Some(ui::HitTarget::RowHeader(row)) => {
                    if app.insert_row_ref_at_caret(row) {
                        return;
                    }
                    app.mode = Mode::Visual(VisualKind::Row);
                    app.selection_anchor = Some((row, app.cursor_col));
                    app.jump_cursor_to(row, app.cursor_col);
                    app.drag_anchor = Some((row, app.cursor_col));
                }
                // T10: column-header click during formula edit inserts a
                // whole-column ref (`B:B`) at the caret. Falls through to
                // T6+T16's V-COLUMN entry otherwise.
                Some(ui::HitTarget::ColumnHeader(col)) => {
                    if app.insert_col_ref_at_caret(col) {
                        return;
                    }
                    app.mode = Mode::Visual(VisualKind::Column);
                    app.selection_anchor = Some((app.cursor_row, col));
                    app.jump_cursor_to(app.cursor_row, col);
                    app.drag_anchor = Some((app.cursor_row, col));
                }
                None => {}
            }
        }
        MouseEventKind::Drag(MouseButton::Left) => {
            match ui::cell_at(app, area, me.column, me.row) {
                Some(ui::HitTarget::Cell { row, col }) => {
                    // T9: drag during formula edit extends the active
                    // pointing ref into a range. anchor stays at the
                    // original click point; target follows the drag.
                    if app.mode == Mode::Edit && app.pointing.is_some() {
                        app.drag_pointing_target(row, col);
                        return;
                    }
                    handle_drag_to_cell(app, row, col);
                }
                // T13: row-header drag during formula edit extends a
                // row-range pointing into `1:5`. Falls through to T6's
                // V-LINE drag-extend otherwise.
                Some(ui::HitTarget::RowHeader(row)) => {
                    if app.mode == Mode::Edit
                        && matches!(
                            app.pointing.map(|p| p.kind),
                            Some(app::PointingKind::Row)
                        )
                    {
                        app.drag_pointing_target(row, 0);
                        return;
                    }
                    if app.drag_anchor.is_some() && matches!(app.mode, Mode::Visual(_)) {
                        app.jump_cursor_to(row, app.cursor_col);
                    }
                }
                // T12: column-header drag during formula edit extends a
                // col-range pointing into `B:E`. Falls through to T6's
                // V-COLUMN drag-extend otherwise.
                Some(ui::HitTarget::ColumnHeader(col)) => {
                    if app.mode == Mode::Edit
                        && matches!(
                            app.pointing.map(|p| p.kind),
                            Some(app::PointingKind::Column)
                        )
                    {
                        app.drag_pointing_target(0, col);
                        return;
                    }
                    if app.drag_anchor.is_some() && matches!(app.mode, Mode::Visual(_)) {
                        app.jump_cursor_to(app.cursor_row, col);
                    }
                }
                // Tab-drag (would be drag-to-reorder, future ticket); ignore.
                Some(ui::HitTarget::Tab(_)) => {}
                // Drag escaped the grid: past the bottom border, into the
                // formula bar, past the right edge, etc. Auto-scroll the
                // viewport one row/column toward the drag and advance the
                // cursor so the Visual selection grows. Each Drag event
                // tick scrolls by 1; OS rate-limiting via mouse motion
                // gives a continuous-feeling scroll without us needing a
                // tick timer.
                None => {
                    handle_drag_past_edge(app, area, me.column, me.row);
                }
            }
        }
        MouseEventKind::Up(MouseButton::Left) => {
            app.drag_anchor = None;
        }
        // T3: scroll wheel moves the viewport without touching the cursor.
        // Shift+ScrollUp/Down → horizontal (terminals that don't emit
        // ScrollLeft/Right natively use this convention).
        MouseEventKind::ScrollUp => {
            if me.modifiers.contains(KeyModifiers::SHIFT) {
                app.scroll_viewport(0, -1);
            } else {
                app.scroll_viewport(-1, 0);
            }
        }
        MouseEventKind::ScrollDown => {
            if me.modifiers.contains(KeyModifiers::SHIFT) {
                app.scroll_viewport(0, 1);
            } else {
                app.scroll_viewport(1, 0);
            }
        }
        MouseEventKind::ScrollLeft => {
            app.scroll_viewport(0, -1);
        }
        MouseEventKind::ScrollRight => {
            app.scroll_viewport(0, 1);
        }
        _ => {}
    }
}

fn handle_left_click_cell(app: &mut App, row: u32, col: u32) {
    match app.mode {
        Mode::Edit => {
            // T8: clicking a cell while editing a formula at an insertable
            // caret position inserts that cell's A1-style ref into the
            // buffer (same path as the keyboard's pointing-mode arrow).
            // Falls through if the caret isn't in formula context — e.g.
            // editing a plain text value, or caret is mid-ref.
            if app.insert_ref_at_caret(row, col) {
                return;
            }
            // Same-cell click while editing: no-op (user clicked the cell
            // they're already editing). Otherwise commit-without-moving
            // and then jump to the click target — same effect as pressing
            // Enter/Tab to commit, then arrowing over.
            if row == app.cursor_row && col == app.cursor_col {
                return;
            }
            app.confirm_edit_advancing(0, 0);
            app.jump_cursor_to(row, col);
        }
        Mode::Nav | Mode::Visual(_) => {
            // A plain click in Visual mode exits Visual and selects only
            // the clicked cell — Excel / Sheets semantics. Keyboard motions
            // (or a click+drag, which fires a Down then a Drag) extend
            // selection; a single click is a fresh-selection gesture.
            // Shift+click for "extend without resetting" is a future
            // ticket; for now plain click always resets.
            app.clear_pending_motion_state();
            app.mode = Mode::Nav;
            app.jump_cursor_to(row, col);
        }
        // Don't intercept clicks while a `:` / `/` / `?` prompt or the
        // color picker is open — those are modal text/UI surfaces.
        Mode::Command | Mode::Shell | Mode::Search(_) | Mode::ColorPicker | Mode::PatchShow => {}
    }
}

/// Handle a `Drag(Left)` whose position is outside the grid (past an
/// edge). Scrolls the viewport one row / column toward the drag and
/// advances the cursor so the Visual selection grows as the user drags.
/// Each Drag event ticks the scroll by 1; OS-level mouse-motion rate
/// limiting gives the user a continuous scroll feel without us needing
/// our own timer.
///
/// In `Visual::Row` (V-LINE from a row-header click), only vertical
/// scroll is applied — the column dimension is irrelevant since V-LINE
/// auto-spans every column.
fn handle_drag_past_edge(app: &mut App, area: Rect, x: u16, y: u16) {
    if app.drag_anchor.is_none() {
        return;
    }
    let (grid, _) = ui::grid_layout(app, area);

    // Decide which axis(es) to scroll. The grid has a 1-cell border on
    // every side; treat "at or past" the border as "past the edge" so
    // dragging onto the border itself starts scrolling.
    let dr: i32 = if y < grid.y {
        -1
    } else if y >= grid.y + grid.height - 1 {
        1
    } else {
        0
    };
    let dc: i32 = if x < grid.x {
        -1
    } else if x >= grid.x + grid.width - 1 {
        1
    } else {
        0
    };

    if dr == 0 && dc == 0 {
        return;
    }

    app.scroll_viewport(dr, dc);

    // V-LINE: only the row dimension matters; column stays put so the
    // anchor remains unambiguous.
    if matches!(app.mode, Mode::Visual(VisualKind::Row)) {
        if dr != 0 {
            let new_row = (app.cursor_row as i32 + dr).clamp(0, MAX_ROW as i32) as u32;
            app.jump_cursor_to(new_row, app.cursor_col);
        }
        return;
    }

    // Cell-mode (or Nav about to transition): advance the cursor by 1 in
    // each scrolled axis. Letting handle_drag_to_cell run does the
    // Nav→Visual transition + anchor wiring identically to an in-grid
    // drag, so the user experiences past-edge as "drag continues, just
    // off-screen".
    let new_row = (app.cursor_row as i32 + dr).clamp(0, MAX_ROW as i32) as u32;
    let new_col = (app.cursor_col as i32 + dc).clamp(0, MAX_COL as i32) as u32;
    handle_drag_to_cell(app, new_row, new_col);
}

/// Handle a `Drag(Left)` event whose drag-to position resolved to a Cell.
/// In Nav, transitions to `Visual::Cell` anchored at the original click
/// point and updates the cursor. In Visual, the existing anchor sticks
/// (drag extends the selection from where the user originally entered
/// Visual) — equivalent to vim `v` + motion.
fn handle_drag_to_cell(app: &mut App, row: u32, col: u32) {
    // Need a remembered click point to anchor against.
    let Some(anchor) = app.drag_anchor else {
        return;
    };
    // No movement off the current cursor → nothing to do (suppresses the
    // initial Drag event some terminals emit on Down with zero motion).
    if (row, col) == (app.cursor_row, app.cursor_col) {
        return;
    }
    // Don't interfere with text-modes (`:` / `/` / Edit). Edit-mode
    // drag will gain semantics in T9 (range-ref insertion); for now
    // commit-and-track behaves like a plain click already did via T1.
    match app.mode {
        Mode::Nav => {
            app.mode = Mode::Visual(VisualKind::Cell);
            app.selection_anchor = Some(anchor);
            app.jump_cursor_to(row, col);
        }
        Mode::Visual(_) => {
            // Anchor stays put (vim semantics — `v` + motion extends).
            app.jump_cursor_to(row, col);
        }
        Mode::Edit
        | Mode::Command
        | Mode::Shell
        | Mode::Search(_)
        | Mode::ColorPicker
        | Mode::PatchShow => {}
    }
}

fn handle_edit_key(app: &mut App, key: event::KeyEvent) {
    let shift = key.modifiers.contains(KeyModifiers::SHIFT);
    let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
    let autocomplete_open = app.autocomplete.is_some();
    let autocomplete_user_selected = app
        .autocomplete
        .as_ref()
        .map(|a| a.user_selected)
        .unwrap_or(false);

    // [sheet.editing.formula-autocomplete] popup steals Up/Down/Tab/Esc.
    // V9 also recognises vim's Ctrl+n / Ctrl+p / Ctrl+y / Ctrl+e — Up/Down
    // remain the primary bindings (per locked-in decision). Enter only
    // accepts after the user has explicitly navigated the list; otherwise
    // it falls through to commit the cell so a passively-shown popup
    // doesn't silently rewrite what the user typed.
    if autocomplete_open {
        match key.code {
            KeyCode::Up => {
                app.autocomplete_select_prev();
                return;
            }
            KeyCode::Down => {
                app.autocomplete_select_next();
                return;
            }
            KeyCode::Char('p') if ctrl => {
                app.autocomplete_select_prev();
                return;
            }
            KeyCode::Char('n') if ctrl => {
                app.autocomplete_select_next();
                return;
            }
            KeyCode::Tab => {
                app.autocomplete_accept();
                return;
            }
            KeyCode::Enter if autocomplete_user_selected => {
                app.autocomplete_accept();
                return;
            }
            KeyCode::Char('y') if ctrl => {
                app.autocomplete_accept();
                return;
            }
            KeyCode::Esc => {
                app.autocomplete_dismiss();
                return;
            }
            KeyCode::Char('e') if ctrl => {
                app.autocomplete_dismiss();
                return;
            }
            _ => {}
        }
    }

    // [sheet.editing.formula-ref-pointing] arrow keys insert/move a ref
    // when the caret is at an insertable position (right after =, op,
    // comma, or open-paren). Shift+Arrow keeps the anchor pinned and
    // extends the ref into a range (`B2` → `B2:B3`), mirroring mouse drag.
    // Falls through to caret movement otherwise.
    let arrow_dir = match key.code {
        KeyCode::Left => Some((0, -1)),
        KeyCode::Right => Some((0, 1)),
        KeyCode::Up => Some((-1, 0)),
        KeyCode::Down => Some((1, 0)),
        _ => None,
    };
    if let Some((dr, dc)) = arrow_dir {
        if app.try_pointing_arrow(dr, dc, shift) {
            return;
        }
    } else {
        // Any non-arrow editing key exits a pending pointing session.
        app.exit_pointing();
    }

    match key.code {
        KeyCode::Enter => {
            // [sheet.navigation.enter-commit-down]
            app.confirm_edit();
        }
        // [sheet.navigation.tab-commit-right]
        KeyCode::BackTab => app.confirm_edit_advancing(0, -1),
        KeyCode::Tab if shift => app.confirm_edit_advancing(0, -1),
        KeyCode::Tab => app.confirm_edit_advancing(0, 1),
        KeyCode::Esc => {
            // [sheet.editing.escape-cancels]
            app.cancel_edit();
        }
        KeyCode::Char(c) => app.edit_insert(c),
        KeyCode::Backspace => app.edit_backspace(),
        KeyCode::Delete => app.edit_delete(),
        KeyCode::Left => app.edit_move_left(),
        KeyCode::Right => app.edit_move_right(),
        KeyCode::Home => app.edit_home(),
        KeyCode::End => app.edit_end(),
        _ => {}
    }
}

fn handle_search_key(app: &mut App, key: event::KeyEvent) {
    let direction = match app.mode {
        Mode::Search(d) => d,
        _ => return,
    };
    match key.code {
        KeyCode::Enter => app.commit_search(direction),
        KeyCode::Esc => app.cancel_search(),
        KeyCode::Char(c) => app.edit_insert(c),
        KeyCode::Backspace => {
            if app.edit_buf.is_empty() {
                // Empty backspace dismisses the prompt — vim convention.
                app.cancel_search();
            } else {
                app.edit_backspace();
            }
        }
        KeyCode::Delete => app.edit_delete(),
        KeyCode::Left => app.edit_move_left(),
        KeyCode::Right => app.edit_move_right(),
        KeyCode::Home => app.edit_home(),
        KeyCode::End => app.edit_end(),
        _ => {}
    }
}

/// Mode::ColorPicker dispatch. hjkl navigates the swatch grid (l/h
/// step within a row, j/k by rows-of-4); Enter applies; Esc / q
/// cancel; `?` toggles hex-input mode; in hex mode, ASCII hex chars
/// extend the buffer and Backspace deletes.
fn handle_color_picker_key(app: &mut App, key: event::KeyEvent) {
    use crate::app::ColorPickerKind;
    const ROW_STRIDE: i32 = 4; // swatches per row in the popup grid
    let in_hex = app
        .color_picker
        .as_ref()
        .map(|s| s.hex_input.is_some())
        .unwrap_or(false);
    match key.code {
        KeyCode::Esc => app.close_color_picker(),
        KeyCode::Char('q') if !in_hex => app.close_color_picker(),
        KeyCode::Enter => app.apply_color_picker(),
        KeyCode::Char('y') if !in_hex => app.apply_color_picker(),
        KeyCode::Char('?') => app.color_picker_toggle_hex(),
        KeyCode::Char('h') if !in_hex => app.color_picker_step(-1),
        KeyCode::Char('l') if !in_hex => app.color_picker_step(1),
        KeyCode::Char('j') if !in_hex => app.color_picker_step(ROW_STRIDE),
        KeyCode::Char('k') if !in_hex => app.color_picker_step(-ROW_STRIDE),
        KeyCode::Backspace if in_hex => app.color_picker_hex_backspace(),
        KeyCode::Char(c) if in_hex => app.color_picker_hex_input(c),
        _ => {
            // Silence the unused import warning when no other branch
            // touches ColorPickerKind.
            let _ = ColorPickerKind::Fg;
        }
    }
}

/// `Mode::PatchShow` dispatch. Esc / q close; j/k scroll; G / gg jump
/// to bottom / top.
fn handle_patch_show_key(app: &mut App, key: event::KeyEvent) {
    let total = app
        .patch_show
        .as_ref()
        .map(|s| s.lines.len())
        .unwrap_or(0);
    match key.code {
        KeyCode::Esc | KeyCode::Char('q') => {
            app.patch_show = None;
            app.mode = app::Mode::Nav;
        }
        KeyCode::Char('j') | KeyCode::Down => {
            if let Some(s) = app.patch_show.as_mut() {
                if s.scroll + 1 < total {
                    s.scroll += 1;
                }
            }
        }
        KeyCode::Char('k') | KeyCode::Up => {
            if let Some(s) = app.patch_show.as_mut() {
                s.scroll = s.scroll.saturating_sub(1);
            }
        }
        KeyCode::Char('G') => {
            if let Some(s) = app.patch_show.as_mut() {
                s.scroll = total.saturating_sub(1);
            }
        }
        KeyCode::Char('g') => {
            if let Some(s) = app.patch_show.as_mut() {
                s.scroll = 0;
            }
        }
        _ => {}
    }
}

fn handle_command_key(app: &mut App, key: event::KeyEvent) -> CommandOutcome {
    match key.code {
        KeyCode::Enter => app.run_command(),
        KeyCode::Esc => {
            app.cancel_command();
            CommandOutcome::Continue
        }
        KeyCode::Char(c) => {
            app.edit_insert(c);
            CommandOutcome::Continue
        }
        KeyCode::Backspace => {
            app.edit_backspace();
            CommandOutcome::Continue
        }
        KeyCode::Delete => {
            app.edit_delete();
            CommandOutcome::Continue
        }
        KeyCode::Left => {
            app.edit_move_left();
            CommandOutcome::Continue
        }
        KeyCode::Right => {
            app.edit_move_right();
            CommandOutcome::Continue
        }
        KeyCode::Home => {
            app.edit_home();
            CommandOutcome::Continue
        }
        KeyCode::End => {
            app.edit_end();
            CommandOutcome::Continue
        }
        _ => CommandOutcome::Continue,
    }
}

/// `Mode::Shell` dispatch — same shape as `handle_command_key`, only
/// difference is Enter routes to `App::run_shell` (subprocess + paste).
fn handle_shell_key(app: &mut App, key: event::KeyEvent) {
    match key.code {
        KeyCode::Enter => {
            let _ = app.run_shell();
        }
        KeyCode::Esc => app.cancel_shell(),
        KeyCode::Char(c) => app.edit_insert(c),
        KeyCode::Backspace => app.edit_backspace(),
        KeyCode::Delete => app.edit_delete(),
        KeyCode::Left => app.edit_move_left(),
        KeyCode::Right => app.edit_move_right(),
        KeyCode::Home => app.edit_home(),
        KeyCode::End => app.edit_end(),
        _ => {}
    }
}

/// Autofit the cursor column, or every column in the active selection
/// when one is in play (Visual::Cell or Visual::Column). V-LINE (row)
/// selections fall through to single-column behavior since the user is
/// selecting rows, not columns.
fn autofit_selection_or_column(app: &mut App) {
    let (c1, c2) = match (app.selection_range(), app.mode) {
        (Some((_, c1, _, c2)), Mode::Visual(VisualKind::Cell | VisualKind::Column)) => (c1, c2),
        _ => (app.cursor_col, app.cursor_col),
    };
    let mut last_w: u16 = 0;
    let mut last_err: Option<String> = None;
    for col in c1..=c2 {
        match app.autofit_column(col) {
            Ok(w) => last_w = w,
            Err(e) => last_err = Some(e),
        }
    }
    if let Some(e) = last_err {
        app.status = format!("Error: {e}");
        return;
    }
    if c1 == c2 {
        app.status = format!("Column {} → {last_w}", column_letter(c1));
    } else {
        let n = c2 - c1 + 1;
        app.status = format!(
            "Autofit {n} columns ({}:{})",
            column_letter(c1),
            column_letter(c2)
        );
    }
}

fn column_letter(col: u32) -> String {
    to_cell_id(0, col)
        .chars()
        .take_while(|c| c.is_ascii_alphabetic())
        .collect()
}

// Dispatch covers both Normal and Visual. When adding a new binding, pick the
// block by which modes should fire it — placing a Ctrl+ binding in #6 instead
// of #3 was the root cause of the Ctrl+= and Ctrl+A regressions. See also
// `examples/vlotus/CLAUDE.md` § "Where new key bindings go".
//
//   1. Pending guards (pending_confirm, Esc)
//   2. pending_f consumer (mode-agnostic, before motion/operator arms)
//   3. Mode-agnostic Ctrl+ arms (Ctrl+=, Ctrl+A) — fire in Nav AND Visual.
//      Add new mode-agnostic Ctrl+ bindings HERE.
//   4. Visual-mode-only commands (y/d/x/c, V toggle)
//   5. Vim prefix accumulators (digit count, g/z/m/'/`/Ctrl+w)
//   6. Nav-side Ctrl+ block (clipboard, undo, jumps) — past the Visual
//      short-circuit, so effectively Nav-only.
//   7. Bare letter arms (motions, operators) — Nav-only.
fn handle_nav_key(app: &mut App, key: event::KeyEvent) {
    let shift = key.modifiers.contains(KeyModifiers::SHIFT);
    let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);

    // [sheet.delete.row-confirm] / [sheet.delete.column-confirm]
    // When a destructive op is pending, a single y/Y commits and any
    // other key cancels.
    if app.pending_confirm.is_some() {
        match key.code {
            KeyCode::Char('y') | KeyCode::Char('Y') => app.confirm_pending(),
            _ => app.cancel_pending_confirm(),
        }
        return;
    }

    // Esc clears every pending vim-prefix flag, exits Visual mode if
    // active, and otherwise clears the clipboard mark.
    if key.code == KeyCode::Esc {
        let had_pending = app.pending_count.is_some()
            || app.pending_g
            || app.pending_z
            || app.pending_f;
        app.clear_pending_motion_state();
        if matches!(app.mode, Mode::Visual(_)) {
            app.exit_visual();
            return;
        }
        if app.clear_clipboard_mark() || had_pending {
            app.status.clear();
        }
        return;
    }

    // `f`-prefix consumer (format axis). Placed BEFORE the Visual /
    // Nav-only blocks so lowercase second-keys (e.g. `fc` center-align)
    // reach this block instead of `c` being absorbed as the Change
    // operator. The corresponding `g`-prefix consumer lives further
    // down because `g` only resolves whole-key second-letters
    // (`gg`/`gv`/`gt`/`gT`) that don't conflict with operators.
    if app.pending_f {
        app.pending_f = false;
        match key.code {
            // `f$`: apply USD/2 to selection. (Was `g$` pre-migrate-f.)
            KeyCode::Char('$') => {
                app.pending_count = None;
                app.apply_format_update(|f| {
                    f.number = Some(crate::format::NumberFormat::Usd { decimals: 2 });
                });
            }
            // `f%`: apply Percent/0. Default 0 decimals matches gsheets'
            // toolbar `%` button (`5%` not `5.00%` for an existing 0.05).
            KeyCode::Char('%') => {
                app.pending_count = None;
                app.apply_format_update(|f| {
                    f.number = Some(crate::format::NumberFormat::Percent { decimals: 0 });
                });
            }
            // `f.` / `f,`: bump decimals + / -. Auto-applies USD/2 first
            // on unformatted cells (gsheets toolbar parity).
            KeyCode::Char('.') => {
                app.pending_count = None;
                app.bump_format_decimals(1);
            }
            KeyCode::Char(',') => {
                app.pending_count = None;
                app.bump_format_decimals(-1);
            }
            // Text-style toggles. Mnemonics map first letter to axis:
            // bold / italic / underline / strikethrough. Excel/Sheets
            // use Ctrl+B/I/U but those are taken in vlotus
            // (full-page-back / Tab / half-page-up).
            KeyCode::Char('b') => {
                app.pending_count = None;
                app.apply_format_update(|f| f.bold = !f.bold);
            }
            KeyCode::Char('i') => {
                app.pending_count = None;
                app.apply_format_update(|f| f.italic = !f.italic);
            }
            KeyCode::Char('u') => {
                app.pending_count = None;
                app.apply_format_update(|f| f.underline = !f.underline);
            }
            KeyCode::Char('s') => {
                app.pending_count = None;
                app.apply_format_update(|f| f.strike = !f.strike);
            }
            // Alignment overrides. Lowercased — the pending_f consumer
            // intercepts before the operator/Nav-only block, so bare
            // `c` (Change operator), `a/A` (Insert), and `l` (motion
            // right) don't compete here. Defensive `!pending_f`
            // guards exist on those bare arms anyway.
            KeyCode::Char('l') => {
                app.pending_count = None;
                app.apply_format_update(|f| {
                    f.align = Some(crate::format::Align::Left);
                });
            }
            KeyCode::Char('c') => {
                app.pending_count = None;
                app.apply_format_update(|f| {
                    f.align = Some(crate::format::Align::Center);
                });
            }
            KeyCode::Char('r') => {
                app.pending_count = None;
                app.apply_format_update(|f| {
                    f.align = Some(crate::format::Align::Right);
                });
            }
            // `fa` clears the explicit alignment override (auto = let
            // classify_display pick: number = right, bool = center,
            // text = left).
            KeyCode::Char('a') => {
                app.pending_count = None;
                app.apply_format_update(|f| f.align = None);
            }
            // `fF` / `fB`: open the modal swatch picker for fg / bg.
            // Capitals (rather than `ff`) avoid the awkward double-tap
            // and mirror the `:fmt fg` / `:fmt bg` mnemonic.
            KeyCode::Char('F') => {
                app.pending_count = None;
                app.open_color_picker(crate::app::ColorPickerKind::Fg);
            }
            KeyCode::Char('B') => {
                app.pending_count = None;
                app.open_color_picker(crate::app::ColorPickerKind::Bg);
            }
            _ => app.pending_count = None,
        }
        return;
    }

    // Ctrl+= autofits — mode-agnostic so a Visual::Cell / Visual::Column
    // selection can autofit every selected column at once. (Bare `=`
    // stays Nav-only via the formula-start arm below; this arm only
    // fires with the Ctrl modifier.) The V8 binding moved off bare `=`
    // to free that key up for formula-start; many terminals also emit
    // Ctrl+= for plain `+`, so `Ctrl+w =` (in the prefix consumer
    // below) is the symmetric fallback.
    if ctrl && matches!(key.code, KeyCode::Char('=')) {
        app.consume_count();
        autofit_selection_or_column(app);
        return;
    }

    // Ctrl+A: expand selection to the contiguous data region (8-connected
    // bbox of populated cells around the cursor); on an empty cell, selects
    // the entire sheet. Mode-agnostic so it works in Nav and Visual.
    if ctrl && matches!(key.code, KeyCode::Char('a')) {
        app.consume_count();
        app.select_data_region();
        return;
    }

    // Visual-mode commands (vim's y/d/c/o, V toggle). Run BEFORE the
    // motion/digit/prefix handlers so e.g. `y` is yank-selection, not a
    // letter that falls through to redo.
    if matches!(app.mode, Mode::Visual(_)) {
        match key.code {
            KeyCode::Char('y') if !ctrl => {
                visual_yank(app);
                return;
            }
            KeyCode::Char('d') | KeyCode::Char('x') if !ctrl => {
                visual_delete(app);
                return;
            }
            KeyCode::Char('c') if !ctrl => {
                visual_change(app);
                return;
            }
            KeyCode::Char('o') if !ctrl => {
                app.consume_count();
                app.swap_visual_corners();
                return;
            }
            // `v` toggles Cell ⇄ exit; `V` cycles Row → Column → exit (and
            // promotes Cell → Row on the way in). `VV` from Nav lands in
            // V-COLUMN, the keyboard route to whole-column selection.
            KeyCode::Char('v') if !ctrl => {
                app.consume_count();
                if matches!(app.mode, Mode::Visual(VisualKind::Cell)) {
                    app.exit_visual();
                } else {
                    app.switch_visual_kind(VisualKind::Cell);
                }
                return;
            }
            KeyCode::Char('V') if !ctrl => {
                app.consume_count();
                match app.mode {
                    Mode::Visual(VisualKind::Row) => {
                        app.switch_visual_kind(VisualKind::Column);
                    }
                    Mode::Visual(VisualKind::Column) => {
                        app.exit_visual();
                    }
                    _ => {
                        app.switch_visual_kind(VisualKind::Row);
                    }
                }
                return;
            }
            // Insert-mode entries are noops in Visual (V4 will repurpose
            // some of these as operator forms).
            KeyCode::Char('i' | 'I' | 'a' | 'A' | 'O' | 's' | 'S') if !ctrl => {
                app.consume_count();
                return;
            }
            _ => {}
        }
    } else {
        // Nav-mode-only entries: visual mode, operators, paste. `gv`
        // lives inside the g-prefix consumer below.
        match key.code {
            KeyCode::Char('v') if !ctrl => {
                app.consume_count();
                app.enter_visual(VisualKind::Cell);
                return;
            }
            KeyCode::Char('V') if !ctrl => {
                app.consume_count();
                app.enter_visual(VisualKind::Row);
                return;
            }
            // Vim operators — set pending and wait for motion / doubled key.
            KeyCode::Char('d') if !ctrl => {
                app.pending_operator = Some(Operator::Delete);
                app.pending_op_count = app.pending_count.take();
                return;
            }
            // Bare `c` is the Change operator; `fc` center-aligns
            // (defensive guard — pending_f consumer runs first).
            KeyCode::Char('c') if !ctrl && !app.pending_f => {
                app.pending_operator = Some(Operator::Change);
                app.pending_op_count = app.pending_count.take();
                return;
            }
            KeyCode::Char('y') if !ctrl => {
                app.pending_operator = Some(Operator::Yank);
                app.pending_op_count = app.pending_count.take();
                return;
            }
            // Vim `Y` is the shorthand for `yy` — yank current cell.
            KeyCode::Char('Y') if !ctrl => {
                let r = app.cursor_row;
                let c = app.cursor_col;
                app.consume_count();
                apply_operator(app, Operator::Yank, r, c, r, c);
                return;
            }
            // `D` / `C` — clear / change to end of row.
            KeyCode::Char('D') if !ctrl => {
                let r = app.cursor_row;
                let c1 = app.cursor_col;
                let c2 = (0..=MAX_COL)
                    .rev()
                    .find(|&c| app.is_filled(r, c))
                    .unwrap_or(MAX_COL)
                    .max(c1);
                app.consume_count();
                apply_operator(app, Operator::Delete, r, c1, r, c2);
                return;
            }
            // Bare `C` is change-to-end-of-row. Defer when pending_g
            // is set so the pending_g consumer's catch-all absorbs
            // `gC` as a quiet no-op (without this guard, the early
            // Nav-only block runs bare C before the pending_g
            // consumer gets a chance — undesirable mode-change for
            // an unbound prefix combo).
            KeyCode::Char('C') if !ctrl && !app.pending_g => {
                let r = app.cursor_row;
                let c1 = app.cursor_col;
                let c2 = (0..=MAX_COL)
                    .rev()
                    .find(|&c| app.is_filled(r, c))
                    .unwrap_or(MAX_COL)
                    .max(c1);
                app.consume_count();
                apply_operator(app, Operator::Change, r, c1, r, c2);
                return;
            }
            // `x` clears in place (Excel's Delete; vim's "advance" is just
            // chars shifting left, which doesn't happen in a grid). `Nx`
            // clears N cells to the right, cursor still ends at start.
            KeyCode::Char('x') if !ctrl => {
                let n = app.consume_count();
                let r = app.cursor_row;
                let start_c = app.cursor_col;
                for i in 0..n {
                    let c = start_c.saturating_add(i);
                    if c > MAX_COL {
                        break;
                    }
                    let _ = app.clear_rect(r, c, r, c);
                }
                app.last_edit = Some(EditAction {
                    kind: EditKind::Delete,
                    anchor_dr: 0,
                    anchor_dc: 0,
                    rect_rows: 1,
                    rect_cols: 1,
                    text: None,
                });
                return;
            }
            KeyCode::Char('X') if !ctrl => {
                let n = app.consume_count();
                for _ in 0..n {
                    if app.cursor_col == 0 {
                        break;
                    }
                    app.cursor_col -= 1;
                    let r = app.cursor_row;
                    let c = app.cursor_col;
                    let _ = app.clear_rect(r, c, r, c);
                }
                app.last_edit = Some(EditAction {
                    kind: EditKind::Delete,
                    anchor_dr: 0,
                    anchor_dc: 0,
                    rect_rows: 1,
                    rect_cols: 1,
                    text: None,
                });
                return;
            }
            // V6: dot-repeat and undo.
            // Bare `u` is undo; `fu` toggles underline. The pending_f
            // consumer runs before this block, so the guard is
            // belt-and-suspenders defensive.
            KeyCode::Char('u') if !ctrl && !app.pending_f => {
                app.consume_count();
                if app.undo() {
                    app.status = "Undo".into();
                } else {
                    app.status = "Nothing to undo".into();
                }
                return;
            }
            // Bare `.` is dot-repeat; `f.` is bump-decimals. The
            // pending_f consumer runs before this block, so the
            // guard is defensive.
            KeyCode::Char('.') if !ctrl && !app.pending_f => {
                app.consume_count();
                if !app.repeat_last_edit() {
                    app.status = "Nothing to repeat".into();
                }
                return;
            }
            KeyCode::Char('p' | 'P') if !ctrl => {
                app.consume_count();
                paste_from_register_or_clipboard(app);
                return;
            }
            // Bare `=` jumpstarts formula authoring: opens Insert with `=`
            // already seeded, so the user is one keystroke into typing a
            // formula. Spreadsheet convention (Excel / Sheets / sc-im).
            // Falls through when Ctrl+w prefix is pending so `Ctrl+w =`
            // (autofit fallback) reaches the prefix consumer below.
            KeyCode::Char('=') if !ctrl && !app.pending_ctrl_w => {
                app.consume_count();
                app.start_edit_blank();
                app.edit_insert('=');
                return;
            }
            // Ctrl+; → today's date as ISO literal at cursor.
            // Ctrl+Shift+; → current datetime as ISO literal.
            // Excel / Sheets convention. Some terminals deliver
            // ctrl-shift-; as Char(':') with CONTROL (since ':' is
            // shift+';'); we accept both forms.
            #[cfg(feature = "datetime")]
            KeyCode::Char(';') if ctrl && !shift => {
                app.consume_count();
                app.insert_today_literal();
                return;
            }
            #[cfg(feature = "datetime")]
            KeyCode::Char(';') if ctrl && shift => {
                app.consume_count();
                app.insert_now_literal();
                return;
            }
            #[cfg(feature = "datetime")]
            KeyCode::Char(':') if ctrl => {
                app.consume_count();
                app.insert_now_literal();
                return;
            }
            // V5: search prompts.
            KeyCode::Char('/') if !ctrl => {
                app.consume_count();
                app.start_search(SearchDir::Forward);
                return;
            }
            KeyCode::Char('?') if !ctrl => {
                app.consume_count();
                app.start_search(SearchDir::Backward);
                return;
            }
            // V5: step n/N through the active search.
            KeyCode::Char('n') if !ctrl => {
                app.consume_count();
                app.search_step(SearchDir::Forward);
                return;
            }
            KeyCode::Char('N') if !ctrl => {
                app.consume_count();
                app.search_step(SearchDir::Backward);
                return;
            }
            // V5: search for current cell's exact value.
            KeyCode::Char('*') if !ctrl => {
                app.consume_count();
                app.search_current_cell(SearchDir::Forward);
                return;
            }
            KeyCode::Char('#') if !ctrl => {
                app.consume_count();
                app.search_current_cell(SearchDir::Backward);
                return;
            }
            // V5: mark-prefix keys. `m{letter}` sets, `` `{letter} ``
            // jumps to exact, `'{letter}` jumps to row.
            KeyCode::Char('m') if !ctrl => {
                app.consume_count();
                app.pending_mark = Some(MarkAction::Set);
                return;
            }
            KeyCode::Char('`') if !ctrl => {
                app.consume_count();
                app.pending_mark = Some(MarkAction::JumpExact);
                return;
            }
            KeyCode::Char('\'') if !ctrl => {
                app.consume_count();
                app.pending_mark = Some(MarkAction::JumpRow);
                return;
            }
            _ => {}
        }
    }

    // V8 column-width prefix consumer. `Ctrl+w >` widens the active
    // column by 1; `Ctrl+w <` narrows by 1. `Ctrl+w =` autofits.
    // Bare count works: `5<C-w>>` adds 5.
    if app.pending_ctrl_w {
        app.pending_ctrl_w = false;
        let n = app.consume_count();
        let col_idx = app.cursor_col;
        let current = app.column_width(col_idx);
        match key.code {
            KeyCode::Char('>') => {
                let new = current.saturating_add(n as u16);
                let _ = app.set_column_width(col_idx, new);
            }
            KeyCode::Char('<') => {
                let new = current.saturating_sub(n as u16).max(1);
                let _ = app.set_column_width(col_idx, new);
            }
            KeyCode::Char('=') => autofit_selection_or_column(app),
            _ => {} // any other key cancels the prefix
        }
        return;
    }

    // V5: mark-pending resolution. `m{a-z}` / `'{a-z}` / `` `{a-z} `` —
    // the second key supplies the mark letter.
    if let Some(action) = app.pending_mark {
        app.pending_mark = None;
        match key.code {
            KeyCode::Char(c) if c.is_ascii_alphabetic() => match action {
                MarkAction::Set => app.set_mark(c),
                MarkAction::JumpExact => app.jump_to_mark(c, false),
                MarkAction::JumpRow => app.jump_to_mark(c, true),
            },
            _ => {} // any other key cancels the mark prefix.
        }
        return;
    }

    // V4: operator-pending resolution. Set by `d`/`c`/`y` in Normal;
    // consumed by the next motion or doubled operator key.
    if let Some(op) = app.pending_operator {
        // `dd` / `cc` / `yy` → operate on current cell.
        let doubled = match op {
            Operator::Delete => matches!(key.code, KeyCode::Char('d')),
            Operator::Change => matches!(key.code, KeyCode::Char('c')),
            Operator::Yank => matches!(key.code, KeyCode::Char('y')),
        };
        if doubled {
            let r = app.cursor_row;
            let c = app.cursor_col;
            let _trailing_count = app.pending_count.take();
            apply_operator(app, op, r, c, r, c);
            app.pending_operator = None;
            app.pending_op_count = None;
            return;
        }
        // Otherwise try to dispatch as a motion. Combine the operator's
        // captured count with any post-`d` count: `5d3w` → 15.
        let trailing = app.pending_count.take();
        let op_count = app.pending_op_count;
        let count_specified = op_count.is_some() || trailing.is_some();
        let count = op_count
            .unwrap_or(1)
            .saturating_mul(trailing.unwrap_or(1))
            .max(1);
        match motion_target_for_operator(app, key, count, count_specified) {
            Some((tr, tc)) => {
                let r1 = app.cursor_row.min(tr);
                let c1 = app.cursor_col.min(tc);
                let r2 = app.cursor_row.max(tr);
                let c2 = app.cursor_col.max(tc);
                apply_operator(app, op, r1, c1, r2, c2);
            }
            None => {
                app.status = "Cancelled".into();
            }
        }
        app.pending_operator = None;
        app.pending_op_count = None;
        return;
    }

    // `g`-prefix consumer: handles `gg`, `gv` (V3), and `gt`/`gT` (V7).
    if app.pending_g {
        app.pending_g = false;
        match key.code {
            KeyCode::Char('g') => match app.pending_count.take() {
                Some(n) => app.goto_row(n.saturating_sub(1)),
                None => app.goto_first_row(),
            },
            // Vim `gv`: re-enter the most-recent Visual selection.
            KeyCode::Char('v') => {
                app.pending_count = None;
                if !app.reselect_last_visual() {
                    app.status = "No previous selection".into();
                }
            }
            // Vim `gt` / `{N}gt`: next sheet, or jump to absolute sheet N.
            KeyCode::Char('t') => match app.pending_count.take() {
                Some(n) => app.switch_sheet(n.saturating_sub(1) as usize),
                None => app.next_sheet(),
            },
            // Vim `gT` (capital): previous sheet. Count is ignored on `gT`
            // in vim too — explicit absolute jump uses `gt`.
            KeyCode::Char('T') => {
                app.pending_count = None;
                app.prev_sheet();
            }
            // `go`: open the URL under the cursor in the system
            // browser. Plain-text URL cells and HYPERLINK custom
            // values both route through `App::url_under_cursor`.
            // (Mnemonic: "go [to link]". Vim's `gx` is the
            // close cousin; we use `go` for the better mnemonic.)
            KeyCode::Char('o') => {
                app.pending_count = None;
                match app.url_under_cursor() {
                    Some(url) => match hyperlink::open_url(&url) {
                        Ok(()) => app.status = format!("Opened {url}"),
                        Err(e) => app.status = format!("Open failed: {e}"),
                    },
                    None => app.status = "No link under cursor".into(),
                }
            }
            _ => app.pending_count = None,
        }
        return;
    }

    // `z`-prefix consumer: zz/zt/zb/zh/zl.
    if app.pending_z {
        app.pending_z = false;
        // The z-prefix scrolls don't take a count; drop any accumulated.
        app.pending_count = None;
        match key.code {
            KeyCode::Char('z') => app.scroll_cursor_to_middle(),
            KeyCode::Char('t') => app.scroll_cursor_to_top(),
            KeyCode::Char('b') => app.scroll_cursor_to_bottom(),
            KeyCode::Char('h') => app.scroll_viewport_left(),
            KeyCode::Char('l') => app.scroll_viewport_right(),
            _ => {}
        }
        return;
    }

    // Digit accumulator. `0` is the "go to col 0" motion *unless* a count
    // is already in progress, in which case it's a digit.
    if let KeyCode::Char(d @ '0'..='9') = key.code {
        let is_motion_zero = d == '0' && app.pending_count.is_none();
        if !ctrl && !is_motion_zero {
            let n = d.to_digit(10).unwrap();
            let prev = app.pending_count.unwrap_or(0);
            app.pending_count = Some(prev.saturating_mul(10).saturating_add(n));
            return;
        }
    }

    // Ctrl+combinations first so they don't get caught by the bare-key arms.
    if ctrl {
        match key.code {
            // [sheet.clipboard.copy]
            KeyCode::Char('c') => {
                copy_selection_to_clipboard(app, ClipMarkMode::Copy);
                return;
            }
            // [sheet.clipboard.cut]
            KeyCode::Char('x') => {
                copy_selection_to_clipboard(app, ClipMarkMode::Cut);
                return;
            }
            // [sheet.clipboard.paste]
            KeyCode::Char('v') => {
                paste_from_clipboard(app);
                return;
            }
            // [sheet.undo.cmd-z]
            // [sheet.undo.redo]
            // Ctrl+Shift+Z is redo (matches the spec's preferred binding);
            // bare Ctrl+Z is undo. Ctrl+Y is the alternate redo.
            KeyCode::Char('z') if shift => {
                if app.redo() {
                    app.status = "Redo".into();
                } else {
                    app.status = "Nothing to redo".into();
                }
                return;
            }
            KeyCode::Char('z') => {
                if app.undo() {
                    app.status = "Undo".into();
                } else {
                    app.status = "Nothing to undo".into();
                }
                return;
            }
            // [sheet.undo.redo]
            KeyCode::Char('y') => {
                if app.redo() {
                    app.status = "Redo".into();
                } else {
                    app.status = "Nothing to redo".into();
                }
                return;
            }
            // V6 vim Ctrl+r: redo (alias of Ctrl+Shift+Z / Ctrl+Y).
            KeyCode::Char('r') => {
                if app.redo() {
                    app.status = "Redo".into();
                } else {
                    app.status = "Nothing to redo".into();
                }
                return;
            }
            // [sheet.tabs.keyboard-switch]
            KeyCode::PageDown => {
                app.next_sheet();
                return;
            }
            KeyCode::PageUp => {
                app.prev_sheet();
                return;
            }
            // V8 column-width: Ctrl+w as a prefix (vim's window-resize key).
            // Next key resolves to `>` / `<` (grow / shrink current column).
            KeyCode::Char('w') => {
                app.pending_ctrl_w = true;
                return;
            }
            // vim V2 Ctrl+d / Ctrl+u / Ctrl+f / Ctrl+b. Count repeats the
            // jump (Vim's behavior — count is *number of pages*).
            KeyCode::Char('d') => {
                let n = app.consume_count();
                for _ in 0..n {
                    app.scroll_half_page(1);
                }
                return;
            }
            KeyCode::Char('u') => {
                let n = app.consume_count();
                for _ in 0..n {
                    app.scroll_half_page(-1);
                }
                return;
            }
            KeyCode::Char('f') => {
                let n = app.consume_count();
                for _ in 0..n {
                    app.scroll_full_page(1);
                }
                return;
            }
            KeyCode::Char('b') => {
                let n = app.consume_count();
                for _ in 0..n {
                    app.scroll_full_page(-1);
                }
                return;
            }
            _ => {}
        }
    }

    match key.code {
        // [sheet.navigation.shift-arrow-jump-extend]
        KeyCode::Left if shift && ctrl => app.move_cursor_jump_selecting(0, -1),
        KeyCode::Down if shift && ctrl => app.move_cursor_jump_selecting(1, 0),
        KeyCode::Up if shift && ctrl => app.move_cursor_jump_selecting(-1, 0),
        KeyCode::Right if shift && ctrl => app.move_cursor_jump_selecting(0, 1),
        // [sheet.navigation.arrow-jump]
        KeyCode::Left if ctrl => app.move_cursor_jump(0, -1),
        KeyCode::Down if ctrl => app.move_cursor_jump(1, 0),
        KeyCode::Up if ctrl => app.move_cursor_jump(-1, 0),
        KeyCode::Right if ctrl => app.move_cursor_jump(0, 1),
        // [sheet.navigation.shift-arrow-extend]
        KeyCode::Left if shift => app.move_cursor_selecting(0, -1),
        KeyCode::Down if shift => app.move_cursor_selecting(1, 0),
        KeyCode::Up if shift => app.move_cursor_selecting(-1, 0),
        KeyCode::Right if shift => app.move_cursor_selecting(0, 1),
        // [sheet.navigation.arrow]
        KeyCode::Left => app.move_cursor(0, -1),
        KeyCode::Down => app.move_cursor(1, 0),
        KeyCode::Up => app.move_cursor(-1, 0),
        KeyCode::Right => app.move_cursor(0, 1),
        // [sheet.navigation.tab-nav-move]
        KeyCode::BackTab => app.move_cursor(0, -1),
        KeyCode::Tab if shift => app.move_cursor(0, -1),
        KeyCode::Tab => app.move_cursor(0, 1),
        // [sheet.editing.f2-or-enter]
        KeyCode::Enter | KeyCode::F(2) => {
            app.consume_count();
            app.start_edit();
        }
        // [sheet.delete.delete-key-clears]
        KeyCode::Delete | KeyCode::Backspace => {
            app.consume_count();
            app.delete_cell();
        }
        // [sheet.clipboard.escape-cancels-mark] handled in the early-return
        // Esc block above so it can also clear pending vim-prefix flags.
        KeyCode::Char(':') => {
            app.consume_count();
            app.start_command();
        }
        // `!` opens the shell prompt — runs a subprocess, sniffs the
        // stdout (JSON / TSV / CSV / plain), and pastes the result at
        // the cursor with the first row as headers.
        KeyCode::Char('!') => {
            app.consume_count();
            app.start_shell();
        }
        // vim V1/V2: hjkl motions (count-multiplied; arrows above don't take a count).
        KeyCode::Char('h') if !ctrl => {
            let n = app.consume_count();
            for _ in 0..n {
                app.move_cursor(0, -1);
            }
        }
        KeyCode::Char('j') if !ctrl => {
            let n = app.consume_count();
            for _ in 0..n {
                app.move_cursor(1, 0);
            }
        }
        KeyCode::Char('k') if !ctrl => {
            let n = app.consume_count();
            for _ in 0..n {
                app.move_cursor(-1, 0);
            }
        }
        // Bare `l` moves right; `fl` left-aligns (defensive guard —
        // pending_f consumer runs first).
        KeyCode::Char('l') if !ctrl && !app.pending_f => {
            let n = app.consume_count();
            for _ in 0..n {
                app.move_cursor(0, 1);
            }
        }
        // vim V2: word motions.
        KeyCode::Char('w') if !ctrl => {
            let n = app.consume_count();
            for _ in 0..n {
                app.word_forward();
            }
        }
        // Bare `b` is word-backward; `fb` toggles bold. Guard is
        // defensive — pending_f consumer runs before this block.
        KeyCode::Char('b') if !ctrl && !app.pending_f => {
            let n = app.consume_count();
            for _ in 0..n {
                app.word_backward();
            }
        }
        KeyCode::Char('e') if !ctrl => {
            let n = app.consume_count();
            for _ in 0..n {
                app.word_end();
            }
        }
        // vim V2: row-internal absolute motions.
        KeyCode::Char('0') if !ctrl => {
            // Reaches here only when no count is pending — see the digit
            // accumulator above.
            app.move_to_row_start();
        }
        KeyCode::Char('^') if !ctrl => {
            app.consume_count();
            app.move_to_first_filled_in_row();
        }
        KeyCode::Char('$') if !ctrl => {
            app.consume_count();
            app.move_to_last_filled_in_row();
        }
        // vim V2: paragraph (between filled runs in current column).
        KeyCode::Char('{') if !ctrl => {
            let n = app.consume_count();
            for _ in 0..n {
                app.paragraph_backward();
            }
        }
        KeyCode::Char('}') if !ctrl => {
            let n = app.consume_count();
            for _ in 0..n {
                app.paragraph_forward();
            }
        }
        // vim V2: viewport.
        KeyCode::Char('H') if !ctrl => {
            app.consume_count();
            app.move_to_viewport_top();
        }
        KeyCode::Char('M') if !ctrl => {
            app.consume_count();
            app.move_to_viewport_middle();
        }
        // Bare `L` is viewport-bottom. Defer when pending_g is set
        // so `gL` cleanly hits the pending_g consumer's catch-all
        // instead of running viewport-bottom on an unbound prefix
        // combo. (Same rationale as bare `C` above.)
        KeyCode::Char('L') if !ctrl && !app.pending_g => {
            app.consume_count();
            app.move_to_viewport_bottom();
        }
        // vim V2: g-prefix (gg) and z-prefix (zz/zt/zb/zh/zl). Don't
        // consume the pending count — the second key resolves it.
        KeyCode::Char('g') if !ctrl => app.pending_g = true,
        KeyCode::Char('z') if !ctrl => app.pending_z = true,
        // `f`-prefix (format) — second key resolves a format axis.
        // Bare `f` is unbound in Normal otherwise; vlotus does not
        // implement vim's `f{char}` row-internal char-find.
        KeyCode::Char('f') if !ctrl => app.pending_f = true,
        // vim V2: G — bare goes to the last filled row of the column;
        // with a count, jumps to that 1-indexed row.
        KeyCode::Char('G') if !ctrl => match app.pending_count.take() {
            Some(n) => app.goto_row(n.saturating_sub(1)),
            None => app.goto_last_filled_row(),
        },
        // vim V1: enter Insert. `i`/`I` put the caret at the start of the
        // existing value; `a`/`A` at the end. Cells are single-line so
        // capital and lowercase forms are aliases. `fi` toggles italic
        // — guard is defensive (pending_f consumer runs first; `fI`
        // falls through the consumer's catch-all and is a no-op).
        KeyCode::Char('i' | 'I') if !ctrl && !app.pending_f => {
            app.consume_count();
            app.start_edit_at_start();
        }
        // Bare `a`/`A` enter Insert at end-of-cell; `fa` clears
        // alignment (defensive guard — pending_f consumer runs first).
        KeyCode::Char('a' | 'A') if !ctrl && !app.pending_f => {
            app.consume_count();
            app.start_edit();
        }
        // vim V1: open below / above with an empty buffer. Defer when
        // pending_g is set so `go` (open URL) reaches the pending_g
        // consumer above — same pattern as bare `C`.
        KeyCode::Char('o') if !ctrl && !app.pending_g => {
            app.consume_count();
            app.move_cursor(1, 0);
            app.start_edit_blank();
        }
        KeyCode::Char('O') if !ctrl => {
            app.consume_count();
            app.move_cursor(-1, 0);
            app.start_edit_blank();
        }
        // vim V1: substitute — clear cell and Insert. `cc` waits for V4's
        // operator state machine. `fs` toggles strikethrough — guard is
        // defensive (pending_f consumer runs first).
        KeyCode::Char('s' | 'S') if !ctrl && !app.pending_f => {
            app.consume_count();
            app.start_edit_blank();
        }
        _ => {
            // Drop any pending count for unhandled keys so it doesn't
            // bleed into the next command.
            app.consume_count();
        }
    }
}

// [sheet.clipboard.copy]
// [sheet.clipboard.cut]
fn copy_selection_to_clipboard(app: &mut App, mode: ClipMarkMode) {
    let tsv = app.yank_tsv();
    let html = app.yank_html();
    let result = Clipboard::new().and_then(|mut cb| cb.set().html(&html, Some(&tsv)));
    match result {
        Ok(()) => {
            app.set_clipboard_mark(mode);
            let verb = match mode {
                ClipMarkMode::Copy => "Copied",
                ClipMarkMode::Cut => "Cut",
            };
            let (r1, c1, r2, c2) = app.effective_rect();
            let rows = r2 - r1 + 1;
            let cols = c2 - c1 + 1;
            app.status = if rows == 1 && cols == 1 {
                format!("{verb} cell")
            } else {
                format!("{verb} {rows}x{cols} range")
            };
        }
        Err(e) => app.status = format!("Clipboard error: {e}"),
    }
}

// [sheet.clipboard.paste]
// [sheet.clipboard.paste-formula-shift]
// [sheet.clipboard.paste-fill-selection]
fn paste_from_clipboard(app: &mut App) {
    let cut_source = app
        .clipboard_mark
        .filter(|m| m.mode == ClipMarkMode::Cut)
        .map(|m| (m.r1, m.c1, m.r2, m.c2));

    match Clipboard::new() {
        Ok(mut cb) => {
            let html = cb.get().html().ok().filter(|h| h.contains("<t"));
            if let Some(grid) = html.as_deref().and_then(app::parse_pasted_grid) {
                match app.apply_pasted_grid(&grid, cut_source) {
                    Ok((rows, cols)) => {
                        app.status = format!("Pasted {rows}x{cols} range");
                        if cut_source.is_some() {
                            // [sheet.clipboard.cut] cut consumed.
                            app.clipboard_mark = None;
                        }
                    }
                    Err(e) => app.status = format!("Paste error: {e}"),
                }
            } else if let Ok(text) = cb.get().text() {
                if !text.is_empty() {
                    let grid = PastedGrid {
                        source_anchor: None,
                        cells: vec![vec![PastedCell {
                            value: text,
                            formula: None,
                        }]],
                    };
                    match app.apply_pasted_grid(&grid, cut_source) {
                        Ok((rows, cols)) => {
                            app.status = if rows == 1 && cols == 1 {
                                "Pasted".into()
                            } else {
                                format!("Pasted into {rows}x{cols} selection")
                            };
                            if cut_source.is_some() {
                                app.clipboard_mark = None;
                            }
                        }
                        Err(e) => app.status = format!("Paste error: {e}"),
                    }
                }
            }
        }
        Err(e) => app.status = format!("Clipboard error: {e}"),
    }
}

// ── Vim yank / paste (V3) ────────────────────────────────────────────

/// Vim Visual `y`: store the selection in the unnamed register, mirror
/// it onto the OS clipboard (so external apps see the same payload),
/// then drop back to Normal mode.
fn visual_yank(app: &mut App) {
    let linewise = matches!(app.mode, Mode::Visual(VisualKind::Row));
    let (r1, c1, r2, c2) = app.yank_selection(linewise);
    sync_register_to_os_clipboard(app);
    let rows = r2 - r1 + 1;
    let cols = c2 - c1 + 1;
    app.exit_visual();
    app.status = if rows == 1 && cols == 1 {
        "Yanked cell".into()
    } else {
        format!("Yanked {rows}x{cols}")
    };
}

/// Vim Visual `d` / `x`: store selection in the register and clear every
/// cell in the rectangle. Returns to Normal.
fn visual_delete(app: &mut App) {
    let linewise = matches!(app.mode, Mode::Visual(VisualKind::Row));
    let (r1, c1, r2, c2) = app.yank_selection(linewise);
    sync_register_to_os_clipboard(app);
    let result = app.clear_rect(r1, c1, r2, c2);
    let rows = r2 - r1 + 1;
    let cols = c2 - c1 + 1;
    app.exit_visual();
    match result {
        Ok(()) => {
            app.status = if rows == 1 && cols == 1 {
                "Deleted cell".into()
            } else {
                format!("Deleted {rows}x{cols}")
            };
        }
        Err(e) => app.status = format!("Error: {e}"),
    }
    app.cursor_row = r1;
    app.cursor_col = c1;
}

/// Vim Visual `c`: clear every cell in the selection and drop into Insert
/// at the rectangle's top-left.
fn visual_change(app: &mut App) {
    let linewise = matches!(app.mode, Mode::Visual(VisualKind::Row));
    let (r1, c1, _r2, _c2) = app.yank_selection(linewise);
    sync_register_to_os_clipboard(app);
    let _ = app.clear_rect(r1, c1, _r2, _c2);
    app.exit_visual();
    app.cursor_row = r1;
    app.cursor_col = c1;
    app.start_edit_blank();
}

/// Vim `yy` / `Y`: yank current row into the register and OS clipboard.
/// Kept around for `:tabs`-style helpers; V4 routes `Y` through
/// `apply_operator` instead.
#[allow(dead_code)]
fn yank_row_to_register(app: &mut App) {
    app.yank_row();
    sync_register_to_os_clipboard(app);
    app.status = "Yanked row".into();
}

/// Vim `p` / `P`: paste from the unnamed register at the cursor; falls
/// back to the OS clipboard when the register is empty.
fn paste_from_register_or_clipboard(app: &mut App) {
    if app.yank_register.is_some() {
        match app.paste_from_register() {
            Ok((rows, cols)) => {
                app.status = if rows == 1 && cols == 1 {
                    "Pasted".into()
                } else {
                    format!("Pasted {rows}x{cols}")
                };
            }
            Err(e) => app.status = format!("Paste error: {e}"),
        }
        return;
    }
    paste_from_clipboard(app);
}

/// Mirror the unnamed register onto the OS clipboard so external apps
/// see the same payload. Reuses the HTML+TSV producers `Ctrl+C` already
/// uses — the cursor needs to point at the yanked region for that to
/// produce the right rect, which it does (yank_* sets the source anchor
/// to the top-left and the cursor sits inside the selection).
fn sync_register_to_os_clipboard(app: &mut App) {
    // The existing yank_html/yank_tsv functions already key off
    // `effective_rect()`, which derives from the live selection. Since
    // visual_* call this *before* exit_visual, the selection is still
    // intact at this point.
    let html = app.yank_html();
    let tsv = app.yank_tsv();
    let _ = Clipboard::new().and_then(|mut cb| cb.set().html(&html, Some(&tsv)));
}

/// Mirror a specific rect onto the OS clipboard. Used by V4 operator
/// yanks where there's no live selection — temporarily hijacks the
/// selection_anchor so yank_html/yank_tsv produce the right region,
/// then restores.
fn sync_rect_to_os_clipboard(app: &mut App, r1: u32, c1: u32, r2: u32, c2: u32) {
    let saved_anchor = app.selection_anchor;
    let saved_row = app.cursor_row;
    let saved_col = app.cursor_col;
    app.selection_anchor = Some((r1, c1));
    app.cursor_row = r2;
    app.cursor_col = c2;
    let html = app.yank_html();
    let tsv = app.yank_tsv();
    app.selection_anchor = saved_anchor;
    app.cursor_row = saved_row;
    app.cursor_col = saved_col;
    let _ = Clipboard::new().and_then(|mut cb| cb.set().html(&html, Some(&tsv)));
}

// ── Vim operator-pending resolution (V4) ────────────────────────────

/// Compute the cell a single motion key would land on, without mutating
/// the cursor. Used by the V4 operator-pending state to build a rect for
/// `d{motion}` / `c{motion}` / `y{motion}`.
///
/// Returns `None` for keys that aren't motions — the caller cancels the
/// pending operator in that case (vim's behavior).
fn motion_target_for_operator(
    app: &App,
    key: event::KeyEvent,
    count: u32,
    count_specified: bool,
) -> Option<(u32, u32)> {
    let ctrl = key.modifiers.contains(KeyModifiers::CONTROL);
    let row = app.cursor_row;
    let col = app.cursor_col;
    let n = count.max(1) as i32;

    if ctrl {
        let visible = app.visible_rows.max(1) as i32;
        let half = (app.visible_rows / 2).max(1) as i32;
        return match key.code {
            KeyCode::Char('d') => Some(((row as i32 + n * half).clamp(0, MAX_ROW as i32) as u32, col)),
            KeyCode::Char('u') => Some(((row as i32 - n * half).max(0) as u32, col)),
            KeyCode::Char('f') => Some(((row as i32 + n * visible).clamp(0, MAX_ROW as i32) as u32, col)),
            KeyCode::Char('b') => Some(((row as i32 - n * visible).max(0) as u32, col)),
            _ => None,
        };
    }

    match key.code {
        KeyCode::Char('h') => Some((row, (col as i32 - n).max(0) as u32)),
        KeyCode::Char('j') => Some(((row + count).min(MAX_ROW), col)),
        KeyCode::Char('k') => Some(((row as i32 - n).max(0) as u32, col)),
        KeyCode::Char('l') => Some((row, (col + count).min(MAX_COL))),
        KeyCode::Char('w') => {
            let mut r = row;
            let mut c = col;
            for _ in 0..count {
                let (nr, nc) = compute_word_forward(r, c, MAX_COL, |row, col| app.is_filled(row, col));
                r = nr;
                c = nc;
            }
            Some((r, c))
        }
        KeyCode::Char('b') => {
            let mut r = row;
            let mut c = col;
            for _ in 0..count {
                let (nr, nc) = compute_word_backward(r, c, |row, col| app.is_filled(row, col));
                r = nr;
                c = nc;
            }
            Some((r, c))
        }
        KeyCode::Char('e') => {
            let mut r = row;
            let mut c = col;
            for _ in 0..count {
                let (nr, nc) = compute_jump_target(r, c, 0, 1, MAX_ROW, MAX_COL, |row, col| {
                    app.is_filled(row, col)
                });
                r = nr;
                c = nc;
            }
            Some((r, c))
        }
        KeyCode::Char('0') => Some((row, 0)),
        KeyCode::Char('^') => Some((
            row,
            (0..=MAX_COL).find(|&c| app.is_filled(row, c)).unwrap_or(0),
        )),
        KeyCode::Char('$') => Some((
            row,
            (0..=MAX_COL)
                .rev()
                .find(|&c| app.is_filled(row, c))
                .unwrap_or(MAX_COL),
        )),
        KeyCode::Char('G') => {
            if count_specified {
                Some(((count.saturating_sub(1)).min(MAX_ROW), col))
            } else {
                // Global last-row-with-data, mirroring App::goto_last_filled_row.
                let last = app.cells.iter().map(|c| c.row_idx).max().unwrap_or(0);
                Some((last, col))
            }
        }
        KeyCode::Char('H') => Some((app.scroll_row, col)),
        KeyCode::Char('M') => Some((app.scroll_row + app.visible_rows / 2, col)),
        KeyCode::Char('L') => Some((
            app.scroll_row + app.visible_rows.saturating_sub(1),
            col,
        )),
        _ => None,
    }
}

/// Apply the resolved operator over the given inclusive rect. Capture
/// into the unnamed register first (so `dw` is yank+clear in one undo
/// step from the user's POV — the apply_changes_recorded inside
/// clear_rect handles the actual undo entry).
fn apply_operator(app: &mut App, op: Operator, r1: u32, c1: u32, r2: u32, c2: u32) {
    let grid = app.build_grid_over(r1, c1, r2, c2);
    app.yank_register = Some(YankRegister {
        grid,
        linewise: false,
    });
    sync_rect_to_os_clipboard(app, r1, c1, r2, c2);

    let rows = r2 - r1 + 1;
    let cols = c2 - c1 + 1;
    let anchor_dr = r1 as i32 - app.cursor_row as i32;
    let anchor_dc = c1 as i32 - app.cursor_col as i32;

    match op {
        Operator::Yank => {
            app.status = if rows == 1 && cols == 1 {
                "Yanked cell".into()
            } else {
                format!("Yanked {rows}x{cols}")
            };
        }
        Operator::Delete => match app.clear_rect(r1, c1, r2, c2) {
            Ok(()) => {
                app.last_edit = Some(EditAction {
                    kind: EditKind::Delete,
                    anchor_dr,
                    anchor_dc,
                    rect_rows: rows,
                    rect_cols: cols,
                    text: None,
                });
                app.cursor_row = r1;
                app.cursor_col = c1;
                app.status = if rows == 1 && cols == 1 {
                    "Deleted cell".into()
                } else {
                    format!("Deleted {rows}x{cols}")
                };
            }
            Err(e) => app.status = format!("Error: {e}"),
        },
        Operator::Change => {
            let _ = app.clear_rect(r1, c1, r2, c2);
            // Record the rect now; confirm_edit will fill in the text on
            // commit (or cancel_edit will downgrade to Delete on Esc).
            app.last_edit = Some(EditAction {
                kind: EditKind::Change,
                anchor_dr,
                anchor_dc,
                rect_rows: rows,
                rect_cols: cols,
                text: None,
            });
            app.cursor_row = r1;
            app.cursor_col = c1;
            app.start_edit_blank();
        }
    }
}

#[cfg(test)]
mod mouse_tests {
    //! Click-handler tests for the T1 mouse foundation. Hit-testing is
    //! covered separately in `ui::tests` (the `cell_at` helper).

    use super::*;
    use crossterm::event::{KeyModifiers as Mods, MouseEvent, MouseEventKind};
    use ratatui::{backend::TestBackend, layout::Rect, Terminal};

    fn make_app() -> App {
        let store = Store::open_in_memory().unwrap();
        App::new(store, "test")
    }

    /// Render once so `visible_rows` / `visible_cols` are populated, then
    /// return the area used for the render.
    fn render(app: &mut App, w: u16, h: u16) -> Rect {
        let backend = TestBackend::new(w, h);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| ui::draw(f, app)).unwrap();
        Rect::new(0, 0, w, h)
    }

    fn click(x: u16, y: u16) -> MouseEvent {
        MouseEvent {
            kind: MouseEventKind::Down(MouseButton::Left),
            column: x,
            row: y,
            modifiers: Mods::NONE,
        }
    }

    #[test]
    fn left_click_in_nav_moves_cursor() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // x=10, y=5 → row 0, col 0 (per ui::tests::cell_at_resolves_grid_interior_to_cell).
        // Move starting cursor away from origin to make the assertion meaningful.
        app.cursor_row = 4;
        app.cursor_col = 3;
        handle_mouse(&mut app, click(10, 5), area);
        assert_eq!(app.cursor_row, 0);
        assert_eq!(app.cursor_col, 0);
    }

    #[test]
    fn left_click_clears_pending_motion_state() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.pending_count = Some(5);
        app.pending_g = true;
        handle_mouse(&mut app, click(10, 5), area);
        assert_eq!(app.pending_count, None);
        assert!(!app.pending_g);
    }

    // ── T6: row/column-header click → row/column visual ───────────────

    #[test]
    fn left_click_on_row_header_enters_v_line() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // x=2 (gutter), y=6 (row 1) → RowHeader(1).
        handle_mouse(&mut app, click(2, 6), area);
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Row)));
        assert_eq!(app.cursor_row, 1);
        assert_eq!(app.selection_anchor, Some((1, 0)));
    }

    #[test]
    fn left_click_on_column_header_enters_v_column() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // x=10, y=4 → ColumnHeader(0).
        handle_mouse(&mut app, click(10, 4), area);
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Column)));
        // V-COLUMN forces row extent to full sheet via selection_range, so
        // the anchor's row doesn't matter for highlighting; cursor stays
        // on whatever row was active so the viewport doesn't scroll.
        let (r1, c1, r2, c2) = app.selection_range().unwrap();
        assert_eq!(c1, 0);
        assert_eq!(c2, 0);
        assert_eq!(r1, 0);
        assert_eq!(r2, MAX_ROW);
        assert_eq!(app.cursor_col, 0);
    }

    #[test]
    fn drag_along_row_headers_extends_v_line_selection() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Click row 2 header, then drag down to row 5 header.
        handle_mouse(&mut app, click(2, 7), area); // y=7 → RowHeader(2)
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Row)));
        assert_eq!(app.cursor_row, 2);
        handle_mouse(&mut app, drag(2, 10), area); // y=10 → RowHeader(5)
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Row)));
        assert_eq!(app.cursor_row, 5);
        assert_eq!(app.selection_anchor, Some((2, 0)));
    }

    #[test]
    fn drag_along_column_headers_extends_column_selection() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(10, 4), area); // ColumnHeader(0)
        let cursor_row_after_down = app.cursor_row;
        handle_mouse(&mut app, drag(35, 4), area); // ColumnHeader(2) approximately
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Column)));
        // Cursor row stays put (column-axis drag preserves row).
        assert_eq!(app.cursor_row, cursor_row_after_down);
        // Column moved to the dragged-to header.
        assert_ne!(app.cursor_col, 0);
        // selection_range forces full-sheet rows; col extent covers 0..=cursor_col.
        let (r1, c1, r2, c2) = app.selection_range().unwrap();
        assert_eq!(r1, 0);
        assert_eq!(r2, MAX_ROW);
        assert_eq!(c1, 0);
        assert_eq!(c2, app.cursor_col);
    }

    #[test]
    fn left_click_on_tab_switches_active_sheet() {
        // T5: clicking a tab in the tabline switches the active sheet.
        let mut app = make_app();
        app.add_sheet("Sheet2").unwrap();
        // `add_sheet` activates the new sheet; switch back to Sheet1 so the
        // click moves us forward.
        app.switch_sheet(0);
        assert_eq!(app.active_sheet, 0);
        let area = render(&mut app, 80, 24);
        // Tab 0 (" test ") is at x=0..6; tab 1 (" Sheet2 ") follows. Click
        // x=10 to land on the second tab.
        handle_mouse(&mut app, click(10, 22), area);
        assert_eq!(app.active_sheet, 1);
    }

    #[test]
    fn left_click_on_active_tab_is_idempotent() {
        // Clicking the already-active tab shouldn't break anything; it's
        // a no-op (App::switch_sheet bails when idx == active_sheet).
        let mut app = make_app();
        app.add_sheet("Sheet2").unwrap();
        app.switch_sheet(0);
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(2, 22), area); // tab 0 (active)
        assert_eq!(app.active_sheet, 0);
    }

    #[test]
    fn left_click_outside_grid_is_noop() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        let (r0, c0) = (app.cursor_row, app.cursor_col);
        // y=1 is in the formula bar.
        handle_mouse(&mut app, click(10, 1), area);
        assert_eq!(app.cursor_row, r0);
        assert_eq!(app.cursor_col, c0);
    }

    #[test]
    fn left_click_in_edit_mode_commits_and_moves() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Type "hi" into A1 in Edit mode (without committing).
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('h');
        app.edit_insert('i');
        assert_eq!(app.mode, Mode::Edit);
        // Click on a different cell — should commit "hi" to A1 and move there.
        // x=10 + 13 = 23 lands inside col 1 (12-wide col + 1 spacing); y=6 = row 1.
        handle_mouse(&mut app, click(23, 6), area);
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.cursor_row, 1);
        assert_eq!(app.cursor_col, 1);
        assert_eq!(app.get_raw(0, 0), "hi");
    }

    #[test]
    fn left_click_on_self_in_edit_mode_is_noop() {
        // Clicking the cell you're already editing shouldn't commit you
        // out of edit mode — that would be infuriating mid-typing.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('h');
        // Click x=10, y=5 → still A1 (where the cursor already is).
        handle_mouse(&mut app, click(10, 5), area);
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.edit_buf, "h");
    }

    #[test]
    fn non_left_button_is_ignored() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        let (r0, c0) = (app.cursor_row, app.cursor_col);
        let me = MouseEvent {
            kind: MouseEventKind::Down(MouseButton::Right),
            column: 10,
            row: 5,
            modifiers: Mods::NONE,
        };
        handle_mouse(&mut app, me, area);
        assert_eq!(app.cursor_row, r0);
        assert_eq!(app.cursor_col, c0);
    }

    fn wheel(kind: MouseEventKind, modifiers: Mods) -> MouseEvent {
        MouseEvent {
            kind,
            column: 10,
            row: 5,
            modifiers,
        }
    }

    // ── T3: scroll-wheel → grid scroll ────────────────────────────────

    #[test]
    fn scroll_down_advances_viewport_without_moving_cursor() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        let (r0, c0) = (app.cursor_row, app.cursor_col);
        let pre_scroll = app.scroll_row;
        handle_mouse(&mut app, wheel(MouseEventKind::ScrollDown, Mods::NONE), area);
        assert_eq!(app.scroll_row, pre_scroll + 1);
        assert_eq!(app.cursor_row, r0);
        assert_eq!(app.cursor_col, c0);
    }

    #[test]
    fn scroll_up_at_origin_clamps_to_zero() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        assert_eq!(app.scroll_row, 0);
        handle_mouse(&mut app, wheel(MouseEventKind::ScrollUp, Mods::NONE), area);
        assert_eq!(app.scroll_row, 0);
    }

    #[test]
    fn shift_scroll_advances_horizontal_viewport() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        let pre = app.scroll_col;
        handle_mouse(&mut app, wheel(MouseEventKind::ScrollDown, Mods::SHIFT), area);
        assert_eq!(app.scroll_col, pre + 1);
        // scroll_row should not have changed.
        assert_eq!(app.scroll_row, 0);
    }

    #[test]
    fn scroll_right_advances_horizontal_viewport() {
        // Native ScrollRight (no Shift modifier needed on terminals that
        // emit it).
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        let pre = app.scroll_col;
        handle_mouse(
            &mut app,
            wheel(MouseEventKind::ScrollRight, Mods::NONE),
            area,
        );
        assert_eq!(app.scroll_col, pre + 1);
    }

    fn drag(x: u16, y: u16) -> MouseEvent {
        MouseEvent {
            kind: MouseEventKind::Drag(MouseButton::Left),
            column: x,
            row: y,
            modifiers: Mods::NONE,
        }
    }

    fn release(x: u16, y: u16) -> MouseEvent {
        MouseEvent {
            kind: MouseEventKind::Up(MouseButton::Left),
            column: x,
            row: y,
            modifiers: Mods::NONE,
        }
    }

    // ── T2: drag → cell visual selection ──────────────────────────────

    #[test]
    fn down_records_drag_anchor() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // x=10, y=5 → A1 (row 0, col 0).
        handle_mouse(&mut app, click(10, 5), area);
        assert_eq!(app.drag_anchor, Some((0, 0)));
    }

    #[test]
    fn up_clears_drag_anchor() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(10, 5), area);
        handle_mouse(&mut app, release(10, 5), area);
        assert_eq!(app.drag_anchor, None);
    }

    #[test]
    fn drag_in_nav_transitions_to_visual_cell_with_correct_anchor() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Down at A1, drag to B1. Default cols are 12-wide; col 0 starts
        // at x=7, col 1 starts at x=20. y=5 → row 0; y=6 → row 1.
        handle_mouse(&mut app, click(10, 5), area);
        handle_mouse(&mut app, drag(23, 6), area);
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
        assert_eq!(app.selection_anchor, Some((0, 0)));
        assert_eq!(app.cursor_row, 1);
        assert_eq!(app.cursor_col, 1);
    }

    #[test]
    fn drag_without_movement_stays_in_nav() {
        // A Drag event arrives at the same cell as Down — some terminals
        // emit one on press without motion. No transition.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(10, 5), area);
        handle_mouse(&mut app, drag(11, 5), area); // same cell, slightly different x
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.selection_anchor, None);
    }

    #[test]
    fn drag_without_prior_down_is_noop() {
        // Stray drag event arriving without a remembered click anchor
        // (e.g. `:set mouse` flipped on mid-drag). Don't enter Visual.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        let (r0, c0) = (app.cursor_row, app.cursor_col);
        handle_mouse(&mut app, drag(23, 6), area);
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.cursor_row, r0);
        assert_eq!(app.cursor_col, c0);
    }

    #[test]
    fn click_in_visual_resets_selection_to_just_clicked_cell() {
        // Excel / Sheets semantics: after a drag-selection, a plain click
        // on a different cell discards the selection and lands the cursor
        // on the clicked cell. Without this, the selection would extend
        // from the original anchor to the click point — surprising users
        // who expect the click to start over.
        let mut app = make_app();
        app.mode = Mode::Visual(VisualKind::Cell);
        app.selection_anchor = Some((0, 0));
        app.cursor_row = 3;
        app.cursor_col = 1;
        let area = render(&mut app, 80, 24);
        // Click on a third cell (col 2-ish, row 4 — past the existing
        // A1:B4 selection).
        handle_mouse(&mut app, click(35, 9), area);
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.selection_anchor, None);
        // Cursor moved to the clicked cell, not the union of old + new.
        assert_eq!(app.cursor_row, 4);
    }

    #[test]
    fn click_then_drag_after_prior_visual_starts_fresh_selection() {
        // Drag-after-Visual should anchor at the new click point, not
        // at the previous Visual selection's anchor. Confirms the
        // "click resets, drag re-anchors" flow end-to-end.
        let mut app = make_app();
        app.mode = Mode::Visual(VisualKind::Cell);
        app.selection_anchor = Some((0, 0));
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(35, 5), area);
        let new_anchor = (app.cursor_row, app.cursor_col);
        assert_eq!(app.selection_anchor, None); // reset by the click
        // Now drag to a further cell — the new selection anchors at the
        // click point, not at the discarded (0, 0) anchor.
        handle_mouse(&mut app, drag(60, 5), area);
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
        assert_eq!(app.selection_anchor, Some(new_anchor));
    }

    // ── T4: double-click → Edit mode ──────────────────────────────────

    use std::time::{Duration, Instant};

    #[test]
    fn is_double_click_pure_helper() {
        let cell = (0u32, 0u32);
        let t0 = Instant::now();
        // Same cell, well within threshold.
        assert!(super::is_double_click(
            Some((t0, cell)),
            t0 + Duration::from_millis(200),
            cell
        ));
        // Same cell, just at threshold: NOT a double-click (strict <).
        assert!(!super::is_double_click(
            Some((t0, cell)),
            t0 + Duration::from_millis(400),
            cell
        ));
        // Different cell, within threshold.
        assert!(!super::is_double_click(
            Some((t0, cell)),
            t0 + Duration::from_millis(100),
            (1, 0)
        ));
        // No prior click.
        assert!(!super::is_double_click(None, t0, cell));
    }

    #[test]
    fn double_click_within_threshold_enters_edit_mode() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Pre-seed A1 with a value so we can assert the buffer was loaded.
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        app.edit_insert('1');
        app.edit_insert('+');
        app.edit_insert('1');
        app.confirm_edit();
        // confirm_edit moves the cursor down — return to A1.
        app.cursor_row = 0;
        app.cursor_col = 0;
        // Simulate a recent click on A1.
        app.last_click = Some((Instant::now(), (0, 0)));
        // Second click on A1 within threshold (microseconds later).
        handle_mouse(&mut app, click(10, 5), area);
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.edit_buf, "=1+1"); // raw value, not computed
        assert_eq!(app.cursor_row, 0);
        assert_eq!(app.cursor_col, 0);
    }

    #[test]
    fn second_click_past_threshold_stays_single_click() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Last click was 1000ms ago — past the 400ms double-click threshold.
        app.last_click = Some((
            Instant::now() - Duration::from_millis(1000),
            (0, 0),
        ));
        handle_mouse(&mut app, click(10, 5), area);
        assert_eq!(app.mode, Mode::Nav); // single click, not Edit
    }

    #[test]
    fn second_click_on_different_cell_is_not_double_click() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Last click was on B1 — clicking A1 within threshold doesn't
        // count (different cells).
        app.last_click = Some((Instant::now(), (0, 1)));
        handle_mouse(&mut app, click(10, 5), area);
        assert_eq!(app.mode, Mode::Nav);
    }

    #[test]
    fn double_click_clears_drag_anchor_and_last_click() {
        // After entering Edit via double-click, neither the drag anchor
        // (T2) nor `last_click` should remain — a third click within
        // 400ms shouldn't re-trigger Edit (we're already there) or extend
        // a stale drag.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.last_click = Some((Instant::now(), (0, 0)));
        handle_mouse(&mut app, click(10, 5), area);
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.last_click, None);
        assert_eq!(app.drag_anchor, None);
    }

    // ── T15: auto-scroll while drag is held past the visible edge ────

    #[test]
    fn drag_past_bottom_scrolls_and_extends_selection() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Down at A1, drag into the grid to enter Visual::Cell.
        handle_mouse(&mut app, click(10, 5), area);
        handle_mouse(&mut app, drag(20, 21), area); // last visible data row
        let pre_scroll = app.scroll_row;
        let pre_cursor = app.cursor_row;
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
        // Drag past the bottom border (y=23 is the status bar in single-sheet).
        handle_mouse(&mut app, drag(20, 23), area);
        assert!(app.scroll_row > pre_scroll);
        assert!(app.cursor_row > pre_cursor);
        // Anchor unchanged; selection extends from (0, 0) to the new cursor.
        assert_eq!(app.selection_anchor, Some((0, 0)));
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
    }

    #[test]
    fn drag_past_right_scrolls_and_extends_selection() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(10, 5), area); // Down at A1
        // Walk into the grid first so we transition to Visual.
        handle_mouse(&mut app, drag(35, 5), area);
        let pre_scroll = app.scroll_col;
        let pre_cursor = app.cursor_col;
        // Drag past the right border (area is 80 wide; grid.x + width - 1 = 79).
        handle_mouse(&mut app, drag(79, 5), area);
        assert!(app.scroll_col > pre_scroll);
        assert!(app.cursor_col > pre_cursor);
        assert_eq!(app.selection_anchor, Some((0, 0)));
    }

    #[test]
    fn drag_past_top_at_origin_clamps_without_panic() {
        // Already at row 0 with scroll_row=0 — dragging up has nothing to
        // reveal. Should clamp gracefully, not panic.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(10, 5), area); // Down at A1, cursor (0, 0)
        handle_mouse(&mut app, drag(10, 1), area); // y=1 is the formula bar
        assert_eq!(app.scroll_row, 0);
        assert_eq!(app.cursor_row, 0);
    }

    #[test]
    fn drag_past_bottom_in_v_line_keeps_column() {
        // V-LINE (Visual::Row from a row-header click) auto-spans columns,
        // so a past-bottom drag should advance the cursor row but leave
        // the column where it was.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Land cursor at column 2 first so we can verify it's preserved.
        app.cursor_col = 2;
        handle_mouse(&mut app, click(2, 5), area); // RowHeader(0) → Visual::Row
        let pre_col = app.cursor_col;
        let pre_scroll = app.scroll_row;
        // Drag past bottom (y=23 = status bar).
        handle_mouse(&mut app, drag(2, 23), area);
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Row)));
        assert!(app.scroll_row > pre_scroll);
        assert_eq!(app.cursor_col, pre_col);
    }

    // ── T8: click cell during formula edit → insert ref at caret ─────

    #[test]
    fn click_cell_during_formula_edit_inserts_ref_at_caret() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Editing A1 with a half-formula buffer.
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        app.edit_insert('1');
        app.edit_insert('+');
        // Click on B2 → x in col 1 (~23), y=6 → row 1.
        handle_mouse(&mut app, click(23, 6), area);
        assert_eq!(app.mode, Mode::Edit); // still editing
        assert_eq!(app.edit_buf, "=1+B2");
        assert!(app.pointing.is_some());
    }

    #[test]
    fn click_cell_with_active_pointing_replaces_ref() {
        // After clicking B2 → "=1+B2" with pointing set, clicking C3
        // should *replace* B2 with C3 instead of appending.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        app.edit_insert('1');
        app.edit_insert('+');
        handle_mouse(&mut app, click(23, 6), area); // B2
        assert_eq!(app.edit_buf, "=1+B2");
        // Click C3 → x in col 2 (~36), y=7 → row 2.
        handle_mouse(&mut app, click(36, 7), area);
        assert_eq!(app.edit_buf, "=1+C3");
    }

    #[test]
    fn click_during_non_formula_edit_falls_through_to_commit() {
        // Editing A1 with plain text "hi" (no `=`); click B2 should
        // commit "hi" and move to B2 — original T1 Edit-mode behavior,
        // not a ref insertion.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('h');
        app.edit_insert('i');
        handle_mouse(&mut app, click(23, 6), area); // B2
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.cursor_row, 1);
        assert_eq!(app.cursor_col, 1);
        assert_eq!(app.get_raw(0, 0), "hi");
    }

    #[test]
    fn click_at_non_insertable_caret_falls_through_to_commit() {
        // Caret in the middle of a ref token (`=B1` with caret at index
        // 2 → between B and 1) — not insertable. Click should commit and
        // move, not insert another ref. Use B1 (not A1) so the formula
        // isn't a self-cycle that the engine would reject.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        app.edit_insert('B');
        app.edit_insert('1');
        app.edit_cursor = 2; // between B and 1
        handle_mouse(&mut app, click(36, 7), area); // C3
        assert_eq!(app.mode, Mode::Nav);
        assert_eq!(app.cursor_row, 2);
        assert_eq!(app.cursor_col, 2);
        assert_eq!(app.get_raw(0, 0), "=B1");
    }

    // ── T10 / T11: header click during formula edit → col/row ref ───

    // ── T9: drag during formula edit → range ref ─────────────────────

    #[test]
    fn drag_during_formula_edit_inserts_range_ref() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        app.edit_insert('S');
        app.edit_insert('U');
        app.edit_insert('M');
        app.edit_insert('(');
        // Click B2 → "=SUM(B2", pointing set with anchor=target=B2.
        handle_mouse(&mut app, click(23, 6), area);
        assert_eq!(app.edit_buf, "=SUM(B2");
        // Drag to B5 → range "=SUM(B2:B5".
        handle_mouse(&mut app, drag(23, 9), area);
        assert_eq!(app.edit_buf, "=SUM(B2:B5");
    }

    #[test]
    fn drag_normalizes_range_min_to_max() {
        // Drag in reverse direction (anchor below/right of target) still
        // produces a top-left → bottom-right range.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        // Click E5 first.
        handle_mouse(&mut app, click(60, 9), area); // col 4, row 4
        let after_click = app.edit_buf.clone();
        assert_eq!(after_click, "=E5");
        // Drag to B2 (above-left).
        handle_mouse(&mut app, drag(23, 6), area);
        // Range normalizes — B2 is the min, E5 is the max.
        assert_eq!(app.edit_buf, "=B2:E5");
    }

    #[test]
    fn drag_back_to_anchor_collapses_to_single_ref() {
        // Drag out into a range, then drag back to the anchor cell —
        // the range collapses to single-cell syntax (no `B2:B2`).
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        handle_mouse(&mut app, click(23, 6), area); // B2
        handle_mouse(&mut app, drag(36, 7), area); // C3 → "=B2:C3"
        assert_eq!(app.edit_buf, "=B2:C3");
        // Drag back to B2 → collapse to "=B2".
        handle_mouse(&mut app, drag(23, 6), area);
        assert_eq!(app.edit_buf, "=B2");
    }

    #[test]
    fn drag_outside_formula_falls_through_to_visual() {
        // Same drag plumbing, but no active pointing → ordinary T2
        // Visual::Cell drag, not range insertion.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(10, 5), area); // not in Edit mode
        handle_mouse(&mut app, drag(23, 6), area);
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Cell)));
    }

    #[test]
    fn col_header_click_during_formula_edit_inserts_col_ref() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        app.edit_insert('S');
        app.edit_insert('U');
        app.edit_insert('M');
        app.edit_insert('(');
        // Click column-B header (x≈25, y=4).
        handle_mouse(&mut app, click(25, 4), area);
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.edit_buf, "=SUM(B:B");
    }

    #[test]
    fn row_header_click_during_formula_edit_inserts_row_ref() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        app.edit_insert('A');
        app.edit_insert('V');
        app.edit_insert('G');
        app.edit_insert('(');
        // Click row-3 header (x in gutter, y=8 → row 3 (1-indexed display)).
        handle_mouse(&mut app, click(2, 8), area);
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.edit_buf, "=AVG(4:4");
    }

    #[test]
    fn col_header_click_replaces_active_pointing_ref() {
        // After a cell click set Cell pointing, a column-header click
        // should overwrite the ref span with the column ref and switch
        // pointing kind to Column so a follow-up drag can extend it.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        handle_mouse(&mut app, click(23, 6), area); // B2 — sets Cell pointing
        assert_eq!(app.edit_buf, "=B2");
        // Column-C header click — replaces the span and switches to Column.
        handle_mouse(&mut app, click(38, 4), area);
        assert_eq!(app.edit_buf, "=C:C");
        let p = app.pointing.expect("Column pointing established");
        assert_eq!(p.kind, app::PointingKind::Column);
    }

    #[test]
    fn col_header_click_outside_formula_falls_through_to_v_column() {
        // Same handler in non-Edit mode is the original T6+T16 path.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(10, 4), area); // ColumnHeader(0)
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Column)));
    }

    #[test]
    fn row_header_click_outside_formula_falls_through_to_v_line() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        handle_mouse(&mut app, click(2, 6), area); // RowHeader(1)
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Row)));
    }

    // ── T12 / T13: header drag during formula edit → col/row range ─

    #[test]
    fn col_header_drag_during_formula_edit_inserts_col_range() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        // Click column-B header → "=B:B".
        handle_mouse(&mut app, click(25, 4), area);
        assert_eq!(app.edit_buf, "=B:B");
        // Drag to column-D header (x≈51, y=4).
        handle_mouse(&mut app, drag(51, 4), area);
        assert_eq!(app.edit_buf, "=B:D");
    }

    #[test]
    fn col_header_drag_back_collapses_range() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        handle_mouse(&mut app, click(25, 4), area); // ColumnHeader B → "=B:B"
        handle_mouse(&mut app, drag(51, 4), area); // → "=B:D"
        // Drag back to column B → collapses to "=B:B".
        handle_mouse(&mut app, drag(25, 4), area);
        assert_eq!(app.edit_buf, "=B:B");
    }

    #[test]
    fn row_header_drag_during_formula_edit_inserts_row_range() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        // Click row-1 header (y=5 → row 0, displayed as 1).
        handle_mouse(&mut app, click(2, 5), area);
        assert_eq!(app.edit_buf, "=1:1");
        // Drag to row-5 header (y=9 → row 4, displayed as 5).
        handle_mouse(&mut app, drag(2, 9), area);
        assert_eq!(app.edit_buf, "=1:5");
    }

    #[test]
    fn col_drag_normalizes_reverse_direction() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        // Click column-D header first.
        handle_mouse(&mut app, click(51, 4), area);
        assert_eq!(app.edit_buf, "=D:D");
        // Drag to column-B header (left of anchor).
        handle_mouse(&mut app, drag(25, 4), area);
        // Range normalizes — B is min, D is max.
        assert_eq!(app.edit_buf, "=B:D");
    }

    #[test]
    fn double_click_is_suppressed_during_edit() {
        // Two fast clicks while editing should not throw away the buffer
        // via start_edit; the first inserts a ref (formula context) and
        // the second replaces or falls through, but mode stays Edit.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.start_edit_blank();
        app.edit_insert('=');
        // Simulate a recent click on B2 — would normally fire T4.
        app.last_click = Some((Instant::now(), (1, 1)));
        handle_mouse(&mut app, click(23, 6), area);
        // Still editing; buffer was extended via T8, not wiped by T4.
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.edit_buf, "=B2");
    }

    #[test]
    fn drag_past_edge_without_prior_down_is_noop() {
        // Drag arriving without a remembered click anchor — typical of
        // a stray event right after `:set mouse` flips on. Don't scroll.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        let pre_scroll = app.scroll_row;
        handle_mouse(&mut app, drag(10, 23), area);
        assert_eq!(app.scroll_row, pre_scroll);
    }

    fn shift_v() -> event::KeyEvent {
        event::KeyEvent::new(KeyCode::Char('V'), Mods::SHIFT)
    }

    #[test]
    fn shift_v_cycles_nav_then_v_line_then_v_column_then_nav() {
        // VV is the keyboard route to V-COLUMN. First press enters V-LINE,
        // second promotes to V-COLUMN, third exits.
        let mut app = make_app();
        handle_nav_key(&mut app, shift_v());
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Row)));
        handle_nav_key(&mut app, shift_v());
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Column)));
        handle_nav_key(&mut app, shift_v());
        assert_eq!(app.mode, Mode::Nav);
    }

    #[test]
    fn shift_v_from_v_cell_promotes_to_v_line() {
        // Pre-existing behavior: pressing V from cell-visual promotes to
        // V-LINE rather than exiting. Lock it in so the cycle change above
        // doesn't accidentally regress this.
        let mut app = make_app();
        app.enter_visual(VisualKind::Cell);
        handle_nav_key(&mut app, shift_v());
        assert!(matches!(app.mode, Mode::Visual(VisualKind::Row)));
    }

    #[test]
    fn equals_in_nav_starts_formula_edit() {
        // `=` from Normal opens Insert with `=` already in the buffer, so
        // the user is one keystroke into typing a formula. Ctrl+= and
        // `Ctrl+w =` cover the autofit role.
        let mut app = make_app();
        handle_nav_key(
            &mut app,
            event::KeyEvent::new(KeyCode::Char('='), Mods::NONE),
        );
        assert_eq!(app.mode, Mode::Edit);
        assert_eq!(app.edit_buf, "=");
        assert_eq!(app.edit_cursor, 1);
    }

    #[test]
    fn ctrl_equals_in_nav_autofits_column() {
        // The autofit binding moved from bare `=` to Ctrl+= (and Ctrl+w =).
        // Seed B0 with content that's wider than the default 12 char width
        // so autofit has somewhere to grow to.
        let mut app = make_app();
        app.cursor_col = 1;
        app.start_edit_blank();
        for ch in "supercalifragilisticexpialidocious".chars() {
            app.edit_insert(ch);
        }
        app.confirm_edit();
        let pre = app.column_width(1);
        app.cursor_col = 1;
        handle_nav_key(
            &mut app,
            event::KeyEvent::new(KeyCode::Char('='), Mods::CONTROL),
        );
        assert!(app.column_width(1) > pre);
        assert!(app.status.starts_with("Column B"));
    }

    #[test]
    fn ctrl_equals_in_v_column_autofits_every_selected_column() {
        // Multi-column selection + Ctrl+= → autofit each column in
        // range, not just the cursor's. V-COLUMN is the cleanest path
        // (whole columns selected); Visual::Cell extending across
        // columns goes through the same code path.
        let mut app = make_app();
        seed_cell(&mut app, 0, 1, "supercalifragilisticexpialidocious");
        seed_cell(&mut app, 0, 2, "another_long_value_for_column_C");
        seed_cell(&mut app, 0, 3, "and_one_more_for_column_D");
        let pre_b = app.column_width(1);
        let pre_c = app.column_width(2);
        let pre_d = app.column_width(3);
        app.cursor_col = 1;
        app.enter_visual(VisualKind::Column);
        app.cursor_col = 3; // extend to column D
        handle_nav_key(
            &mut app,
            event::KeyEvent::new(KeyCode::Char('='), Mods::CONTROL),
        );
        assert!(app.column_width(1) > pre_b);
        assert!(app.column_width(2) > pre_c);
        assert!(app.column_width(3) > pre_d);
        assert!(app.status.contains("3 columns"));
        assert!(app.status.contains("B:D"));
    }

    #[test]
    fn ctrl_w_equals_in_v_column_autofits_every_selected_column() {
        // Symmetric multi-column behavior via the `Ctrl+w =` fallback.
        let mut app = make_app();
        seed_cell(&mut app, 0, 1, "supercalifragilisticexpialidocious");
        seed_cell(&mut app, 0, 2, "another_long_value_for_column_C");
        let pre_b = app.column_width(1);
        let pre_c = app.column_width(2);
        app.cursor_col = 1;
        app.enter_visual(VisualKind::Column);
        app.cursor_col = 2;
        app.pending_ctrl_w = true;
        handle_nav_key(
            &mut app,
            event::KeyEvent::new(KeyCode::Char('='), Mods::NONE),
        );
        assert!(app.column_width(1) > pre_b);
        assert!(app.column_width(2) > pre_c);
        assert!(app.status.contains("2 columns"));
    }

    #[test]
    fn ctrl_w_equals_autofits_as_fallback() {
        // Ctrl+= isn't reliably emitted by every terminal; `Ctrl+w =`
        // is the symmetric fallback alongside `Ctrl+w >` / `Ctrl+w <`.
        let mut app = make_app();
        app.cursor_col = 2;
        app.start_edit_blank();
        for ch in "long_value_that_needs_more_room".chars() {
            app.edit_insert(ch);
        }
        app.confirm_edit();
        app.cursor_col = 2;
        let pre = app.column_width(2);
        app.pending_ctrl_w = true;
        handle_nav_key(
            &mut app,
            event::KeyEvent::new(KeyCode::Char('='), Mods::NONE),
        );
        assert!(app.column_width(2) > pre);
    }

    /// Type `value` into `(row, col)` and re-park the cursor on it
    /// (`confirm_edit_advancing` moves down by one).
    fn seed_cell(app: &mut App, row: u32, col: u32, value: &str) {
        app.cursor_row = row;
        app.cursor_col = col;
        app.start_edit_blank();
        for c in value.chars() {
            app.edit_insert(c);
        }
        app.confirm_edit();
        app.cursor_row = row;
        app.cursor_col = col;
    }

    fn key(c: char) -> event::KeyEvent {
        event::KeyEvent::new(KeyCode::Char(c), Mods::NONE)
    }

    #[test]
    fn bare_f_in_nav_sets_pending_f() {
        let mut app = make_app();
        handle_nav_key(&mut app, key('f'));
        assert!(app.pending_f);
        // No collateral state changes.
        assert!(!app.pending_g);
        assert!(app.pending_count.is_none());
        assert!(matches!(app.mode, crate::app::Mode::Nav));
    }

    #[test]
    fn esc_clears_pending_f() {
        let mut app = make_app();
        app.pending_f = true;
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Esc, Mods::NONE));
        assert!(!app.pending_f);
    }

    #[test]
    fn unknown_f_combo_clears_pending_f() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        // `q` is not a format axis key.
        handle_nav_key(&mut app, key('q'));
        assert!(!app.pending_f);
        // Nothing applied.
        assert!(app.get_format(0, 0).is_none());
    }

    #[test]
    fn f_dollar_applies_usd_format() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "1.25");
        handle_nav_key(&mut app, key('f'));
        assert!(app.pending_f);
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('$'), Mods::SHIFT));
        assert!(!app.pending_f);
        assert_eq!(
            app.get_format_json_raw(0, 0).as_deref(),
            Some(r#"{"n":{"k":"usd","d":2}}"#)
        );
        assert_eq!(app.displayed_for(0, 0), "$1.25");
    }

    #[test]
    fn f_dot_increases_decimals_f_comma_decreases() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "$1.25");
        // f. → decimals 2 → 3
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('.'));
        assert_eq!(app.displayed_for(0, 0), "$1.250");
        // f, → decimals 3 → 2
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key(','));
        assert_eq!(app.displayed_for(0, 0), "$1.25");
    }

    #[test]
    fn old_g_dollar_no_longer_applies_usd() {
        // Regression for migrate-f F1: hard-swap means `g$` is gone.
        // Pressing `g` sets pending_g, then `$` should fall through the
        // pending_g consumer's catch-all (which clears the flag) without
        // touching the cell's format. The bare `$` motion runs only on
        // the *next* keypress after pending_g clears, but the catch-all
        // returns before reaching the late match — so this is a no-op.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "1.25");
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('$'), Mods::SHIFT));
        assert!(app.get_format(0, 0).is_none());
        assert_eq!(app.displayed_for(0, 0), "1.25");
    }

    #[test]
    #[allow(non_snake_case)]
    fn fF_opens_color_picker_for_fg() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('F'), Mods::SHIFT));
        assert!(matches!(app.mode, crate::app::Mode::ColorPicker));
        let state = app.color_picker.as_ref().expect("picker open");
        assert!(matches!(state.kind, crate::app::ColorPickerKind::Fg));
    }

    #[test]
    #[allow(non_snake_case)]
    fn fB_opens_color_picker_for_bg() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('B'), Mods::SHIFT));
        assert!(matches!(app.mode, crate::app::Mode::ColorPicker));
        let state = app.color_picker.as_ref().expect("picker open");
        assert!(matches!(state.kind, crate::app::ColorPickerKind::Bg));
    }

    #[test]
    fn picker_navigates_and_applies_swatch() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('F'), Mods::SHIFT));
        // Step right twice (cursor 0 → 2).
        handle_color_picker_key(&mut app, key('l'));
        handle_color_picker_key(&mut app, key('l'));
        assert_eq!(app.color_picker.as_ref().unwrap().cursor, 2);
        // Apply.
        handle_color_picker_key(&mut app, event::KeyEvent::new(KeyCode::Enter, Mods::NONE));
        assert!(matches!(app.mode, crate::app::Mode::Nav));
        let fmt = app.get_format(0, 0).expect("fg applied");
        assert_eq!(fmt.fg, Some(crate::format::COLOR_PRESETS[2].1));
    }

    #[test]
    fn picker_esc_cancels_without_applying() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('F'), Mods::SHIFT));
        handle_color_picker_key(&mut app, event::KeyEvent::new(KeyCode::Esc, Mods::NONE));
        assert!(matches!(app.mode, crate::app::Mode::Nav));
        assert!(app.color_picker.is_none());
        // Cell got no format change.
        assert!(app.get_format(0, 0).is_none());
    }

    #[test]
    fn picker_hex_input_applies_typed_color() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('F'), Mods::SHIFT));
        handle_color_picker_key(&mut app, event::KeyEvent::new(KeyCode::Char('?'), Mods::NONE));
        for c in "fa3".chars() {
            handle_color_picker_key(&mut app, key(c));
        }
        handle_color_picker_key(&mut app, event::KeyEvent::new(KeyCode::Enter, Mods::NONE));
        let fmt = app.get_format(0, 0).expect("fg applied");
        assert_eq!(fmt.fg, Some(crate::format::Color::rgb(0xff, 0xaa, 0x33)));
    }

    #[test]
    fn picker_bad_hex_reports_status_and_stays_open() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('F'), Mods::SHIFT));
        handle_color_picker_key(&mut app, event::KeyEvent::new(KeyCode::Char('?'), Mods::NONE));
        // Type just one digit (`#a`) — apply should fail with Bad hex.
        handle_color_picker_key(&mut app, key('a'));
        handle_color_picker_key(&mut app, event::KeyEvent::new(KeyCode::Enter, Mods::NONE));
        assert!(app.status.contains("Bad hex"));
        assert!(matches!(app.mode, crate::app::Mode::ColorPicker));
    }

    #[test]
    fn f_percent_applies_percent_format() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "0.05");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('%'), Mods::SHIFT));
        assert_eq!(
            app.get_format(0, 0).unwrap().number,
            Some(crate::format::NumberFormat::Percent { decimals: 0 })
        );
    }

    #[test]
    fn fb_toggles_bold() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('b'));
        assert!(app.get_format(0, 0).unwrap().bold);
        // Toggle off.
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('b'));
        assert!(!app.get_format(0, 0).unwrap().bold);
    }

    #[test]
    fn fi_fu_fs_toggle_their_flags() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('i'));
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('u'));
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('s'));
        let fmt = app.get_format(0, 0).unwrap();
        assert!(fmt.italic);
        assert!(fmt.underline);
        assert!(fmt.strike);
        // Bold not toggled — only the targeted flags.
        assert!(!fmt.bold);
    }

    #[test]
    fn old_gb_no_longer_toggles_bold() {
        // Regression for migrate-f F2: hard-swap means `gb` is gone.
        // `g` sets pending_g; the consumer's catch-all clears it on
        // `b`. The bare `b` arm has !pending_f (not !pending_g) so
        // post-consumer the late-match `b` would run word-back —
        // BUT the pending_g consumer returns early after handling,
        // so we never reach the late match. Either way: format
        // unchanged.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, key('b'));
        assert!(app.get_format(0, 0).is_none_or(|f| !f.bold));
    }

    #[test]
    fn fl_fc_fr_set_alignment() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('l'));
        assert_eq!(app.get_format(0, 0).unwrap().align, Some(crate::format::Align::Left));
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('c'));
        assert_eq!(app.get_format(0, 0).unwrap().align, Some(crate::format::Align::Center));
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('r'));
        assert_eq!(app.get_format(0, 0).unwrap().align, Some(crate::format::Align::Right));
    }

    #[test]
    fn fa_clears_alignment() {
        // fa is the explicit "auto" — drops format.align back to None
        // so classify_display picks per-type alignment.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('l'));
        assert_eq!(app.get_format(0, 0).unwrap().align, Some(crate::format::Align::Left));
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('a'));
        assert!(app.get_format(0, 0).unwrap().align.is_none());
    }

    #[test]
    fn bare_c_still_operator_when_no_pending_f() {
        // Negative test: bare `c` followed by a motion still acts as
        // the Change operator (clear rect + enter Edit). Seed two
        // filled cells so `cw` has a target.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "abc");
        seed_cell(&mut app, 0, 1, "def");
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, key('c'));
        assert!(matches!(app.pending_operator, Some(Operator::Change)));
    }

    #[test]
    fn bare_a_still_inserts_when_no_pending_f() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "abc");
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, key('a'));
        assert!(matches!(app.mode, crate::app::Mode::Edit));
    }

    #[test]
    fn bare_l_still_moves_right_when_no_pending_f() {
        let mut app = make_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, key('l'));
        assert_eq!(app.cursor_col, 1);
    }

    #[test]
    #[allow(non_snake_case)]
    fn bare_C_still_change_to_end_when_no_pending_g() {
        // The `!pending_g` guard on bare `C` makes `gC` a quiet
        // catch-all no-op. Bare `C` (no prefix) must still run
        // change-to-end-of-row.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        seed_cell(&mut app, 0, 1, "y");
        seed_cell(&mut app, 0, 2, "z");
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('C'), Mods::SHIFT));
        assert!(matches!(app.mode, crate::app::Mode::Edit));
    }

    #[test]
    #[allow(non_snake_case)]
    fn bare_L_still_viewport_bottom_when_no_pending_g() {
        // Symmetric to bare_C — `gL` is a no-op via consumer
        // catch-all; bare `L` still moves to viewport-bottom.
        let mut app = make_app();
        let _area = render(&mut app, 80, 24);
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('L'), Mods::SHIFT));
        assert!(app.cursor_row > 0);
    }

    #[test]
    fn old_uppercase_g_align_no_longer_applies() {
        // Regression for migrate-f F3: `gL` / `gC` / `gR` / `gA` are
        // gone. Pressing them should leave format.align untouched.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('L'), Mods::SHIFT));
        assert!(app.get_format(0, 0).is_none_or(|f| f.align.is_none()));
    }

    #[test]
    fn bare_b_still_word_back_when_no_pending_f() {
        // Negative test: the !pending_f guard on the `b` arm mustn't
        // break the bare word-back motion. Seed two filled cells so
        // word_backward has somewhere to go.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "a");
        seed_cell(&mut app, 0, 5, "b");
        app.cursor_row = 0;
        app.cursor_col = 5;
        handle_nav_key(&mut app, key('b'));
        assert_eq!(app.cursor_col, 0);
    }

    #[test]
    fn bare_u_still_undo_when_no_pending_f() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "x");
        // Edit it so undo has something to roll back.
        app.cursor_row = 0;
        app.cursor_col = 0;
        app.delete_cell();
        handle_nav_key(&mut app, key('u'));
        assert_eq!(app.get_raw(0, 0), "x");
    }

    #[test]
    fn bare_i_still_inserts_when_no_pending_f() {
        // Negative test: bare `i` enters Edit mode (insert at start).
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "abc");
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, key('i'));
        assert!(matches!(app.mode, crate::app::Mode::Edit));
    }

    #[test]
    fn bare_s_still_substitutes_when_no_pending_f() {
        // Negative test: bare `s` enters Edit mode with empty buffer.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "abc");
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, key('s'));
        assert!(matches!(app.mode, crate::app::Mode::Edit));
        // Substitute opens with empty buffer (cancel preserves original).
        assert_eq!(app.edit_buf, "");
    }

    #[test]
    fn f_number_keys_clear_pending_f() {
        // Each number-axis key clears the pending_f flag on entry to
        // the consumer, so they don't leave the prefix stuck.
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "1");
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('$'), Mods::SHIFT));
        assert!(!app.pending_f);
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('.'));
        assert!(!app.pending_f);
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key(','));
        assert!(!app.pending_f);
    }

    #[test]
    fn f_format_keys_clear_pending_f() {
        // Each format-axis key clears the pending_f flag on entry to
        // the consumer (one binding per axis exercised here).
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "1");
        // Number axis.
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('$'), Mods::SHIFT));
        assert!(!app.pending_f);
        // Style axis.
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('b'));
        assert!(!app.pending_f);
        // Align axis.
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, key('c'));
        assert!(!app.pending_f);
        // Color axis.
        handle_nav_key(&mut app, key('f'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('F'), Mods::SHIFT));
        assert!(!app.pending_f);
    }

    #[test]
    fn pending_g_consumer_only_handles_motion_keys_after_migrate_f() {
        // Regression: after migrate-f, pending_g routes ONLY gg/gv/gt/gT.
        // Pressing a former-format-axis key under pending_g should be a
        // no-op (catch-all clears the flag without any side effect).
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "1.25");

        // gb (was bold)
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, key('b'));
        assert!(app.get_format(0, 0).is_none_or(|f| !f.bold));

        // g$ (was USD)
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('$'), Mods::SHIFT));
        assert!(app.get_format(0, 0).is_none_or(|f| f.number.is_none()));

        // gC (was center)
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, event::KeyEvent::new(KeyCode::Char('C'), Mods::SHIFT));
        assert!(app.get_format(0, 0).is_none_or(|f| f.align.is_none()));

        // gf (was fg picker) — bare `f` after pending_g hits the
        // catch-all (no `f` arm in the g-consumer post-migration).
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, key('f'));
        assert!(matches!(app.mode, crate::app::Mode::Nav));

        // gg still works.
        app.cursor_row = 5;
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, key('g'));
        assert_eq!(app.cursor_row, 0);
    }

    #[test]
    fn lowercase_x_clears_in_place_with_count() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "a");
        seed_cell(&mut app, 0, 1, "b");
        seed_cell(&mut app, 0, 2, "c");
        seed_cell(&mut app, 0, 3, "d");
        app.cursor_row = 0;
        app.cursor_col = 1;

        // Bare `x` clears only the cursor cell; cursor stays put.
        handle_nav_key(&mut app, key('x'));
        assert_eq!(app.get_raw(0, 0), "a");
        assert_eq!(app.get_raw(0, 1), "");
        assert_eq!(app.get_raw(0, 2), "c");
        assert_eq!(app.get_raw(0, 3), "d");
        assert_eq!((app.cursor_row, app.cursor_col), (0, 1));

        // Move to C1 and run `2x` — clears C1 + D1 (cursor + 1), cursor stays.
        app.cursor_col = 2;
        handle_nav_key(&mut app, key('2'));
        handle_nav_key(&mut app, key('x'));
        assert_eq!(app.get_raw(0, 2), "");
        assert_eq!(app.get_raw(0, 3), "");
        assert_eq!((app.cursor_row, app.cursor_col), (0, 2));
    }

    // ── go / ctrl-click hyperlink open (T2) ──────────────────────────

    #[test]
    fn url_under_cursor_returns_url_for_text_url_cell() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "https://example.com");
        app.cursor_row = 0;
        app.cursor_col = 0;
        assert_eq!(
            app.url_under_cursor().as_deref(),
            Some("https://example.com")
        );
    }

    #[test]
    fn url_under_cursor_returns_url_portion_for_hyperlink_cell() {
        // HYPERLINK(url, label) cell stores `url\u{1F}label` in
        // cell.computed; url_under_cursor must return the URL portion,
        // not the label, so `go` opens the right thing.
        let mut app = make_app();
        seed_cell(
            &mut app,
            0,
            0,
            "=HYPERLINK(\"https://example.com\", \"click here\")",
        );
        app.cursor_row = 0;
        app.cursor_col = 0;
        assert_eq!(
            app.url_under_cursor().as_deref(),
            Some("https://example.com"),
        );
        // And the displayed value is the label, not the URL.
        assert_eq!(app.displayed_for(0, 0), "click here");
    }

    #[test]
    fn url_under_cursor_none_for_plain_text() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "hello");
        app.cursor_row = 0;
        app.cursor_col = 0;
        assert_eq!(app.url_under_cursor(), None);
    }

    #[test]
    fn url_under_cursor_none_for_empty_cell() {
        let mut app = make_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        assert_eq!(app.url_under_cursor(), None);
    }

    #[test]
    fn go_with_url_under_cursor_flashes_opened_status() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "https://example.com");
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, key('g'));
        assert!(app.pending_g);
        handle_nav_key(&mut app, key('o'));
        assert!(!app.pending_g, "pending_g should clear after go");
        assert!(
            app.status.starts_with("Opened "),
            "expected 'Opened ...' status, got {:?}",
            app.status
        );
        assert!(app.status.contains("https://example.com"));
    }

    #[test]
    fn go_with_no_url_flashes_no_link_status() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "plain text");
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, key('o'));
        assert!(!app.pending_g);
        assert_eq!(app.status, "No link under cursor");
    }

    #[test]
    fn go_on_empty_cell_flashes_no_link_status() {
        let mut app = make_app();
        app.cursor_row = 3;
        app.cursor_col = 3;
        handle_nav_key(&mut app, key('g'));
        handle_nav_key(&mut app, key('o'));
        assert_eq!(app.status, "No link under cursor");
    }

    #[test]
    fn bare_o_without_pending_g_still_opens_below() {
        // Regression guard: `go` hooks into the pending_g consumer,
        // so bare `o` (open new line below + Insert) must still run.
        // Press `o` from row 0, confirm cursor moved down and Edit
        // mode is active.
        let mut app = make_app();
        app.cursor_row = 0;
        app.cursor_col = 0;
        handle_nav_key(&mut app, key('o'));
        assert_eq!(app.cursor_row, 1);
        assert!(matches!(app.mode, crate::app::Mode::Edit));
    }

    fn ctrl_click(x: u16, y: u16) -> MouseEvent {
        MouseEvent {
            kind: MouseEventKind::Down(MouseButton::Left),
            column: x,
            row: y,
            modifiers: Mods::CONTROL,
        }
    }

    #[test]
    fn ctrl_click_on_url_cell_opens_url() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "https://example.com");
        let area = render(&mut app, 80, 24);
        // Park cursor far from A1 to make the cursor-jump assertion meaningful.
        app.cursor_row = 4;
        app.cursor_col = 3;
        handle_mouse(&mut app, ctrl_click(10, 5), area);
        // Cursor moved to clicked cell.
        assert_eq!(app.cursor_row, 0);
        assert_eq!(app.cursor_col, 0);
        // Status reflects the open.
        assert!(
            app.status.starts_with("Opened "),
            "expected 'Opened ...' status, got {:?}",
            app.status
        );
    }

    #[test]
    fn ctrl_click_on_non_url_cell_falls_through_to_select() {
        let mut app = make_app();
        seed_cell(&mut app, 0, 0, "just text");
        let area = render(&mut app, 80, 24);
        app.cursor_row = 4;
        app.cursor_col = 3;
        handle_mouse(&mut app, ctrl_click(10, 5), area);
        // Cursor still moved (normal click path took over).
        assert_eq!(app.cursor_row, 0);
        assert_eq!(app.cursor_col, 0);
        // Status NOT set to opened.
        assert!(!app.status.starts_with("Opened "));
    }
}
