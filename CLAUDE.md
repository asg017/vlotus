# vlotus — Vim-style Terminal UI

Standalone terminal spreadsheet application using `ratatui`. Not part of the Datasette plugin — this is a separate tool for CLI use.

## Usage

```bash
cargo run -p vlotus [db_path]            # Interactive TUI
cargo run -p vlotus eval db_path cell    # One-shot cell evaluation
cargo run -p vlotus tutor                # Bundled vimtutor-style lessons
cargo run -p vlotus patch <op> [args...] # Work with .lpatch files (apply / show / invert / combine)
```

## Modes

`Mode::Nav` (default), `Mode::Edit`, `Mode::Command` (`:` prompt), `Mode::Search(SearchDir)` (`/` `?`), `Mode::Visual(VisualKind::Cell|Row|Column)`, and `Mode::ColorPicker` (modal swatch picker opened by `fF` / `fB`). The run loop in `main.rs::run_loop` dispatches by mode; `handle_nav_key` covers Normal *and* Visual (motions/operators are mode-aware via `selection_anchor` and `pending_*` flags). `V` from Normal cycles into V-LINE; a second `V` promotes V-LINE → V-COLUMN, so `VV` is the keyboard route to whole-column visual selection (column-header click is the mouse route).

Vim grammar: count prefix (`pending_count`), `g`/`z`/`m`/`'`/`` ` ``/`Ctrl+w` two-keystroke prefixes, operator-pending (`pending_operator` + motion target via `motion_target_for_operator` in `main.rs`), search/marks, `.` dot-repeat (`last_edit`).

Detailed reference for every key, command, App method, and backing unit test: see `KEYMAP.md`. Subtle behaviors (e.g. `s<Esc>` vs `cc<Esc>` differ; `Y` yanks a cell not a row; forward operator motions are inclusive) are flagged there.

**When you add a new key binding or `:command`, update `KEYMAP.md` in the same change.** It's the canonical reference and goes stale fast otherwise. Match the existing table format (binding · what it does · backing test name).

## Where new key bindings go

`main.rs::handle_nav_key` serves both Normal and Visual. Block placement determines which modes see your binding — misplacing it is the most common bug class here (the Ctrl+= autofit and Ctrl+A expand fixes both initially landed in the wrong block and had to be relocated). The function has a header comment listing the dispatch order; the decision tree:

1. **Two-keystroke prefix that conflicts with a bare-key arm** (e.g. `fc`, `fa` where `c`/`a` are operators / Insert) → goes in the `pending_f` consumer block, *before* the operator and Visual blocks. Pattern: set `pending_f = true` on the first key in the bare-letter arm; consume in the `pending_f` block.
2. **Mode-agnostic `Ctrl+<key>`** that should fire in both Nav and Visual (autofit, select-all, future ones) → goes in the **mode-agnostic Ctrl+ block** (around the existing Ctrl+= and Ctrl+A arms), which sits *before* the Visual short-circuit.
3. **Visual-mode-only command** (yank, delete, V toggle) → goes in the Visual short-circuit block.
4. **Nav-only `Ctrl+<key>`** (clipboard, undo, jumps) → goes in the Nav-side `if ctrl { match … }` block, which is past the Visual short-circuit.
5. **Bare letter motion / operator** → bottom of the function. Add a `!pending_f` guard if the letter is one of the alignment letters consumed by `f`-prefix.

Quick sanity test: in Visual mode, does your binding still fire? If it should and doesn't, you're in block 4 instead of 2. If it shouldn't but does, you're in block 2 instead of 4.

## Mouse

**On by default**, matching Excel / VisiData / sc-im. Click to select a cell, click-and-drag to enter `Visual::Cell`, scroll wheel to scroll, double-click to edit, click a tab to switch sheets, click row/column headers for V-LINE / column-mode visual. `:set nomouse` releases capture for the rest of the session if the user wants the terminal's native text-selection (Cmd/Option+drag → copy) back; `:set mouse` re-enables; `:set mouse?` reports current state. `run_loop` diffs `app.mouse_enabled` after each event tick and emits `EnableMouseCapture` / `DisableMouseCapture` on transitions; `run_terminal` only emits a final `DisableMouseCapture` on shutdown if the session ended with capture on. Hit-testing lives in `ui::cell_at` and returns `HitTarget::{Cell, RowHeader, ColumnHeader, Tab}`.

## Tabline

When `app.sheets.len() > 1`, `ui::draw` inserts a 1-line tabline between the grid and the status bar (Excel / Sheets convention) that lists every sheet as ` {name} ` separated by `│`. The active tab renders in reverse-video; inactive tabs in `Color::DarkGray`. Truncation: names cap at 16 chars (`…` suffix); when the full list exceeds the terminal width, `tabline_window` picks a contiguous slice that includes the active tab and signals elided runs with leading / trailing `…`. Hidden when only one sheet exists, so single-sheet workbooks keep their full grid height.

## Cell formatting

`format::CellFormat` is a struct with composable axes — `number: Option<NumberFormat>` (Usd/Percent), bool style flags (`bold`, `italic`, `underline`, `strike`), `align: Option<Align>`, and `fg`/`bg: Option<Color>`. Each axis is independent: a cell can be USD-formatted + bold + center-aligned + red text, all at once. Format JSON lives in the `cell_format.format_json` column (joined into `Store::load_sheet` via LEFT JOIN on `(sheet_name, row, col)`). JSON layout: `{"n":{"k":"usd","d":2},"b":true,"a":"center","fg":"#ff0000","bg":"#1e1e2e"}` — unset fields are omitted. Hand-rolled parser in `format.rs` (no serde dep).

Mutation path: `App::apply_format_update(F)` reads a cell's existing format, runs the closure to mutate one axis, writes back. `:fmt usd` on a bold cell preserves bold — every command merges. Full-clear via `:fmt clear` sends the literal `"null"` JSON sentinel which `Store::apply` interprets as "delete the cell_format row" (versus `None` which preserves the existing format). `capture_undo_entry` writes the same sentinel when the prior cell had no format, so undo of "applied a format to an unformatted cell" actually drops the format instead of preserving it.

Render path: `App::displayed_for(r, c)` is the single source of truth for the formatted display string, used by both `ui::draw_grid` and `autofit_column` so column-width math agrees with what's drawn. Style flags layer ratatui `Modifier::{BOLD,ITALIC,UNDERLINED,CROSSED_OUT}` on top of the state-driven cell style — a bold cell stays bold under cursor / selection / search highlights. Format `fg` always wins (red text stays red even under selection bg); format `bg` defers to highlight bgs so cursor / selection / search still paint correctly. Alignment override lives in `ui::align_text_override` and falls back to `align_text(kind, …)` when `format.align == None`.

Edit-commit auto-detect: `format::try_parse_typed_input` peels `$1.25` → 1.25 + USD/2 and `4.5%` → 0.045 + Percent/1; `=`-formulas (including `=$A$1` absolute refs) pass through unchanged. `cell_change_from_typed` is the shared helper used by `confirm_edit_advancing` and `repeat_last_edit` so dot-repeat re-runs auto-detect at the new cell.

Color picker (`Mode::ColorPicker`, opened by `fF` for fg / `fB` for bg): `App::color_picker: Option<ColorPickerState>` holds the live state (kind, swatch cursor, hex-input buffer, target rect captured at open-time). `handle_color_picker_key` dispatches hjkl nav / Enter apply / Esc cancel / `?` hex toggle. Render overlay is `ui::draw_color_picker`, a centered popup with a 4×6 swatch grid sourced from `format::COLOR_PRESETS` (Catppuccin Mocha aliased with natural-language names — `red` / `orange` / `purple` / `cyan` route to mocha::RED / PEACH / MAUVE / SKY).

All cell-format keys live under the **`f`-prefix** (number / style / align / color axes — see `KEYMAP.md`'s "Cell formatting" section). The `pending_f` consumer block is placed *before* the operator/Nav-only dispatch in `handle_nav_key` so lowercase second-keys (`fc` center, `fa` auto-align, `fl` left-align, `fr` right-align) reach the consumer instead of being absorbed by the Change operator / Insert / motion-right arms. Bare `b/i/u/s/c/a/A/l/.` carry defensive `!pending_f` guards even though the consumer placement makes them redundant.

When you add a new format axis or `:fmt` command, the canonical spots to update: `format::CellFormat` struct + JSON parser, `apply_format_update` callers in `run_command`, render integration in `ui::draw_grid`, and `KEYMAP.md` in the same change.

## Column widths

`column_meta.width` is interpreted as **character count** by vlotus. `App::column_width(col_idx)` is the canonical accessor; `set_column_width` and `autofit_column` are the mutators. Default is `DEFAULT_COL_WIDTH = 12`. `ui::draw_grid` walks per-column widths to compute `visible_cols`, charging each visible column its declared width plus `COL_SPACING = 1` (the cell ratatui's `Table` inserts between adjacent constraints). `active_cell_rect` and the `Table::column_spacing` setting agree on the same constant so the autocomplete popup anchors correctly. Skipping the spacing accounting silently overflows the inner Table area and ratatui's solver compensates by shrinking the widest column — clipping right-aligned numerics like `100` to `10`.

## Snapshot tests

`src/snapshots.rs` covers ~20 styled UI states via `ratatui::backend::TestBackend` + `insta`. Most use plain-text snapshots; the four scenes whose feature is purely styling (visual rect, V-LINE, search match tint, clipboard mark perimeter) use `render_with_highlights` which appends a per-row range listing of cells by category. Format-related scenes: `snapshot_currency_formatted_grid`, `snapshot_full_format_grid`, `snapshot_color_picker_open`, `snapshot_color_picker_hex_mode`. To regenerate after intentional UI changes:

```bash
INSTA_UPDATE=new cargo test -p vlotus      # writes .snap.new
# review .snap.new files, then rename to .snap (or use `cargo insta review`)
```

## Dependencies

lotus-core + lotus-datetime (git deps from `asg017/liblotus`), ratatui, crossterm, rusqlite, clap, arboard, thiserror, fallible-streaming-iterator, open, csv, serde_json. Dev: insta. (`csv` powers `shell::parse_csv_grid` for RFC 4180-correct parsing of `!`-shell stdout. `serde_json` backs the undo log; the `!` shell sniffer's JSON parser is hand-rolled to preserve object-key insertion order, which serde_json's default `Map` doesn't.)

## Shell paste (`!` prompt)

`src/shell.rs` is the whole feature: `run` / `run_with_cap` spawn `sh -c` (or `cmd /C` on Windows) with stdin nulled and stdout/stderr piped, capping captured stdout at 100 MB and killing the child if exceeded. A side thread drains stderr concurrently to avoid deadlocking on a full pipe. `detect_payload` sniffs JSON → NDJSON → TSV → CSV → plain (first match wins) and returns a `PastedGrid` with `source_anchor: None` and every `formula: None` — values land verbatim, never as live formulas. NDJSON reuses `json_grid_from_value` by synthesizing an `Array` of per-line values, so the array-of-objects / array-of-scalars rules are shared with the single-document path.

`App::run_shell` orchestrates: spawn → sniff → `App::insert_shell_payload(&PastedGrid)` → `apply_pasted_grid`. The split lets tests exercise paste behavior without touching subprocess plumbing. `Mode::Shell` shadows `Mode::Command` everywhere relevant: ui.rs prompt prefix + mode banner, main.rs mouse-guard / drag-guard / formula-bar title.

Status reports the **requested** paste size, not the actually-written count — same contract as the clipboard paste path. Out-of-bounds rows/cols are silently dropped by `build_paste_changes`.

JSON parser caveat: hand-rolled (`JsonParser` in `shell.rs`) to preserve insertion order via `Vec<(String, JsonValue)>` for objects. serde_json was rejected because its default `Map` is `BTreeMap`-backed (sorted keys), and `preserve_order` changes the shape of the type globally. Limited grammar — strings (with `\u{XXXX}` escapes, no surrogate pairs), numbers (preserved as the source literal), bool/null, arrays, objects. Nested arrays/objects in cell values re-serialize via `json_serialize` → compact JSON.

## Hyperlinks

`src/hyperlink.rs` owns three things:
1. `looks_like_url` — scheme-list detection (`http://`, `https://`,
   `mailto:`, `ftp://`, `ftps://`, `file://`); render-time hint only,
   no engine state.
2. `open_url` — wraps the `open` crate (macOS `/usr/bin/open`, Linux
   `xdg-open`, Windows `cmd /c start`). `cfg(test)` short-circuits to
   `Ok(())` so the suite never spawns a real browser.
3. `HyperlinkType` + `HyperlinkFn` — a vlotus-local custom type +
   function pair registered on every `Sheet::new` site (see
   `store/cells.rs::recalculate`). `=HYPERLINK(url, label)` produces
   a `CustomValue` whose `data` packs `url + '\u{1F}' + label`. The
   unit-separator (`hyperlink::SEP`) survives SQLite TEXT storage and
   round-trips through `recalculate` unchanged. `App::displayed_for`
   strips the URL prefix so the grid shows the label;
   `App::url_for_cell` returns the URL portion so `go` / ctrl-click
   open the right thing. Keep `lotus-url` and the engine untouched —
   HYPERLINK is intentionally not exposed to lotus-wasm / lotus-pyo3.

When adding a new vlotus-local custom type or function, register it
alongside `hyperlink::register` (and the cfg-gated
`datetime::register`) in `store/cells.rs::recalculate` so
both the TUI and `vlotus eval` paths see it (both go through
`Store::recalculate`).

## Datetime extension

Optional `datetime` Cargo feature, on by default for the native binary
(`cargo build -p vlotus --no-default-features` opts out). Wires
`lotus-datetime` (six `j*` types + ~40 functions) onto every
`Sheet::new` site via `crate::datetime::register(&mut sheet)`, sitting
next to `crate::hyperlink::register` in `store/cells.rs::recalculate`.
The vlotus side of the integration is intentionally tiny — handlers
and functions live in the `lotus-datetime` crate; `src/datetime.rs`
just re-exports `register` and a `is_datetime_tag` predicate over the
six tag constants from `lotus_datetime::tags`.

Display: `recalculate` writes `sheet.registry().display(cv)` into
`cell.computed` for non-hyperlink Custom values, so `jspan` renders
as `1y 2mo 3d` and `jdatetime` swaps `T` for space. The hyperlink
type is carved out with a comment because its `display()` returns
just the label, but `App::displayed_for` / `url_for_cell` decode the
encoded `url + SEP + label` payload from `cell.computed`. Eventual
follow-up: source URLs from `cell.raw` and drop the carve-out.

Auto-style: `theme::DATETIME_FG` (peach), painted in `ui::draw_grid`
in a branch parallel to the hyperlink underline. `App::datetime_tag_for_cell`
is the probe; it reads from `Store::custom_cells`, an in-memory
`HashMap<(sheet, row, col), CustomCell>` repopulated on every
`recalculate` (zero on-disk schema). `CustomCell` carries both the
type tag and the canonical `cv.data` so the strftime override
(`:fmt date`) can call `lotus_datetime::format_custom_value` without
re-parsing the cell text.

Format axis: `format::CellFormat::date: Option<String>` is a per-cell
strftime pattern persisted as the `df` JSON key in
`cell_format.format_json`. Applied by `App::apply_date_format` (called
from `displayed_for` before `format::render`). No-op for cells whose
type isn't a datetime, so attaching a pattern to a text cell is
harmless until that cell becomes a datetime.

Live-time entry: `App::insert_today_literal` / `insert_now_literal`
type the current ISO date / datetime as a literal (not a formula),
matching Excel/Sheets `Ctrl+;` / `Ctrl+Shift+;`. They call
`lotus_datetime::today_iso` / `now_iso` (gated on jiff's `_has-now`
feature, which `system-tz` enables on the native binary). The literal
auto-detects as `jdate` / `jdatetime`.

Engine probe: `Sheet::type_tag(&str) -> Option<&str>` (added to
`lotus-core`) lets any consumer holding a live Sheet ask "what type
tag is at this cell?". vlotus snapshots this into `Store::custom_cells`
during recalc since its Sheet is short-lived.

Tutor: L16 ("Dates") is a cfg-gated lesson appended to `LESSONS` in
`tutor.rs`. The `seed_creates_one_sheet_per_lesson` test counts
`LESSONS.len()` so it tracks automatically.

## Storage

`src/store/` owns SQLite-backed persistence. `Store::open` writes the
schema (`sheet`, `cell`, `cell_format`, `column_meta`, `meta`,
`undo_entry` — composite natural PKs throughout) and stamps
`application_id = 0x564C4F54` ("VLOT") + `user_version = 1`, then
opens the long-lived dirty-buffer txn (`BEGIN IMMEDIATE`). All edits
land inside that txn until `:w` (commit) or `:q!` (rollback). The
Drop guard rolls back any pending buffer; SQLite's WAL also discards
on next open if the process is killed.

Public methods: `apply` / `load_sheet` / `recalculate` /
`get_computed` (cells), `list_sheets` / `create_sheet` /
`delete_sheet` (sheets), `load_columns` / `set_column_width`
(columns), `insert_rows` / `delete_rows` / `insert_cols` /
`delete_cols` (structural ops), `record_undo_group` / `pop_undo` /
`apply_redo` / `clear_undo_log` (undo log).

## Patch sessions (`.lpatch` files)

`src/store/patch.rs` wires SQLite's session extension into the
store. `:patch new <file>` starts an `sqlite3_session` attached to
the user-data tables (filter excludes `undo_entry` / `meta`); every
authored mutation accumulates as a changeset. `:patch save` writes
`Session::changeset_strm` to the file (tmp + rename for atomicity);
`:patch close` saves and stops; `:patch detach` drops without
saving; `:patch invert` applies the inverse to the live workbook
and resets the session; `:patch pause`/`:patch resume` toggle
recording mid-session.

The session is held inside `Store` next to the `Connection` via a
documented `unsafe { mem::transmute }` lifetime cast — sound because
the `patch` field is declared first (Drop ordering) and
`Store::drop` releases the session before the conn. `:q!`
(`Store::rollback`) invalidates an active patch since rollback
diverges the changeset from the workbook.

`Store::apply` / structural ops bracket their `recalculate()` call
in `with_session_disabled` so `cell.computed` / `cell.owner_*`
writes don't bloat patches.

CLI surface in `src/patch_cli.rs`: `vlotus patch apply <db>
<patch>` (with optional `--invert` and `--on-conflict
omit|replace|abort`), `vlotus patch show <patch>` (renders
`Sheet1!A1: "foo" → "bar"` lines), `vlotus patch invert <in>
<out>`, `vlotus patch combine <out> <patch>...`,
`vlotus patch diff <db-from> <db-to> <out>` (uses
`session::Session::diff` against an `ATTACH`ed source DB to
derive a patch from two workbook snapshots).

`:patch apply <file>` lands an external patch into the live
dirty buffer for review (user `:w` to commit, `:q!` to discard).
`:patch break <new-file>` chunks a long recording into multiple
patch files. `:patch show` opens a modal popup
(`Mode::PatchShow`, state on `App::patch_show`); j/k scroll, q /
Esc close.

## Undo / redo

Undo state lives on disk in the `undo_entry` table inside the
dirty-buffer txn — every cell-edit / column-resize call site captures
the inverse via `App::snapshot_undo_ops` and writes a group via
`Store::record_undo_group`. `App::undo` calls `Store::pop_undo` and
pushes the popped `UndoGroup` onto an in-memory `redo_stack`;
`App::redo` pops from the redo stack and calls `Store::apply_redo`,
which re-records the inverses so subsequent `u` walks back the other
way. Cap is `UNDO_DEPTH = 50` distinct groups (oldest pruned in
`record_undo_group`). `Store::commit` clears the log so each save
window starts fresh; `Store::rollback` discards uncommitted entries
along with the rest of the buffer.
