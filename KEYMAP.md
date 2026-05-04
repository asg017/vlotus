# vlotus Vim Keymap

Reference for every key, command, and prefix in vlotus. Each entry includes:
- **What** — user-facing behavior, including non-obvious edge cases.
- **How** — the App method or main.rs handler that implements it.
- **Tests** — unit-test names in `examples/vlotus/src/app.rs::tests` (or noted otherwise).

Conventions:
- `{N}` is a numeric count prefix (digits in Normal mode accumulate into `App::pending_count`).
- `{a-z}` is a single ASCII letter.
- `Ctrl+X` is `KeyCode::Char('x')` with `KeyModifiers::CONTROL`.
- "Cell A1" = row 0, col 0 in 0-indexed coordinates; status bar shows the 1-indexed cell ID.

---

## Modes

| Mode | Trigger | Banner | Notes |
|------|---------|--------|-------|
| `Mode::Nav` | default; `Esc` from anywhere | `-- NORMAL --` (cyan) | Vim Normal — letters are commands, not text. |
| `Mode::Edit` | `i` / `a` / `I` / `A` / `o` / `O` / `s` / `S`; `Enter` / `F2` | `-- INSERT --` (yellow) | `edit_buf` holds the staged value. |
| `Mode::Visual(Cell)` | `v` from Normal | `-- VISUAL --` (magenta) | Free rectangle. |
| `Mode::Visual(Row)` | `V` from Normal | `-- V-LINE --` (magenta) | Selection columns pinned 0..=`MAX_COL`. |
| `Mode::Visual(Column)` | `VV` from Normal; column-header click | `-- V-COLUMN --` (magenta) | Selection rows pinned 0..=`MAX_ROW`. |
| `Mode::Command` | `:` from Normal/Visual | `-- COMMAND --` (yellow) | Ex prompt; `edit_buf` holds the typed line. |
| `Mode::Shell` | `!` from Normal | `-- SHELL --` (yellow) | Shell prompt; `edit_buf` holds the command. Enter runs it via `shell::run` and pastes the sniffed output at the cursor. |
| `Mode::Search(SearchDir)` | `/` (forward) / `?` (backward) | `-- SEARCH --` (yellow) | Pattern prompt. |

Run-loop dispatcher lives in `main.rs::run_loop`; one `handle_*_key` per mode (Nav and Visual share `handle_nav_key` because most motions are mode-aware, not mode-specific).

`Esc` is universal cancel: clears every `pending_*` flag (count / g / z / f / operator / mark / ctrl_w), exits Visual, and dismisses the clipboard mark.

Two-keystroke prefixes:
- `g…` — motion / nav (`gg`, `gv`, `gt`, `gT`, `go`)
- `z…` — viewport (`zz`, `zt`, `zb`, `zh`, `zl`)
- `f…` — **format** (number / style / align / color axes — see "Cell formatting" below)
- `m{a-z}` / `'{a-z}` / `` `{a-z} `` — marks
- `Ctrl+w …` — column widths

---

## Counts

A leading digit in Normal mode accumulates into `App::pending_count`. Bare `0` is the "go to column 0" motion **unless** a count is already in progress, in which case it's a digit. Most motions multiply by the count; absolute motions (`G`, `{N}gt`) treat it as a target.

| App method | Tests |
|------------|-------|
| `App::consume_count()` | `consume_count_takes_and_resets` |
| `App::clear_pending_motion_state()` | `clear_pending_motion_state_clears_all_prefix_flags` |

---

## Motions — Normal & Visual

In Visual mode every motion below extends the selection rectangle. The cursor moves; the anchor stays put. `move_cursor` / `jump_cursor_to` / `move_cursor_jump` all check `matches!(self.mode, Mode::Visual(_))` to skip the anchor reset.

### Cursor (one-cell)

| Key | What | Counts | Method | Tests |
|-----|------|--------|--------|-------|
| `h` | left | yes | `move_cursor(0, -1)` | (motion path) |
| `j` | down | yes | `move_cursor(1, 0)` | |
| `k` | up | yes | `move_cursor(-1, 0)` | |
| `l` | right | yes | `move_cursor(0, 1)` | |
| Arrow keys | same as hjkl | no | `move_cursor` | |
| `Tab` / `Shift+Tab` | right / left (also commits in Insert) | no | `move_cursor` | |
| `Shift+Arrow` | extend selection | no | `move_cursor_selecting` | |
| `Ctrl+Arrow` | jump to filled-run boundary | no | `move_cursor_jump` | `jump_*` (5 tests) |
| `Ctrl+Shift+Arrow` | extend to boundary | no | `move_cursor_jump_selecting` | |

### Word (filled-run boundary)

Spreadsheet "word" = contiguous run of filled cells in the cursor's row. Forward motions are inclusive of the target cell; spreadsheets simplify vim's exclusive-on-forward distinction.

| Key | What | Counts | Method | Tests |
|-----|------|--------|--------|-------|
| `w` | forward to start of next run | yes | `word_forward` → `compute_word_forward` | `word_forward_walks_past_run_then_skips_empties`, `word_forward_stays_at_edge_with_no_more_content` |
| `b` | back to start of current/prev run | yes | `word_backward` → `compute_word_backward` | `word_backward_lands_on_run_start`, `word_backward_at_col_zero_is_noop` |
| `e` | end of current/next run | yes | `word_end` → `compute_jump_target(0,1)` | covered by `jump_*` tests |

### Row-internal absolute

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `0` | column 0 | `move_to_row_start` → `jump_cursor_to(row, 0)` | (covered indirectly) |
| `^` | first filled cell in row, else col 0 | `move_to_first_filled_in_row` | `first_filled_in_row_finds_left_edge_of_content`, `first_filled_in_empty_row_falls_back_to_col_zero` |
| `$` | last filled cell in row, else `MAX_COL` | `move_to_last_filled_in_row` | `last_filled_in_row_finds_right_edge_of_content` |

### Whole-sheet

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `gg` | row 0 of current column | `goto_first_row` (after `pending_g` consumed) | `goto_first_row_and_last_filled_row` |
| `{N}gg` | row N (1-indexed) of current column | `goto_row(N-1)` | `goto_row_clamps_to_max_row` |
| `G` | **last row in the sheet** with any data; row 0 if empty | `goto_last_filled_row` | `goto_first_row_and_last_filled_row`, `goto_last_filled_row_on_empty_column_uses_global_max`, `goto_last_filled_row_on_empty_sheet_lands_at_row_0` |
| `{N}G` | row N (1-indexed) | `goto_row(N-1)` | (above) |
| `{` | previous "paragraph" in current column (start of preceding filled run) | `paragraph_backward` | `paragraph_motions_skip_to_next_run_start` |
| `}` | next "paragraph" in current column (end of following filled run) | `paragraph_forward` | (above) |
| `Ctrl+a` | expand selection to the contiguous data region — the bounding box of the 8-connected block of populated cells around the cursor. On an empty cell, selects the entire sheet (`A1:Z1000`). Mode-agnostic; lands in `Visual::Cell` regardless of prior mode. | `select_data_region` | `select_data_region_on_block_enters_visual_with_correct_rect`, `select_data_region_on_empty_cell_selects_entire_sheet`, `select_data_region_on_isolated_cell_is_one_by_one`, `select_data_region_from_v_line_re_anchors_as_visual_cell` (helper: `compute_data_region` + 6 unit tests) |

### Viewport (window-relative)

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `H` | top of viewport | `move_to_viewport_top` | `viewport_motions_use_scroll_window` |
| `M` | middle of viewport | `move_to_viewport_middle` | (above) |
| `L` | bottom of viewport | `move_to_viewport_bottom` | (above) |
| `Ctrl+d` | half page down (count repeats) | `scroll_half_page(1)` | `scroll_half_page_moves_cursor_by_half_visible_rows` |
| `Ctrl+u` | half page up | `scroll_half_page(-1)` | (above) |
| `Ctrl+f` | full page down | `scroll_full_page(1)` | `scroll_full_page_clamps_at_grid_top` |
| `Ctrl+b` | full page up | `scroll_full_page(-1)` | (above) |
| `zz` | center cursor row vertically (no cursor move) | `scroll_cursor_to_middle` | `z_prefix_scrolls_keep_cursor_put` |
| `zt` | cursor row → top of viewport | `scroll_cursor_to_top` | (above) |
| `zb` | cursor row → bottom | `scroll_cursor_to_bottom` | (above) |
| `zh` | scroll viewport one column left (cursor stays put) | `scroll_viewport_left` | |
| `zl` | scroll viewport one column right | `scroll_viewport_right` | |

`H` / `M` / `L` and the `z*` prefix don't take a count.

---

## Insert mode entries (Normal → Edit)

Every entry sets `Mode::Edit` and seeds `edit_buf` differently. Esc cancels (preserves original cell value); Enter commits and moves down; Tab/Shift+Tab commit and move right/left.

| Key | What | Method | Notes |
|-----|------|--------|-------|
| `i` / `I` | caret at start of existing value | `start_edit_at_start` | Aliases — cells are one line. |
| `a` / `A` | caret at end of existing value | `start_edit` | Aliases. |
| `o` | move down 1, Insert with empty buffer | `move_cursor(1,0)` + `start_edit_blank` | Doesn't pre-clear cell content; commit overwrites. |
| `O` | move up 1, Insert with empty buffer | `move_cursor(-1,0)` + `start_edit_blank` | |
| `s` / `S` | Insert with empty buffer at current cell | `start_edit_blank` | **Subtle:** `s<Esc>` preserves the original value (nothing committed). `cc<Esc>` (operator path, see below) does NOT preserve it because `apply_operator(Change)` clears the rect before opening Insert. |
| `=` | Insert with `=` already seeded — jumpstart formula authoring | `start_edit_blank` + `edit_insert('=')` | `equals_in_nav_starts_formula_edit`. Spreadsheet convention (Excel / Sheets / sc-im). Falls through when `Ctrl+w` prefix is pending so `Ctrl+w =` reaches the column-width consumer. |
| `Enter` / `F2` | Insert with caret at end (same as `a`) | `start_edit` | Universal-spreadsheet UX. |

While in Edit mode, the autocomplete popup steals `Up` / `Down` / `Tab` / `Enter` / `Esc` when open. V9 added `Ctrl+n` / `Ctrl+p` (next/prev), `Ctrl+y` (accept), `Ctrl+e` (dismiss) as vim-style supplements.

Tests: `start_edit_*` paths are covered indirectly by `confirm_edit_records_insert_for_dot_repeat`, `cancel_edit_downgrades_change_to_delete`, `autocomplete_*` tests.

### While editing a formula — pointing mode

When the buffer starts with `=` and the caret sits at an "insertable" position (right after `=`, an operator `+ - * / ^ & < >`, `(`, or `,`, with the next non-whitespace char being EOL, `)`, or `,`), arrow keys insert a cell ref instead of moving the caret. Any non-arrow key exits the pointing session. Mouse equivalents (click / drag a cell or row/column header during formula edit) are documented in the **Mouse** section below as T8–T13.

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `Arrow` (Up/Down/Left/Right) — first press at insertable caret | insert a cell ref pointing at `cursor + delta` (e.g. `=1+` then Down → `=1+B2`) and start a Cell-kind pointing session | `App::try_pointing_arrow(dr, dc, false)` | `pointing_arrow_inserts_ref_at_insertable_caret`, `pointing_arrow_clamps_at_grid_edge`, `pointing_does_not_start_in_middle_of_ref` |
| `Arrow` — subsequent press while pointing | replace the inserted ref with a fresh single-cell ref at `target + delta` (e.g. `=1+B2` then Down → `=1+B3`); collapses any active range and resets the anchor | `App::try_pointing_arrow(dr, dc, false)` → `replace_pointing_ref` | `pointing_arrow_moves_existing_ref` |
| `Shift+Arrow` while pointing (Cell kind) | extend the ref into a range with the anchor pinned at the original cell (`=1+B2` then Shift+Down → `=1+B2:B3`); stepping back through the anchor collapses then re-extends the other side | `App::try_pointing_arrow(dr, dc, true)` → `rewrite_pointing_text` | `shift_arrow_during_pointing_extends_range`, `shift_arrow_back_through_anchor_collapses_then_extends_other_side`, `plain_arrow_during_pointing_still_collapses_after_shift_extension`, `shift_arrow_with_no_pointing_starts_pointing_like_plain_arrow`, `shift_arrow_at_non_insertable_caret_falls_through` |
| Any non-arrow key | exit pointing (typed char/Backspace/Enter/Esc/Tab still do their normal edit-mode thing) | `App::exit_pointing` | `non_arrow_key_exits_pointing` |

Insertable-position rules live in `is_insertable_at` (`app.rs`): inside a string, after a `)`, or in the middle of an identifier are all non-insertable, so `="hello "` + arrow doesn't insert a ref and `=A1` + arrow doesn't either.

---

## Visual mode

`v` enters cell-rectangle visual; `V` enters row-visual (V-LINE — cols pinned 0..=`MAX_COL` in `selection_range`); pressing `V` again from V-LINE promotes to V-COLUMN (rows pinned 0..=`MAX_ROW`), so `VV` from Normal is the keyboard route to whole-column selection. Anchor = cursor at entry; subsequent motions extend the rect. `gv` re-enters the most-recent visual range (saved on `exit_visual` to `App::last_visual`). Column-header click also enters V-COLUMN directly (mouse path, see Mouse section).

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `v` (Normal) | enter Visual::Cell | `enter_visual(VisualKind::Cell)` | `enter_visual_anchors_at_cursor` |
| `V` (Normal) | enter Visual::Row | `enter_visual(VisualKind::Row)` | `vline_selection_pins_columns_to_full_grid` |
| `VV` (Normal) | enter Visual::Column (V-LINE → V-COLUMN cycle) | `switch_visual_kind(VisualKind::Column)` | `shift_v_cycles_nav_then_v_line_then_v_column_then_nav` |
| `gv` (Normal) | re-enter last Visual | `reselect_last_visual` | `reselect_last_visual_restores_rect`, `reselect_last_visual_returns_false_when_none_saved` |
| `o` (Visual) | swap anchor and cursor | `swap_visual_corners` | `swap_visual_corners_swaps_anchor_and_cursor` |
| `v` (Visual) | promote Row/Column → Cell, or exit if already Cell | `switch_visual_kind` / `exit_visual` | `motion_in_visual_extends_rectangle` |
| `V` (Visual) | cycle Cell→Row, Row→Column, Column→exit | `switch_visual_kind` / `exit_visual` | `shift_v_cycles_nav_then_v_line_then_v_column_then_nav`, `shift_v_from_v_cell_promotes_to_v_line` |
| `Esc` (Visual) | exit to Normal (saves last_visual) | `exit_visual` | `exit_visual_saves_last_visual` |
| `y` (Visual) | yank selection, exit | `visual_yank` (main.rs) → `yank_selection`, OS clipboard sync | `yank_selection_round_trips_through_paste` |
| `d` / `x` (Visual) | clear selection + yank, exit | `visual_delete` → `clear_rect` | `clear_rect_clears_every_cell_in_one_undo_step` |
| `c` (Visual) | clear selection, drop into Insert at top-left | `visual_change` | (above) |
| Insert-mode entries (`i` `I` `a` `A` `O` `s` `S`) | noop in Visual | dispatcher absorbs and consumes count | |

In V-LINE, motions that move horizontally still reposition the cursor visually (status bar tracks col), but `selection_range` always returns full-row rects. V-COLUMN is symmetric: vertical motions move the cursor but `selection_range` returns full-column rects.

---

## Operators (V4)

Vim's `{op}{motion}` grammar. Pressing `d` / `c` / `y` in Normal sets `pending_operator` and stashes any `pending_count` into `pending_op_count`. The next key is interpreted by `motion_target_for_operator` (in `main.rs`), which walks the same motion table without mutating the cursor and returns a target. `apply_operator` builds the inclusive rect from cursor to target and acts.

Counts compose: `5d3w` → `pending_op_count = 5`, then `3` → `pending_count = 3`, then `w` → resolver runs `w` 5×3=15 times and clears the resulting rect.

Doubled letter (`dd` / `cc` / `yy`) operates on the cursor cell only (per locked-in decision; vim normally treats them as line-wise).

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `d{motion}` | clear cells in motion rect (yanks first) | resolver → `clear_rect`, captures into register | `clear_rect_clears_every_cell_in_one_undo_step` |
| `dd` | clear current cell | resolver doubled path | (above) |
| `D` | clear cursor → end of row (alias `d$`) | direct call to `apply_operator(Delete, ...)` | |
| `c{motion}` | clear motion rect, drop into Insert at top-left | resolver → `clear_rect` + `start_edit_blank` | `cancel_edit_downgrades_change_to_delete` |
| `cc` | doubled — clear current cell + Insert | resolver doubled path | |
| `C` | clear cursor → end of row + Insert | direct call | |
| `y{motion}` | yank cells in motion rect (no clear) | resolver → `build_grid_over` | `yank_selection_round_trips_through_paste` |
| `yy` / `Y` | yank current cell | resolver doubled path / direct | (Note: V3 originally had `Y` = yank-row; V4 changed it for vim parity. Use `Vy` for row yank.) |
| `x` | clear current cell in place. `Nx` clears N cells to the right, cursor stays at start | `clear_rect(r,c,r,c)` per cell | `lowercase_x_clears_in_place_with_count` |
| `X` | move left, then clear | symmetric | |
| `p` / `P` | paste from register at cursor (falls back to OS clipboard) | `paste_from_register_or_clipboard` (main.rs) | `paste_from_empty_register_is_noop` |

Yank register also syncs to the OS clipboard via `sync_register_to_os_clipboard` / `sync_rect_to_os_clipboard`, so external paste sees the same data.

`Delete` / `Backspace` (in Normal) also clear the cursor cell (`delete_cell`) and record `last_edit` — covered by `delete_cell_records_repeatable_action`.

---

## Undo / redo / dot-repeat (V6)

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `u` | undo | `undo` | `undo_round_trip_single_cell`, `undo_chain_of_edits_replays_in_reverse`, `undo_stack_capped_at_max_depth` |
| `Ctrl+r` | redo | `redo` | `new_action_clears_redo_branch` |
| `Ctrl+z` / `Ctrl+y` / `Ctrl+Shift+z` | undo / redo (legacy bindings, kept for muscle memory) | same | |
| `.` | repeat last change at cursor | `repeat_last_edit` | `delete_cell_records_repeatable_action`, `confirm_edit_records_insert_for_dot_repeat`, `cancel_edit_downgrades_change_to_delete`, `repeat_last_edit_returns_false_when_none` |

`last_edit: Option<EditAction>` records {kind, anchor offset, rect dims, optional text}. `apply_operator` records Delete/Change shape; `confirm_edit_advancing` fills in Change text on commit, or records a fresh Insert for non-operator inserts. Yank doesn't record (vim's `.` ignores yank).

Limitation: `.` after `o foo<Esc>` writes saved text at the current cursor — it doesn't replicate the move-down-then-insert step.

---

## Search & marks (V5)

### Search

Operates on each cell's **computed** display value (`get_display`). Smart-case is *not* implemented; default is case-insensitive.

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `/{pattern}<Enter>` | forward search | `start_search(Forward)` → `commit_search` | `search_finds_first_match_after_cursor_and_wraps` |
| `?{pattern}<Enter>` | backward search | `start_search(Backward)` | (above) |
| `n` | next match (in original direction) | `search_step(Forward)` | `search_step_backward_with_capital_n_reverses` |
| `N` | previous match (reverses original direction) | `search_step(Backward)` | (above) |
| `*` | search forward for cursor cell's exact value | `search_current_cell(Forward)` | `search_current_cell_uses_displayed_value` |
| `#` | search backward for cursor cell's exact value | `search_current_cell(Backward)` | (above) |
| `:noh` / `:nohlsearch` | clear search highlight | `clear_search` | `clear_search_drops_state` |

Match state lives in `App::search: Option<SearchState>`. `ui::draw_grid` paints matching cells with `Color::LightYellow` background.

While the search prompt is open (`Mode::Search`), `handle_search_key` mirrors `handle_command_key`: Enter commits, Esc cancels, empty Backspace dismisses. Status bar swaps to `/{buf}` or `?{buf}` with the terminal cursor inside.

### Marks

Per-letter cell bookmarks. `pending_mark` is set by `m` / `'` / `` ` ``; the next letter consumes it.

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `m{a-z}` | set mark to cursor | `set_mark` | `marks_set_and_jump` |
| `` `{a-z} `` | jump to exact cell | `jump_to_mark(letter, false)` | (above) |
| `'{a-z}` | jump to row, column 0 | `jump_to_mark(letter, true)` | (above) |
| (any non-letter after prefix) | cancel | dispatcher drops `pending_mark` | |

Status reports `"Mark not set: 'z"` when an unset letter is queried (`jump_to_unset_mark_reports_status`).

---

## Sheets / tabs (V7)

Layered on the multi-sheet workbook (T5a). Bindings + Ex aliases all route to `add_sheet` / `next_sheet` / `prev_sheet` / `switch_sheet` / `delete_active_sheet`.

| Key / Cmd | What | Method | Tests |
|-----------|------|--------|-------|
| `gt` | next sheet | `next_sheet` (after `pending_g` consumed) | `next_and_prev_sheet_wrap` |
| `gT` | prev sheet | `prev_sheet` | (above) |
| `{N}gt` | switch to sheet N (1-indexed, absolute) | `switch_sheet(N-1)` | `switch_sheet_isolates_cell_data`, `switch_sheet_clears_clipboard_mark` |
| `Ctrl+PageDown` / `Ctrl+PageUp` | next / prev sheet (legacy) | same | |
| `:tabnew [name]` / `:tabe` / `:tabedit` | alias for `:sheet new` | `add_sheet` | `add_sheet_appends_and_switches` |
| `:tabnext` / `:tabn` | alias for `gt` | `next_sheet` | |
| `:tabprev` / `:tabp` / `:tabprevious` / `:tabN` | alias for `gT` | `prev_sheet` | |
| `:tabfirst` / `:tabfir` / `:tabrewind` / `:tabr` | jump to first sheet | `switch_sheet(0)` | `tabfirst_jumps_to_first_sheet` |
| `:tablast` / `:tabl` | jump to last sheet | `switch_sheet(len - 1)` | `tablast_jumps_to_last_sheet` |
| `:tabclose` / `:tabc` | alias for `:sheet del` | `delete_active_sheet` | `delete_active_sheet_drops_and_falls_back`, `delete_active_sheet_refuses_when_only_one` |
| `:tabs` | alias for `:sheet ls` | (inline in `run_command`) | |

When the workbook has more than one sheet, a one-row tabline below the grid (above the status bar, Excel / Sheets convention) lists every sheet as ` {name} `, separated by `│`. The active tab renders in reverse-video; inactive tabs in `Color::DarkGray`. If the row overflows the terminal width, neighboring tabs are kept around the active one and elided runs are flagged with leading / trailing `…`. Names longer than 16 characters are truncated with `…`. Hidden when only one sheet exists. Backed by `tabline_window` in `ui.rs` (tested by `tabline_window_*` unit tests; rendering covered by the `snapshot_multi_tab_workbook_second_active` and `snapshot_tabline_overflow_with_active_in_view` snapshots).

---

## Shell paste (`!` prompt)

`!` from Normal opens a yellow `-- SHELL --` prompt. Enter runs the typed command through `sh -c` (or `cmd /C` on Windows), captures stdout, and pastes the sniffed grid at the cursor. **The cursor row becomes the header row; data fills below.** The whole paste is a single undo group (inherited from `apply_pasted_grid`).

Sniffer dispatch order — first match wins:

1. **JSON** — leading non-whitespace is `[` or `{` and the input parses end-to-end. Recognised shapes:
   - `[{...}, {...}, ...]` — keys of the **first object** (insertion order) become headers; subsequent objects are rows. Missing keys → empty cells. Scalar values stringified (`"true"` / `"false"` / `""` for null / number literal preserved). Nested arrays/objects re-serialize as compact JSON.
   - `{...}` — single object, treated as a 1-element array of itself.
   - `[scalar, scalar, ...]` — single column with synthetic `value` header.
   - Mixed-shape arrays return `None` and fall through to the next sniffer.
2. **NDJSON / JSON Lines** — same gate (input starts with `[` or `{`), but tries one JSON value per line. Common output of `jq -c`, `gh api --paginate`, `kubectl get -o json | jq -c '.items[]'`. Each line must parse end-to-end as JSON; blank lines are skipped. Same shape rules as the regular JSON path applied to the synthesized array of per-line values: all-objects → first line's keys as headers; all-scalars → synthetic `value` header; mixed shapes fall through. Single-line inputs are deferred to the regular JSON branch.
3. **TSV** — first non-empty line contains `\t` and `\t`-count ≥ `,`-count.
4. **CSV** — line contains `,`. RFC 4180-correct via the `csv` crate (quoted fields, `""` escapes, embedded newlines).
5. **Plain** — split lines, single column with synthetic `value` header.

Status messages:

| Outcome | Status |
|---------|--------|
| Success | `!: pasted {requested_rows}×{requested_cols}` (the **requested** size, not the clamped count — same contract as the clipboard paste path) |
| Empty stdout / unparseable | `!: empty output` |
| Non-zero exit | `!: exit {N}: {stderr}` (stderr truncated to 200 chars, `...` appended on overflow) |
| Spawn failure | `!: spawn failed: {io_error}` |
| Stdout > 100 MB | `!: output too large ({bytes}; cap 104857600)` (child is killed) |

v1 limitations:

- **Non-interactive only.** Stdin is `null`; the TUI alt-screen stays intact while the subprocess runs but no tty pass-through is attempted. Don't run `vim`, `less`, `psql` (without `-c`), etc.
- **Blocks the UI** for the duration of the subprocess. There's no spinner or cancellation — long commands lock vlotus until they exit. Pipe through `head` / `timeout` if needed.
- **Stderr discarded on exit 0.** A noisy 0-exit command's stderr disappears.

Implementation: `src/shell.rs` (`run` / `run_with_cap` / `detect_payload` / hand-rolled JSON parser). `App::run_shell` in `src/app.rs` orchestrates; `App::insert_shell_payload` is the subprocess-free entry point used by tests. Backing tests live in `shell::runner_tests` (cfg(unix), 9 tests), `shell::parser_tests` (18 tests), and `app::tests::shell_*` / `app::tests::run_shell_*` (7 tests).

---

## Hyperlinks

Cells whose computed text is a recognised URL (`http://`, `https://`, `mailto:`, `ftp://`, `ftps://`, `file://`) auto-render with underline + cyan foreground — the standard "links are blue" affordance. Detection lives in `hyperlink::looks_like_url`; styling layers on top of the type / format / state stack in `ui::draw_grid` and is skipped under cursor / selection / search highlights so contrast still holds.

A second flavour comes from the `=HYPERLINK(url, label)` formula: the engine produces a `"hyperlink"` custom value whose `data` packs `url` and `label` separated by a single `\u{001F}` (ASCII unit-separator). `App::displayed_for` strips the URL prefix so the grid shows the label; `App::url_for_cell` returns the URL portion so `go` and ctrl-click follow the right link. The custom type and function are registered locally on every `Sheet::new` site (`store/cells.rs::recalculate`); `lotus-url`, `lotus-wasm`, and `lotus-pyo3` are untouched, so HYPERLINK is a vlotus-only feature.

| Key / Action | What | Method | Tests |
|--------------|------|--------|-------|
| `go` | open URL in cell under cursor in system browser; status flashes "Opened ..." or "No link under cursor" (mnemonic: "go [to link]") | `App::url_under_cursor` + `hyperlink::open_url` (after `pending_g` consumed) | `go_with_url_under_cursor_flashes_opened_status`, `go_with_no_url_flashes_no_link_status`, `go_on_empty_cell_flashes_no_link_status`, `bare_o_without_pending_g_still_opens_below`, `url_under_cursor_returns_url_portion_for_hyperlink_cell` |
| Ctrl+click on URL cell | move cursor to clicked cell + open URL; falls through to plain selection if no URL | inline branch in `handle_mouse` Down(Left) | `ctrl_click_on_url_cell_opens_url`, `ctrl_click_on_non_url_cell_falls_through_to_select` |
| `=HYPERLINK(url)` / `=HYPERLINK(url, label)` | formula — produces a hyperlink-typed cell. The grid shows `label` (or `url` when omitted/empty); `go` and ctrl-click follow the URL. Numeric / boolean labels are stringified. | `hyperlink::HyperlinkFn` | `one_arg_uses_url_as_label`, `two_arg_packs_url_and_label`, `empty_label_falls_back_to_url`, `numeric_label_stringifies`, `register_makes_hyperlink_callable_in_a_formula`, `concat_of_hyperlink_uses_label`, `hyperlink_formula_round_trips_through_storage` |

Production builds open via the cross-platform `open` crate (macOS `/usr/bin/open`, Linux `xdg-open`, Windows `cmd /c start`). `cfg(test)` short-circuits the dispatch so the test suite never spawns a real browser. Render-side coverage: `hyperlink::tests`, `snapshot_hyperlink_autostyle`, `snapshot_hyperlink_function_renders_label`.

---

## Dates and times

Compiled in by default via the `datetime` Cargo feature (build with `cargo build -p vlotus --no-default-features` to opt out). The `lotus-datetime` extension gets registered on every `Sheet::new` site alongside `hyperlink::register` (`store/cells.rs::recalculate`); six new cell types — `jdate`, `jtime`, `jdatetime`, `jzoned`, `jtimezone`, `jspan` — and ~40 formula functions become available.

Auto-detect on cell entry (priority order from `lotus_datetime::register`):

- `2025-04-27T12:30[America/New_York]` → `jzoned`
- `2025-04-27T12:30:45` → `jdatetime`
- `2025-04-27` → `jdate`
- `12:30` / `12:30:45` → `jtime`
- `1y 2mo 3d`, `P1Y` → `jspan`
- `jtimezone` is parse-on-request only (would otherwise shadow strings like `UTC`)

Auto-style: cells whose computed value is one of the six types render with `theme::DATETIME_FG` (peach) — the analog of the hyperlink sapphire. Layered after explicit cell-format `fg` and skipped under cursor / selection / search highlights. Probe lives in `App::datetime_tag_for_cell`, backed by `Store::custom_cells` (an in-memory map repopulated on every `recalculate`; no schema change).

Storage round-trip: `cell.raw` keeps the user's typed string; `cell.computed` keeps the friendly display form via `sheet.registry().display(cv)` (jspan `1y 2mo 3d`, jdatetime space-separated). The hyperlink type is carved out in `recalculate` because its display strips the URL the click handler needs.

Search: `/` and `?` test the cell's `raw_value` (not `computed`), so `/2025-04-27` finds typed ISO literals; the friendly span form `/1y` doesn't false-match the canonical `P1Y`.

| Key / Action | What | Method | Tests |
|---|---|---|---|
| Type ISO date / time / datetime / span literal | parsed at ingestion, displayed canonical | `lotus_datetime::*Handler::parse_literal` | `parse_literal_claims_iso_date_in_full_pipeline`, `jdate_round_trips_through_storage`, `jspan_renders_friendly_in_computed`, `jdatetime_renders_with_space_separator` |
| `=DATE(y, m, d)` / `=TIME(h, m, s)` / `=DATETIME(...)` / `=NOW()` / `=TODAY()` | construct typed values; NOW/TODAY use the system zone | upstream | upstream |
| `=A1 - A2` (two dates) | yields a `jspan`; renders friendly | upstream | upstream |
| `=A1 + B1` (date + span) | yields a `jdate` | upstream | upstream |
| `=YEAR(d)` / `=MONTH(d)` / `=DAY(d)` / `=WEEKDAY(d)` etc. | accessors | upstream | upstream |
| `=FORMAT(d, "%a %b %d")` | strftime-format any datetime value | upstream | upstream |
| `:today` | type today's date as ISO literal at cursor (lands as `jdate`); pinned in time | `App::insert_today_literal` (calls `lotus_datetime::today_iso`) | `today_command_inserts_current_date_literal` |
| `:now` | type current local datetime as ISO literal (lands as `jdatetime`) | `App::insert_now_literal` (calls `lotus_datetime::now_iso`) | `now_command_inserts_current_datetime_literal` |
| `Ctrl+;` | same as `:today` (Excel/Sheets binding) | `KeyCode::Char(';') if ctrl && !shift` in `handle_nav_key` | `today_command_inserts_current_date_literal` (binding shares the App method) |
| `Ctrl+Shift+;` (or `Ctrl+:`) | same as `:now` (terminals vary on which form they emit, both bound) | `KeyCode::Char(';') if ctrl && shift` and `KeyCode::Char(':') if ctrl` | as above |
| `:fmt date <strftime>` | per-cell strftime override; renders datetime cells through the pattern (e.g. `:fmt date %a %b %d` → `Sun Apr 27`). Stored as `df` in `cell_format.format_json`. No-op for non-datetime cells. | `apply_format_update(\|f\| f.date = Some(...))` then `App::displayed_for` → `apply_date_format` → `lotus_datetime::format_custom_value` | `fmt_date_applies_strftime_to_jdate_cell`, `fmt_date_no_op_for_non_datetime_cell`, `fmt_date_persists_through_format_json_round_trip`, `json_round_trip_date_only` |
| `:fmt nodate` | clear the per-cell strftime override | `apply_format_update(\|f\| f.date = None)` | `fmt_nodate_clears_strftime_override` |

`=TODAY()` and `=NOW()` re-evaluate against the system clock on every recalc — a workbook saved on Monday and re-opened on Tuesday silently shifts. For a pinned date use `:today` / `Ctrl+;`. The native build links jiff with `system-tz`, which uses `/etc/localtime` (or the platform equivalent) for the local zone.

---

## Ex commands (the `:` prompt)

Multi-token commands live in `App::run_command`; pure-token commands are parsed by `execute_command` (no `&mut self` needed). Unknown commands report `"Unknown command: :foo"` in the status bar.

| Command | What | Tests |
|---------|------|-------|
| `:q` / `:quit` | quit. Refused with a status warning if the dirty buffer has uncommitted edits (`has_unsaved_changes`) or this is a touched in-memory session (`should_warn_in_memory`). Tutor sessions skip the in-memory guard; ephemeral sessions (`:memory:` / `tutor`) skip the unsaved-changes guard. | `command_quit_aliases`, `quit_in_memory_*`, `quit_in_tutor_session_skips_in_memory_warning`, `quit_with_dirty_buffer_is_refused`, `quit_file_backed_dirty_buffer_is_refused_then_allowed_after_w` |
| `:q!` / `:quit!` | rollback the dirty buffer and quit unconditionally. Invalidates any active `:patch new` recording. | `command_quit_bang_aliases_force`, `quit_bang_overrides_*` |
| `:wq` / `:x` | commit + quit. Same guards as `:q`. | `wq_in_memory_touched_is_refused` |
| `:wq!` / `:x!` | commit + quit unconditionally | `command_quit_bang_aliases_force`, `wq_bang_forces_quit_in_memory` |
| `:w` / `:write` | commit the dirty-buffer txn (`COMMIT; BEGIN IMMEDIATE`) and clear the undo log; with a path argument writes a CSV/TSV instead | `export_sheet_writes_csv`, `quit_file_backed_dirty_buffer_is_refused_then_allowed_after_w` |
| `:42` (or any pure-numeric) | jump cursor to row N (1-indexed) | `jump_to_target_handles_row_and_cell` |
| `:A1` (any cell-id) | jump to that cell | (above) |
| `:goto <target>` / `:cell <target>` | long form of `:42` / `:A1` | `jump_to_target_handles_row_and_cell` |
| `:colwidth N` | set current column to N chars | `handle_colwidth_parses_numeric_and_auto` |
| `:colwidth <letter> N` | set specific column | (above) |
| `:colwidth auto` / `:colwidth <letter> auto` | autofit | `autofit_column_uses_longest_displayed_value`, `autofit_empty_column_uses_default` |
| `:noh` / `:nohlsearch` | clear search highlight | `clear_search_drops_state` |
| `:sheet new [name]` | create + switch | `add_sheet_appends_and_switches` |
| `:sheet del` / `:sheet delete` | delete active | `delete_active_sheet_*` |
| `:sheet rename <name>` / `:sheet ren <name>` | rename active sheet (FK cascade rewrites cells + formats; clears undo log) | `rename_active_sheet_*`, `sheet_rename_command_runs_through_run_command` |
| `:sheet ls` / `:sheet list` | list (active wrapped in `[brackets]`) | |
| `:row ins above|below` / `:row insert above|below` | structural row insert | `insert_row_above_cursor_shifts_content_down` |
| `:col ins left|right` / `:col insert left|right` | structural column insert | `insert_col_below_right_keeps_cursor_col_intact` |
| `:row del` | structural row delete (y/n confirm) | `delete_row_requires_confirm_before_executing`, `cancel_pending_confirm_leaves_data_intact` |
| `:col del` | structural column delete (y/n confirm) | `delete_col_with_formula_dependent_rewrites_to_ref_error` |
| `:reset` | (tutor only) reseed the active lesson | (no test; covered by tutor seed tests indirectly) |
| `:next` / `:prev` / `:previous` | aliases for `gt` / `gT` (tutor convenience) | |
| `:help` | show a one-line keymap reminder in the status bar | |
| `:set mouse` | enable mouse capture (click-to-select, drag, scroll, …). On by default; this command is mostly useful to re-enable after `:set nomouse`. | `set_mouse_re_enables_capture` |
| `:set nomouse` | disable mouse capture so the terminal's native text-selection (Cmd/Option+drag → copy) works. | `set_nomouse_disables_capture` |
| `:set mouse?` | report current state in the status bar (`mouse=on` / `mouse=off`) | `set_mouse_query_reports_current_state` |
| `:fmt usd [N]` | apply USD currency format to selection (default 2 decimals; `N` is 0–10) | `fmt_usd_attaches_format_to_active_cell`, `fmt_usd_with_explicit_decimals_parses`, `fmt_usd_rejects_bad_decimals` |
| `:fmt percent [N]` | apply percent format (default 0 decimals). `4.5%` displays as such; raw stored as 0.045. | `fmt_percent_attaches_format`, `fmt_percent_with_decimals_renders_correctly` |
| `:fmt+` / `:fmt-` | bump decimals on selection up/down. On unformatted cells, applies USD/2 first then bumps (gsheets toolbar parity). | `fmt_plus_minus_bumps_decimals`, `fmt_plus_on_unformatted_applies_usd_then_bumps` |
| `:fmt {bold,italic,underline,strike}` / `:fmt no…` | toggle a text-style flag across selection. Flags compose with each other and with number/align/color. | `fmt_bold_attaches_flag`, `fmt_nobold_clears_flag`, `fmt_style_flags_compose` |
| `:fmt {left,center,right}` / `:fmt auto` | set explicit alignment override. `auto` (or `noalign`) clears it so classify_display picks (number = right, boolean = center, text = left). | `fmt_left_overrides_classify_right_align`, `fmt_center_and_right_round_trip`, `fmt_auto_clears_explicit_alignment` |
| `:fmt fg <color>` / `:fmt bg <color>` | set foreground / background color. Accepts a Catppuccin preset name (red, blue, orange…) or hex (`#rgb` or `#rrggbb`). | `fmt_fg_named_color`, `fmt_bg_hex_color`, `fmt_fg_bad_color_reports_status` |
| `:fmt nofg` / `:fmt nobg` | clear just the named color axis; preserve every other axis. | `fmt_nofg_clears_only_fg` |
| `:fmt clear` / `:fmt none` | drop EVERY format axis on selection (number + style + align + colors). Raw values preserved. | `fmt_clear_drops_format`, `fmt_clear_drops_style_flags_too` |
| `:patch new <file>` | start recording authored edits into `<file>` (creates an SQLite session attached to `cell` / `cell_format` / `column_meta` / `sheet`; `undo_entry` and `meta` are excluded). Status bar shows ` [patch: <basename>] `. | `patch_records_authored_changes_only` |
| `:patch save` / `:patch save --commit` / `:patch save --rollback` | flush the changeset to the patch file. Default keeps the dirty buffer; `--commit` follows with `:w`; `--rollback` follows with `:q!` (shelve). | `patch_records_authored_changes_only` |
| `:patch close` | save (keep-dirty) and stop recording | `patch_records_authored_changes_only` |
| `:patch detach` | stop recording WITHOUT saving — discards the in-memory changeset | (no test) |
| `:patch invert` | apply the inverse of the recorded changeset to the workbook (rolls authoring back to start of recording); the patch session resets so further edits start fresh | `patch_invert_round_trips_authored_state` |
| `:patch pause` / `:patch resume` | toggle recording without dropping the session — edits while paused don't appear in the patch | `patch_pause_resume_brackets_edits_out` |
| `:patch status` | report current patch path / paused state in the status bar | (no test) |
| `:patch apply <file>` | replay an external patch into the dirty buffer (review then `:w` to commit, `:q!` to discard). Bracketed by `with_session_disabled` so the apply doesn't fold into an active recording. | `patch_apply_command_lands_changes_into_dirty_buffer` |
| `:patch break <new-file>` | save current patch (keep-dirty), close it, open a fresh patch at `<new-file>` — for chunking a long edit session into bite-sized patches | (no test) |
| `:patch show` | open a modal popup rendering the in-progress changeset (`Sheet!A1: + "value"` style). j/k scroll, gg/G top/bottom, q / Esc close. | (no test) |

---

## Cell formatting

`format::CellFormat` is a struct with composable axes — number / style flags / alignment / fg / bg colors. Every axis is independent: a cell can be USD-formatted **and** bold **and** center-aligned **and** red, all at once. Applies via `apply_format_update(F)` which reads → mutates → writes, so axis-specific commands preserve every other axis. Format JSON lives in the `cell_format.format_json` column (joined into `Store::load_sheet` via LEFT JOIN on `(sheet_name, row, col)`).

### Number formats

| Key / Cmd | What | Tests |
|-----------|------|-------|
| `f$` | apply USD/2 to selection (mnemonic "**f**ormat dollar") | `f_dollar_applies_usd_format` |
| `f%` | apply Percent/0 to selection | `f_percent_applies_percent_format` |
| `f.` | bump decimals +1 (auto-applies USD/2 first on unformatted cells) | `f_dot_increases_decimals_f_comma_decreases` |
| `f,` | bump decimals −1 | (above) |
| `:fmt usd [N]` | apply USD with N decimals | `fmt_usd_attaches_format_to_active_cell` |
| `:fmt percent [N]` | apply Percent with N decimals | `fmt_percent_attaches_format` |
| `:fmt+` / `:fmt-` | bump decimals on whichever number format is active | `fmt_plus_minus_bumps_decimals` |
| (auto-detect at edit commit) | `$1.25` → 1.25 + USD/2; `4.5%` → 0.045 + Percent/1; `-$1,234.56` and `-3%` work too. `=` formulas untouched. | `currency_input_parses_and_formats_round_trip`, `percent_input_auto_detects_at_edit_commit`, `formula_with_dollar_unchanged` |

Negative renders: `-$1.25` and `-3.0%` (gsheets default). `classify_display` recognises both `$`/`-$` and `N%`/`-N%` as `DisplayKind::Number` for right-alignment.

### Text styles

| Key / Cmd | What | Tests |
|-----------|------|-------|
| `fb` | toggle **b**old | `fb_toggles_bold` |
| `fi` | toggle **i**talic | `fi_fu_fs_toggle_their_flags` |
| `fu` | toggle **u**nderline | (above) |
| `fs` | toggle **s**trikethrough | (above) |
| `:fmt bold` / `:fmt nobold` | explicit on/off | `fmt_bold_attaches_flag`, `fmt_nobold_clears_flag` |
| `:fmt italic` / `:fmt noitalic` | (likewise) | (above) |
| `:fmt underline` / `:fmt nounderline` | | (above) |
| `:fmt strike` / `:fmt nostrike` | | (above) |

Render: ratatui `Modifier::{BOLD,ITALIC,UNDERLINED,CROSSED_OUT}` layers on top of the state-driven cell style, so a bold cell stays bold under cursor / selection / search highlights.

### Alignment

| Key / Cmd | What | Tests |
|-----------|------|-------|
| `fl` | **l**eft-align | `fl_fc_fr_set_alignment` |
| `fc` | **c**enter-align | (above) |
| `fr` | **r**ight-align | (above) |
| `fa` | **a**uto-align (clear explicit override) | `fa_clears_alignment` |
| `:fmt left/center/right/auto` | command equivalents | `fmt_left_overrides_classify_right_align`, `fmt_center_and_right_round_trip`, `fmt_auto_clears_explicit_alignment` |

Lowercased letters are safe under `f`-prefix because the consumer block runs *before* the operator/Nav-only dispatch in `handle_nav_key`. Bare `c` (Change operator), `a/A` (Insert at end), and `l` (motion right) carry defensive `!pending_f` guards but the consumer would intercept them either way.

`Option<Align>` — `None` falls back to `classify_display` (number = right, boolean = center, text/error = left). Explicit alignment wins and persists across number-format changes.

### Colors

| Key / Cmd | What | Tests |
|-----------|------|-------|
| `fF` | open **F**oreground color picker (modal `Mode::ColorPicker`) | `fF_opens_color_picker_for_fg` |
| `fB` | open **B**ackground color picker | `fB_opens_color_picker_for_bg` |
| `:fmt fg <color>` | set fg from preset name or `#rgb` / `#rrggbb` hex | `fmt_fg_named_color`, `fmt_fg_bad_color_reports_status` |
| `:fmt bg <color>` | set bg | `fmt_bg_hex_color` |
| `:fmt nofg` / `:fmt nobg` | clear that axis only | `fmt_nofg_clears_only_fg` |

In the picker: `hjkl` navigate the 4×6 swatch grid (l/h ±1 cell, j/k ±row), `Enter` apply, `Esc` cancel, `?` toggle hex-input mode, hex chars + Enter apply a typed color. See `picker_navigates_and_applies_swatch`, `picker_hex_input_applies_typed_color`, `picker_esc_cancels_without_applying`.

Preset palette: Catppuccin Mocha names (Mauve / Peach / Sky / etc.) plus natural-language aliases (purple / orange / cyan / etc.) — see `format::COLOR_PRESETS`.

Render order: format `fg` always wins (a red-text cell stays red even under selection bg); format `bg` defers to state-driven highlights so cursor / selection / search bgs still paint correctly.

### Clear-all

`:fmt clear` (alias `:fmt none`) wipes EVERY axis at once — number, style, align, fg, bg. Raw value is preserved. Sends the literal `"null"` sentinel which `set_cells_and_recalculate` interprets via a non-COALESCE upsert branch (drops `format_json` to NULL). See `fmt_clear_drops_format`, `fmt_clear_drops_style_flags_too`.

### Persistence + undo

Format JSON layout: `{"n":{"k":"usd","d":2},"b":true,"a":"center","fg":"#ff0000","bg":"#1e1e2e"}`. Unset fields are omitted. Hand-rolled parser/emitter (no serde dep).

`:fmt …` undo: applying a format to a previously-unformatted cell, then `u` undoing, must actually drop the format. `capture_undo_entry` writes the `"null"` clear-sentinel for that case; `Store::apply` interprets the sentinel as "delete the cell_format row" (versus `None` which preserves the existing format). See `fmt_undo_restores_unformatted_state`.

### Snapshots

- `snapshot_currency_formatted_grid` — number formats + alignment fallbacks
- `snapshot_color_picker_open` — picker overlay in swatch mode
- `snapshot_color_picker_hex_mode` — picker after `?` toggle

---

## Column widths (variable-width columns)

`column_meta.width` is interpreted as **character count** by vlotus. `App::column_width(col_idx)` is the canonical accessor; defaults to `DEFAULT_COL_WIDTH = 12` for columns without a row, clamped to `MIN_COL_WIDTH..=MAX_COL_WIDTH` (1..=80).

| Key / Cmd | What | Method | Tests |
|-----------|------|--------|-------|
| `Ctrl+=` | autofit current column — or every column in the active Visual::Cell / Visual::Column selection | `autofit_selection_or_column` → `autofit_column` per col | `ctrl_equals_in_nav_autofits_column`, `ctrl_equals_in_v_column_autofits_every_selected_column`, `autofit_column_uses_longest_displayed_value`, `autofit_empty_column_uses_default` |
| `Ctrl+w =` | autofit current column (terminal-portable fallback for `Ctrl+=`); same multi-column behavior in Visual selections | `autofit_selection_or_column` | `ctrl_w_equals_autofits_as_fallback`, `ctrl_w_equals_in_v_column_autofits_every_selected_column` |
| `Ctrl+w >` (count-multiplied) | widen current column by N | `set_column_width(col, current+N)` | `set_column_width_persists_and_clamps` |
| `Ctrl+w <` (count-multiplied) | narrow current column by N | `set_column_width(col, current-N)` | (above) |
| `:colwidth N` / `:colwidth <letter> N` | set absolute width | `handle_colwidth` → `set_column_width` | `handle_colwidth_parses_numeric_and_auto` |
| `:colwidth auto` / `:colwidth <letter> auto` | fit to longest displayed value + 1 char pad | `autofit_column` | `autofit_column_uses_longest_displayed_value` |

Note: bare `=` no longer autofits — it now jumpstarts formula authoring (see "Insert mode entries" below). Use `Ctrl+=` (or `Ctrl+w =` if your terminal doesn't deliver that combo) for autofit.

Renderer: `ui::draw_grid` walks per-column widths to compute `visible_cols` (stops when the next column wouldn't fit) and renders each column at its declared width. Each visible column is charged `width + COL_SPACING` (where `COL_SPACING = 1` matches ratatui's `Table::column_spacing`); `active_cell_rect` uses the same constant so the autocomplete popup anchors correctly under non-uniform layouts. Truncation handles narrow widths: width 1 = single char (no ellipsis room), width ≥ 2 = N-1 chars + `…`.

---

## Mouse

**On by default.** `:set nomouse` disables for the rest of the session if you want the terminal's native text-selection back (see `:set` rows above). When enabled, vlotus interprets these gestures inside the grid; outside the grid (formula bar, status bar, borders) they're no-ops.

| Gesture | What | Handler | Tests |
|---------|------|---------|-------|
| Left-click cell | move cursor; in Edit mode commit-and-jump (same-cell click is no-op) | `handle_left_click_cell` (main.rs) | `left_click_in_nav_moves_cursor`, `left_click_in_edit_mode_commits_and_moves`, `left_click_on_self_in_edit_mode_is_noop` |
| Left-click + drag | enter `Visual::Cell` anchored at the click point; cursor follows the drag. Drag past the visible edge auto-scrolls the viewport one row/column per Drag-event tick, advancing the cursor so the selection grows. In Visual already, anchor sticks (vim `v` + motion semantics). | `handle_drag_to_cell`, `handle_drag_past_edge` (main.rs) | `drag_in_nav_transitions_to_visual_cell_with_correct_anchor`, `drag_past_bottom_scrolls_and_extends_selection`, `drag_past_right_scrolls_and_extends_selection`, `drag_past_top_at_origin_clamps_without_panic`, `drag_past_bottom_in_v_line_keeps_column` |
| Scroll wheel up / down | scroll viewport ±1 row; cursor doesn't follow (Excel / VisiData convention). Clamped at row 0 / `MAX_ROW`. | `App::scroll_viewport(±1, 0)` | `scroll_down_advances_viewport_without_moving_cursor`, `scroll_up_at_origin_clamps_to_zero` |
| Shift + scroll, or `ScrollLeft` / `ScrollRight` | scroll viewport ±1 column. The Shift convention covers terminals that don't emit horizontal scroll natively. | `App::scroll_viewport(0, ±1)` | `shift_scroll_advances_horizontal_viewport`, `scroll_right_advances_horizontal_viewport` |
| Double-click cell | enter Edit mode at the cell with its raw value pre-loaded into the buffer (formula text, not computed value). Two clicks on the same cell within `DOUBLE_CLICK_MS = 400ms` count as a double-click. Suppressed while already in Edit mode so clicks during formula entry can't wipe the in-flight buffer. | `is_double_click` predicate + `App::start_edit` (main.rs) | `is_double_click_pure_helper`, `double_click_within_threshold_enters_edit_mode`, `double_click_is_suppressed_during_edit` |
| Click cell *during formula edit* | insert that cell's A1-style ref at the edit caret (e.g. typing `=1+`, then clicking B2 → `=1+B2`). Reuses the keyboard pointing-mode infrastructure: subsequent clicks replace the pointing-ref instead of appending. Falls through to commit-and-move if the caret isn't in formula context (no leading `=`, mid-ref, etc). | `App::insert_ref_at_caret` | `click_cell_during_formula_edit_inserts_ref_at_caret`, `click_cell_with_active_pointing_replaces_ref`, `click_during_non_formula_edit_falls_through_to_commit`, `click_at_non_insertable_caret_falls_through_to_commit` |
| Drag cells *during formula edit* | extend the active pointing-ref into a range (`=SUM(B2`, drag to B5 → `=SUM(B2:B5`). Drag back to anchor collapses to single. Range normalizes top-left → bottom-right regardless of drag direction. | `App::drag_pointing_target` + `rewrite_pointing_text` | `drag_during_formula_edit_inserts_range_ref`, `drag_normalizes_range_min_to_max`, `drag_back_to_anchor_collapses_to_single_ref` |
| Click column header *during formula edit* | insert `B:B` at the caret and start a `Column`-kind pointing session so a follow-up column-header drag extends to `B:E`. Replaces any active pointing of any kind. | `App::insert_col_ref_at_caret` | `col_header_click_during_formula_edit_inserts_col_ref`, `col_header_click_replaces_active_pointing_ref` |
| Drag column headers *during formula edit* | extend the column-range pointing — drag from B to E during edit → `B:E`. | `App::drag_pointing_target` (Column kind) | `col_header_drag_during_formula_edit_inserts_col_range`, `col_header_drag_back_collapses_range`, `col_drag_normalizes_reverse_direction` |
| Click row header *during formula edit* | insert `1:1` at the caret (1-indexed display) and start a `Row`-kind pointing session. | `App::insert_row_ref_at_caret` | `row_header_click_during_formula_edit_inserts_row_ref` |
| Drag row headers *during formula edit* | extend the row-range pointing — drag from row 1 to row 5 → `1:5`. | `App::drag_pointing_target` (Row kind) | `row_header_drag_during_formula_edit_inserts_row_range` |
| Click tab (multi-sheet) | switch the active sheet (`App::switch_sheet`). Clicks on the leading/trailing `…` truncation marker fall through as a no-op for now. | `App::switch_sheet` | `left_click_on_tab_switches_active_sheet`, `left_click_on_active_tab_is_idempotent` |
| Click row-number gutter | enter `Visual::Row` (V-LINE) at that row. V-LINE auto-spans every column via `App::selection_range`. | (inline in `handle_mouse`) | `left_click_on_row_header_enters_v_line` |
| Drag along row gutter | extend the V-LINE selection — cursor row follows the drag, column unchanged. | (inline in `handle_mouse`) | `drag_along_row_headers_extends_v_line_selection` |
| Click column header | enter `Visual::Column` (V-COLUMN) — `selection_range` forces the row extent to `0..=MAX_ROW` so yank/delete touch every row of the column without scrolling the cursor away from the user's current row. | (inline in `handle_mouse`) | `left_click_on_column_header_enters_v_column`, `vcolumn_selection_pins_rows_to_full_grid` |
| Drag along column headers | extend the cell-mode selection in the column axis — cursor column follows the drag, row unchanged. | (inline in `handle_mouse`) | `drag_along_column_headers_extends_column_selection` |

Pending epic children (`att tree i38d0opr`): the formula-edit ref-insertion variants (T8-T13).

---

## Clipboard (system-clipboard, kept alongside vim register)

These pre-date the vim work (T2a/T2b) but stay bound for muscle memory. They paint a "marching ants" perimeter (`clipboard_mark`) at copy/cut time so the user can see what's queued.

| Key | What | Method | Tests |
|-----|------|--------|-------|
| `Ctrl+c` | copy selection (or cursor cell) | `copy_selection_to_clipboard(Copy)` (main.rs) | |
| `Ctrl+x` | cut (paint mark; clears source on next paste) | `copy_selection_to_clipboard(Cut)` | `paste_with_cut_clear_is_one_undo_step` |
| `Ctrl+v` | paste OS clipboard | `paste_from_clipboard` (main.rs) | |
| `Esc` | cancel a clipboard mark | `clear_clipboard_mark` | (used in dispatcher) |

Vim `y` / `p` (above) layer on top: `y` populates *both* the internal register and the OS clipboard; `p` reads the internal register first and falls back to the OS clipboard.

---

## Tutor

`cargo run -p vlotus -- tutor` opens a 14-lesson curriculum in an in-memory workbook. Each lesson is one sheet (named `L1 Movement`, `L2 Insert`, …, `L14 Closing`); column A holds instructions readable in full via the formula bar; columns B–H hold practice data. Mutations evaporate on `:q`.

| Cmd | What | Tests |
|-----|------|-------|
| `cargo run -p vlotus -- tutor` | open the lesson workbook | `seed_creates_one_sheet_per_lesson`, `each_lesson_has_at_least_a_title_cell`, `lesson_ids_are_unique` (in `tutor::tests`) |
| `:reset` | reseed the active lesson to its starting state | |
| `:next` / `:prev` | aliases for `gt` / `gT` (discoverable for new users) | |

Lessons are defined as a `const &[Lesson]` in `examples/vlotus/src/tutor.rs`. Each lesson is a `(row, col, value)` array — easy to grep, easy to diff. Adding a new lesson: append a `Lesson { id, name, cells }` entry to `LESSONS`.

---

## UI snapshot tests

`src/snapshots.rs` covers ~14 styled UI states via `ratatui::backend::TestBackend` + `insta`. Most use plain-text snapshots; the four scenes whose feature is purely styling (visual rect, V-LINE, search match tint, clipboard mark perimeter) use `render_with_highlights` which appends a per-row range listing of cells by category (cursor / selection / mark perimeter / search match / pointing target).

Regenerating after intentional UI changes:

```bash
INSTA_UPDATE=new cargo test -p vlotus   # writes .snap.new alongside .snap
# review .snap.new files, then rename to .snap
# (or: cargo install cargo-insta && cargo insta review)
```

Snapshots live in `examples/vlotus/src/snapshots/`. Classification of a cell's highlight category is restricted to the data-grid rows (`y=5..h-2`) so the Normal-mode banner (Cyan+Black+Bold — same style as a pointing target) doesn't false-match.

---

## State on `App` (developer reference)

Every vim feature is backed by a field on `App`. Grouped by phase:

```text
// V1 (mode + entry):           mode, edit_buf, edit_cursor
// V2 (counts, prefixes):       pending_count, pending_g, pending_z
// V3 (visual + register):      yank_register, last_visual
// V4 (operators):              pending_operator, pending_op_count
// V5 (search + marks):         search, marks, pending_mark
// V6 (dot-repeat):              last_edit
// V8 (column widths):          columns (Vec<ColumnMeta>), pending_ctrl_w
```

`clear_pending_motion_state()` zeroes every `pending_*` flag in one call — used by Esc and by mid-sequence aborts.

---

## Subtle behaviors and known gaps

These deserve docs because they'll bite someone:

1. **`s<Esc>` vs `cc<Esc>`** are different. `s` opens Insert without clearing the cell (so cancel preserves the original value); `cc` goes through `apply_operator(Change)` which clears first, so cancel leaves the cell empty. Vim treats them as identical (both = "delete + insert"); we don't, currently.

2. **`Y` / `yy` yank one cell, not the row.** V3 originally had them as row-yank for vim parity, V4 changed it so doubled-letter operators are consistent (`dd` = cell, so `yy` = cell). For row yank, use `Vy`.

3. **Forward operator motions are inclusive on the target cell.** Vim's charwise/linewise distinction (where `dw` excludes the target) is not implemented; for spreadsheets, "delete from cursor through `w`'s landing cell" reads cleaner.

4. **`.` after `o foo<Esc>`** writes the saved text at the *current* cursor — it doesn't replicate the move-down step.

5. **`G` is global.** Bare `G` finds the last row with data anywhere in the sheet (max `row_idx` across all cells), not the last filled row in the cursor's column.

6. **Smart-case search not implemented.** Patterns are unconditionally case-insensitive.

7. **`{N}` count after the operator** combines: `5d3w` deletes 15 words. `5dgg` does *not* work — multi-key motions (`gg`, `gt`) aren't supported in operator-pending; they just cancel.

8. **`Ctrl+w` is a prefix**, not a single-shot binding. Bare `Ctrl+w` (without a follow-up) is a no-op that leaves a pending flag until the next key.

9. **Tutor `:reset` only works on tutor sheets** — sheet id has to start with `tutor-l`. Outside the tutor it reports `"Not a tutor sheet"`.

10. **Currency auto-detect peels but doesn't validate as a money type.** `$1.25` becomes raw `1.25` + USD/2 format; `=A1+1` in a USD-formatted cell renders as `$N` because the format renders the formula's numeric result. There's no "currency type" — only a presentation layer. Multiplying USD cells doesn't track units.

11. **`f.` reaches the f-prefix consumer cleanly.** Bare `.` is dot-repeat; the `f`-prefix consumer block sits before the early Nav-mode arms in `handle_nav_key`, so `f`-prefixed `.` is intercepted before the dot-repeat arm runs. `f$` and `f,` likewise — the consumer placement is the trick that lets every format axis use lowercase second-keys without elaborate defer guards.

12. **`:fmt` on an empty selection cell is silently skipped.** `set_cells_and_recalculate` deletes empty-raw rows, so attaching a format to a blank cell would silently disappear. Status reports "No cells to format" when the selection contains only blanks.

13. **Format axes compose; `:fmt …` merges.** `:fmt usd` on a bold cell preserves bold. `:fmt fg red` on a percent cell preserves the number format. The merge path is `apply_format_update(F)` which reads the existing format, mutates one field, and writes back. Use `:fmt clear` for the full reset.

14. **`:fmt clear` wipes EVERY axis** (number + style + align + fg + bg). Today's behavior matches Sheets's "Clear formatting"; the older single-axis interpretation is gone.

15. **Unbound `f{key}` combos cancel cleanly.** The `pending_f` consumer's catch-all clears the flag for unrecognised second-keys, so e.g. `fx` (where `x` is unbound under `f`) drops the prefix and does nothing else. `fB` opens the bg picker (F4), `fb` toggles bold — capitals and lowercase are distinct.

16. **Bare-arm defer guards.** Bare `b/i/u/s/c/a/A/l/.` carry `!pending_f` guards (defensive — pending_f consumer runs before the bare arms anyway). Bare `C` and `L` carry `!pending_g` guards — these aren't defensive, they're load-bearing: bare C/L live in the *early* Nav-only block which runs *before* the pending_g consumer, so without the guards `gC` would run change-to-end and `gL` would jump viewport-bottom instead of being absorbed as a clean no-op by the consumer's catch-all.

17. **Color picker captures selection at open-time, not apply-time.** `fF` opens the picker bound to whatever was selected when you pressed it. Moving the cursor while the picker is open doesn't change which cells get the color; Enter applies to the captured rect.

18. **Format `fg` always wins; `bg` defers.** A red-text cell stays red even under selection bg. A blue-bg cell turns whatever color the selection / cursor / search wants while highlighted, and shows blue otherwise.
