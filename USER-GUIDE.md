# vlotus user guide

User-facing reference: keybindings, commands, mouse gestures, the
bundled tutor, and the optional `datetime` feature. For the
exhaustive keymap (every key + the App method that backs it + the
unit-test names that exercise each, plus subtle behaviors), see
[KEYMAP.md](./KEYMAP.md). For build / development notes, see
[README.md](./README.md).

## Tutor

`vlotus tutor` opens a 16-lesson curriculum that walks through the vim-style
keymap from movement (L1) through a closing combined exercise (L15) and
dates / strftime / `:fmt date` (L16, included with the default `datetime`
feature). Each lesson is one sheet — read column A in the formula bar at the
top of the window, follow each step, then press `gt` to advance to the next
lesson (or `{N}gt` to jump). `:reset` reverts the active lesson to its
initial state. Mutations are kept in memory only and disappear on `:q`.

## Dates and times

The `datetime` Cargo feature (on by default) loads `lotus-datetime`,
adding six `j*` cell types (`jdate`, `jtime`, `jdatetime`, `jzoned`,
`jtimezone`, `jspan`) and ~40 formula functions. Type `2025-04-27`
into a cell — it lands as a `jdate` and renders peach. `Ctrl+;` and
`Ctrl+Shift+;` insert today's date / current datetime as literals
(Excel/Sheets convention). `:fmt date %a %b %d` overrides the display
strftime per cell; `:fmt nodate` clears it. Build without datetime via
`cargo build -p vlotus --no-default-features` for a smaller binary.
See [KEYMAP.md](KEYMAP.md#dates-and-times) for the full reference.

## Keys

The full reference (every key + command, the App method that backs each, the
unit-test names that exercise each, plus subtle behaviors and gaps) lives in
[**KEYMAP.md**](./KEYMAP.md). Quick reference follows.

| Key                       | Action                                            |
| ------------------------- | ------------------------------------------------- |
| h / j / k / l             | Move one cell (left / down / up / right)          |
| Arrows                    | Same as hjkl                                      |
| {N}{motion}               | Repeat the motion N times (e.g. `5j`, `10l`, `3w`)|
| w / b / e                 | Forward / back to next "word" (filled-run boundary)|
| 0                         | Move to column 0 of current row                   |
| ^                         | Move to first filled cell of current row          |
| $                         | Move to last filled cell of current row           |
| gg                        | First row of current column (or `{N}gg` → row N)  |
| G                         | Last row in the sheet that has data (or `{N}G` → N)|
| { / }                     | Previous / next "paragraph" (run) in column       |
| H / M / L                 | Top / middle / bottom of viewport                 |
| Ctrl+d / Ctrl+u           | Half-page down / up                               |
| Ctrl+f / Ctrl+b           | Full-page down / up                               |
| zz / zt / zb              | Re-scroll cursor to middle / top / bottom         |
| zh / zl                   | Scroll viewport one column left / right           |
| Ctrl+Arrow                | Jump to content boundary / grid edge              |
| Shift+Arrow               | Extend selection by one cell                      |
| Ctrl+Shift+Arrow          | Extend selection to content boundary              |
| Tab / Shift+Tab           | Move one column right / left (commits in edit)    |
| i / I                     | Start editing, caret at start of value            |
| a / A                     | Start editing, caret at end of value              |
| o / O                     | Move down / up and start editing with empty cell  |
| s / S                     | Clear cell and start editing                      |
| =                         | Start editing with `=` seeded (jumpstart formula) |
| v                         | Enter Visual mode (rectangular selection)         |
| V                         | Enter V-LINE mode (whole rows)                    |
| VV                        | Enter V-COLUMN mode (whole columns)               |
| gv                        | Re-select the previous Visual range               |
| x / X                     | Clear current cell / clear cell to the left       |
| d{motion}                 | Clear cells through motion (yanks to register)    |
| dd                        | Clear current cell                                |
| D                         | Clear cells from cursor to end of row             |
| c{motion}                 | Change — clear through motion + Insert            |
| cc / S                    | Clear current cell + Insert                       |
| C                         | Change to end of row                              |
| y{motion}                 | Yank cells through motion                         |
| yy / Y                    | Yank current cell                                 |
| p / P                     | Paste from register (falls back to OS clipboard)  |

In Visual / V-LINE / V-COLUMN mode (motions extend the selection):

| Key                       | Action                                            |
| ------------------------- | ------------------------------------------------- |
| Any motion (hjkl, w, …)   | Extend the selection                              |
| o                         | Swap anchor and cursor (move the other corner)    |
| y                         | Yank selection, exit to Normal                    |
| d / x                     | Clear selection, exit to Normal                   |
| c                         | Clear selection, drop into Insert at top-left     |
| v                         | Promote Row/Column → Cell, or exit if Cell        |
| V                         | Cycle Cell→Row, Row→Column, Column→exit           |
| Esc                       | Exit Visual without clearing                      |
| Enter / F2                | Start editing (or commit + move down in edit)     |
| Esc                       | Cancel edit / clear clipboard mark                |
| Delete / Backspace        | Clear cell (in nav mode)                          |
| Ctrl+C                    | Copy selection (paints clipboard mark)            |
| Ctrl+X                    | Cut selection (mark + clear on next paste)        |
| Ctrl+V                    | Paste OS clipboard (HTML table or text)           |
| u / Ctrl+Z                | Undo last cell change                             |
| Ctrl+r / Ctrl+Shift+Z / Ctrl+Y | Redo                                         |
| .                         | Repeat the last change at the cursor              |
| /pattern / ?pattern       | Forward / backward search prompt                  |
| n / N                     | Next / previous match                             |
| * / #                     | Search forward / back for current cell's value    |
| m{a-z}                    | Set a mark at the current cell                    |
| `` `{a-z} ``              | Jump to a mark's exact cell                       |
| '{a-z}                    | Jump to a mark's row, column 0                    |
| `:noh` / `:nohlsearch`    | Clear the search highlight                        |
| Ctrl+PgUp / Ctrl+PgDn     | Switch to previous / next sheet tab               |
| gt / gT                   | Next / previous sheet (vim style)                 |
| {N}gt                     | Switch to sheet N (1-indexed)                     |
| `:`                       | Open the command line                             |

Command line (`:` opens it):

| Command                   | Effect                                            |
| ------------------------- | ------------------------------------------------- |
| `:q`, `:quit`             | Quit                                              |
| `:w`                      | No-op (writes are immediate)                      |
| `:sheet new [name]`       | Create a new sheet and switch to it               |
| `:sheet del`              | Delete the active sheet (refused if it's the last)|
| `:sheet ls`               | List all sheets in the workbook                   |
| `:tabnew [name]` / `:tabe`| Alias for `:sheet new`                            |
| `:tabnext` / `:tabn`      | Alias for `gt`                                    |
| `:tabprev` / `:tabp`      | Alias for `gT`                                    |
| `:tabclose` / `:tabc`     | Alias for `:sheet del`                            |
| `:tabs`                   | Alias for `:sheet ls`                             |
| `:42`                     | Jump to row 42 (1-indexed)                        |
| `:A1` / `:goto A1`        | Jump to cell A1                                   |
| `:w <path>.csv`           | Export the active sheet to a CSV file             |
| `:w <path>.tsv`           | Export to a TSV file                              |
| `:colwidth <n>`           | Set the current column to N chars (1..=80)        |
| `:colwidth <letter> <n>`  | Set a specific column's width                     |
| `:colwidth auto`          | Fit current column to longest displayed value     |
| `:set mouse` / `:set nomouse` / `:set mouse?` | Toggle mouse capture (on by default; nomouse releases capture so the terminal's native Cmd/Option+drag → copy works) |
| `:fmt usd [N]`            | Apply USD currency format (default 2 decimals)    |
| `:fmt percent [N]`        | Apply percent format (default 0 decimals)         |
| `:fmt+` / `:fmt-`         | Bump decimals up / down on selection              |
| `:fmt {bold,italic,underline,strike}` / `:fmt no…` | Toggle text-style flags          |
| `:fmt {left,center,right,auto}` | Set explicit alignment (or clear it)        |
| `:fmt fg <name\|hex>` / `:fmt bg <name\|hex>` | Set fg / bg color (Catppuccin presets or `#rrggbb`) |
| `:fmt nofg` / `:fmt nobg` | Clear just the named color axis                   |
| `:fmt clear` / `:fmt none`| Drop EVERY format axis on selection               |
| `:help`                   | Show a short keymap reminder in the status bar    |

Cell formatting keys (all under `f`-prefix — mnemonic "**f**ormat" — since `Ctrl+B/I/U` are taken by full-page-back / Tab / half-page-up):

| Key                       | Action                                            |
| ------------------------- | ------------------------------------------------- |
| f$                        | Apply USD format to selection                     |
| f%                        | Apply percent format                              |
| f. / f,                   | Bump decimals + / -                               |
| fb / fi / fu / fs         | Toggle bold / italic / underline / strikethrough  |
| fl / fc / fr / fa         | Left / center / right / auto-align                |
| fF / fB                   | Open foreground / background color picker         |

Editing also auto-detects format from typed input: `$1.25` stores raw `1.25` + USD/2; `4.5%` stores raw `0.045` + Percent/1; `=`-prefixed formulas (incl. `=$A$1` absolute refs) pass through unchanged.

In the color picker (`fF` / `fB` open it):

| Key                       | Action                                            |
| ------------------------- | ------------------------------------------------- |
| h / j / k / l             | Navigate the swatch grid                          |
| Enter / y                 | Apply selected color                              |
| Esc / q                   | Cancel without applying                           |
| ?                         | Toggle hex-input mode (`#rgb` or `#rrggbb`)       |

Column-width keys:

| Key                       | Action                                            |
| ------------------------- | ------------------------------------------------- |
| Ctrl+=                    | Autofit current column to its longest value       |
| Ctrl+w =                  | Autofit (terminal-portable fallback for Ctrl+=)   |
| Ctrl+w >                  | Widen the current column by 1 char (count repeats)|
| Ctrl+w <                  | Narrow the current column by 1 char               |
| `:row ins above\|below`    | Insert a blank row at the cursor                  |
| `:row del`                | Delete the cursor's row (`y` to confirm)          |
| `:col ins left\|right`     | Insert a blank column at the cursor               |
| `:col del`                | Delete the cursor's column (`y` to confirm)       |

## Mouse

Mouse capture is **on by default** (matches Excel / VisiData / sc-im). `:set nomouse` releases it for the rest of the session if you want native terminal text-selection (Cmd/Option+drag → copy) back; `:set mouse` re-enables.

| Gesture                                | Effect                                                                                  |
| -------------------------------------- | --------------------------------------------------------------------------------------- |
| Click cell                             | Move cursor                                                                             |
| Double-click cell                      | Enter Edit mode with the cell's raw value pre-loaded                                    |
| Click + drag cells                     | Visual::Cell selection; drag past the visible edge auto-scrolls                         |
| Click row gutter (number column)       | V-LINE selection of that row                                                            |
| Drag along row gutter                  | Extend V-LINE to span multiple rows                                                     |
| Click column header (letter row)       | V-COLUMN selection (forces every row of that column into the range)                     |
| Drag along column headers              | Extend V-COLUMN to span multiple columns                                                |
| Scroll wheel                           | Scroll viewport vertically (cursor stays put). Shift+scroll or `ScrollLeft/Right` for horizontal |
| Click tab in tabline (multi-sheet)     | Switch active sheet                                                                     |
| Click cell *while editing a formula*   | Insert that cell's ref at the caret (`=1+` + click B2 → `=1+B2`)                        |
| Drag cells *while editing*             | Extend to a range ref (`B2:E5`)                                                         |
| Click row/column header *while editing*| Insert whole-row (`1:1`) or whole-column (`B:B`) ref                                    |
| Drag headers *while editing*           | Extend to row/column range ref (`1:5`, `B:E`)                                           |

Behaviors are tagged in source with `// [sheet.x.y]` comments matching the
spec IDs in
[datasette-sheets/specs](https://github.com/simonw/datasette-sheets/tree/main/specs).
