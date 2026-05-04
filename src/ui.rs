use ratatui::{
    layout::{Constraint, Direction, Layout, Rect},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, Clear, Paragraph, Row, Table},
    Frame,
};
use std::collections::HashMap;

use lotus_core::{
    extract_refs, formula_tokens, CompletionKind, FunctionInfo, SignatureHelp, TokenKind,
};
use crate::store::coords::to_cell_id;

use crate::app::{App, MarkAction, Mode, Operator, SearchDir, VisualKind};
use crate::theme;

/// Width reserved for the row-number gutter (e.g. " 999"). Applied as a
/// `Constraint::Length` so it doesn't grow with viewport size.
const ROW_HEADER_WIDTH: u16 = 5;
/// Cells inserted between adjacent Table columns (passed to
/// `Table::column_spacing`). Read by `visible_cols` and `active_cell_rect`
/// so all three agree on the layout.
const COL_SPACING: u16 = 1;

/// Per-frame highlight assignments for cell refs in the formula being edited.
/// See [`compute_ref_highlights`].
#[derive(Default)]
struct RefHighlights {
    /// Byte offset (in the formula string, including the leading `=`) where
    /// a ref starts → its assigned palette color. Used by the formula-bar
    /// colorizer to override the default cyan for `CellRef` / `Range`
    /// tokens with the matching ref color.
    by_start: HashMap<usize, Color>,
    /// (row, col) → assigned color, for every concrete cell covered by any
    /// ref. Whole-column / whole-row refs aren't expanded (their `cells`
    /// vec is empty); their tokens still get colorized in the formula bar
    /// via `by_start`.
    by_cell: HashMap<(u32, u32), Color>,
}

/// Walk the refs in `formula` (left-to-right) and assign each unique
/// ref-text to the next [`REF_PALETTE`] color. Returns empty highlights
/// for non-formula input. Two refs with the same text (e.g. `A1` appearing
/// twice) share a color so the user can match them visually.
fn compute_ref_highlights(formula: &str) -> RefHighlights {
    if !formula.starts_with('=') {
        return RefHighlights::default();
    }
    let refs = extract_refs(formula);
    let mut text_color: HashMap<String, Color> = HashMap::new();
    let mut by_start: HashMap<usize, Color> = HashMap::new();
    let mut by_cell: HashMap<(u32, u32), Color> = HashMap::new();
    let mut next = 0usize;
    for r in &refs {
        // Match by uppercase text so `=a1+A1` collapses to one color.
        let key = r.text.to_uppercase();
        let color = *text_color.entry(key).or_insert_with(|| {
            let c = theme::REF_PALETTE[next % theme::REF_PALETTE.len()];
            next += 1;
            c
        });
        by_start.insert(r.start, color);
        for cc in &r.cells {
            // First ref to claim a cell wins — keeps coloring stable when
            // overlapping refs (e.g. `A1` plus `A1:B2`) cover the same cell.
            by_cell.entry((cc.row, cc.col)).or_insert(color);
        }
    }
    RefHighlights { by_start, by_cell }
}

pub fn draw(f: &mut Frame, app: &mut App) {
    let area = f.area();
    let show_tabline = app.sheets.len() > 1;

    // Tabline sits below the grid (Excel / Google Sheets convention).
    let constraints: Vec<Constraint> = if show_tabline {
        vec![
            Constraint::Length(3), // formula bar
            Constraint::Min(5),    // grid
            Constraint::Length(1), // tabline
            Constraint::Length(1), // status
        ]
    } else {
        vec![
            Constraint::Length(3), // formula bar
            Constraint::Min(5),    // grid
            Constraint::Length(1), // status
        ]
    };
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints(constraints)
        .split(area);

    // [sheet.editing.formula-ref-highlight]
    // Compute once per frame and share between the formula bar (colorize
    // ref tokens) and the grid (highlight referenced cells). Outside Edit
    // mode the result is empty, so both call sites become no-ops.
    let ref_highlights = if app.mode == Mode::Edit {
        compute_ref_highlights(&app.edit_buf)
    } else {
        RefHighlights::default()
    };

    draw_formula_bar(f, app, chunks[0], &ref_highlights);
    let grid_chunk = chunks[1];
    draw_grid(f, app, grid_chunk, &ref_highlights);
    let status_chunk = if show_tabline {
        draw_tabline(f, app, chunks[2]);
        chunks[3]
    } else {
        chunks[2]
    };
    draw_status(f, app, status_chunk);

    // Edit-mode overlays render last so they sit on top of the grid.
    if app.mode == Mode::Edit {
        // [sheet.editing.formula-signature-help] sits just below the
        // formula bar so it doesn't fight with the autocomplete popup.
        if let Some(sig) = &app.signature {
            draw_signature_help(f, sig, grid_chunk);
        }
        // [sheet.editing.formula-autocomplete]
        if app.autocomplete.is_some() {
            draw_autocomplete_popup(f, app, grid_chunk);
        }
    }

    // Color picker is a modal overlay above the grid.
    if app.mode == Mode::ColorPicker && app.color_picker.is_some() {
        draw_color_picker(f, app, grid_chunk);
    }
    // Patch-show popup overlays the grid in `Mode::PatchShow`.
    if app.mode == Mode::PatchShow && app.patch_show.is_some() {
        draw_patch_show(f, app, grid_chunk);
    }
}

fn draw_formula_bar(f: &mut Frame, app: &App, area: Rect, highlights: &RefHighlights) {
    let cell_id = to_cell_id(app.cursor_row, app.cursor_col);
    let raw = app.get_raw(app.cursor_row, app.cursor_col);
    let editing = app.mode == Mode::Edit;

    // [sheet.formula-bar.live-sync] both the cell editor and the formula bar
    // read from the same `edit_buf`, so any change in one is visible in the
    // other on the next render frame.
    let prefix = format!("{cell_id}: ");
    let body = if editing {
        app.edit_buf.as_str()
    } else {
        raw.as_str()
    };

    // [sheet.editing.formula-name-coloring]
    // [sheet.editing.formula-string-coloring]
    // [sheet.editing.formula-ref-highlight]
    let mut spans: Vec<Span<'_>> = vec![Span::raw(prefix.clone())];
    spans.extend(colorize_formula(body, highlights));

    let (title, border_color) = match app.mode {
        Mode::Edit => (" Insert ", theme::MODE_EDIT),
        Mode::Visual(VisualKind::Cell) => (" Visual ", theme::MODE_VISUAL),
        Mode::Visual(VisualKind::Row) => (" V-Line ", theme::MODE_VISUAL),
        Mode::Visual(VisualKind::Column) => (" V-Column ", theme::MODE_VISUAL),
        Mode::Command => (" Command ", theme::MODE_COMMAND),
        Mode::Shell => (" Shell ", theme::MODE_COMMAND),
        Mode::Search(_) => (" Search ", theme::MODE_SEARCH),
        Mode::ColorPicker => (" Picker ", theme::MODE_COMMAND),
        Mode::PatchShow => (" Patch ", theme::MODE_VISUAL),
        Mode::Nav => (" Formula ", theme::BORDER_DEFAULT),
    };
    let block = Block::default()
        .borders(Borders::ALL)
        .title(title)
        .border_style(Style::default().fg(border_color));

    let para = Paragraph::new(Line::from(spans)).block(block);
    f.render_widget(para, area);

    // Show cursor in formula bar when editing a cell.
    if editing {
        let chars_before: usize = app.edit_buf[..app.edit_cursor].chars().count();
        let cursor_x = area.x + 1 + prefix.len() as u16 + chars_before as u16;
        let cursor_y = area.y + 1;
        f.set_cursor_position((cursor_x, cursor_y));
    }
}

/// Tokenise `text` with `lotus_core::formula_tokens` and produce a
/// styled span for each token. Non-formula input (no leading `=`) is
/// returned as a single default-styled span. The leading `=` is
/// emitted as its own default-styled span (`formula_tokens` reports
/// positions starting *after* the `=`).
///
/// When `highlights` is non-empty (Edit mode), `CellRef` / `Range` tokens
/// whose start position is in `highlights.by_start` are recolored with
/// their assigned palette color so each ref in the formula visually
/// matches the highlighted cell in the grid.
fn colorize_formula<'a>(text: &'a str, highlights: &RefHighlights) -> Vec<Span<'a>> {
    if !text.starts_with('=') {
        return vec![Span::raw(text.to_string())];
    }
    let tokens = formula_tokens(text);
    let mut out: Vec<Span<'a>> = Vec::with_capacity(tokens.len() + 2);
    let mut cursor = 0usize;
    for t in &tokens {
        let start = (t.start as usize).min(text.len());
        let end = (t.end as usize).min(text.len());
        if start > cursor {
            out.push(Span::raw(text[cursor..start].to_string()));
        }
        let style = ref_style(t.kind, start, highlights).unwrap_or_else(|| style_for(t.kind));
        out.push(Span::styled(text[start..end].to_string(), style));
        cursor = end;
    }
    if cursor < text.len() {
        out.push(Span::raw(text[cursor..].to_string()));
    }
    out
}

/// If `kind` is a ref-shaped token whose start matches a highlighted ref,
/// return the override style — bold + the ref's palette color. `None`
/// means "fall back to the default `style_for` mapping".
fn ref_style(kind: TokenKind, start: usize, highlights: &RefHighlights) -> Option<Style> {
    if !matches!(kind, TokenKind::CellRef | TokenKind::Range) {
        return None;
    }
    let color = *highlights.by_start.get(&start)?;
    Some(Style::default().fg(color).add_modifier(Modifier::BOLD))
}

/// What an `(x, y)` mouse coordinate resolves to in the vlotus UI.
/// Returned by [`cell_at`]. Clicks on the formula bar, status bar,
/// borders, or the top-left corner of the grid resolve to `None`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HitTarget {
    /// Interior of the grid: a data cell at `(row, col)`.
    Cell { row: u32, col: u32 },
    /// The row-number gutter on the left edge of the grid.
    RowHeader(u32),
    /// The column-letter header row at the top of the grid.
    ColumnHeader(u32),
    /// One of the tabs in the tabline (only present when `sheets.len() > 1`).
    Tab(usize),
}

/// Compute the same vertical layout split that [`draw`] uses, returning
/// `(grid_rect, optional tabline_rect)`. Shared between [`cell_at`] and
/// the auto-scroll path in `main.rs::handle_mouse` so they agree on
/// where the grid lives.
pub fn grid_layout(app: &App, area: Rect) -> (Rect, Option<Rect>) {
    let show_tabline = app.sheets.len() > 1;
    let constraints: Vec<Constraint> = if show_tabline {
        vec![
            Constraint::Length(3),
            Constraint::Min(5),
            Constraint::Length(1),
            Constraint::Length(1),
        ]
    } else {
        vec![
            Constraint::Length(3),
            Constraint::Min(5),
            Constraint::Length(1),
        ]
    };
    let chunks = Layout::default()
        .direction(Direction::Vertical)
        .constraints(constraints)
        .split(area);
    let grid = chunks[1];
    let tabline = if show_tabline { Some(chunks[2]) } else { None };
    (grid, tabline)
}

/// Map a screen-space `(x, y)` (in absolute terminal coordinates, matching
/// `crossterm::event::MouseEvent::{column, row}`) to a logical hit target.
///
/// Mirrors the layout split that [`draw`] computes from `area` and the
/// inner walk that [`active_cell_rect`] / [`draw_grid`] use to position
/// per-column data. Returns `None` for clicks on the formula bar, the
/// status bar, grid borders, or the top-left grid corner where the
/// row-gutter and column-header axes meet.
///
/// `area` is the full terminal area (origin `(0, 0)`); both [`draw`] and
/// `cell_at` derive the grid / tabline rects from it.
pub fn cell_at(app: &App, area: Rect, x: u16, y: u16) -> Option<HitTarget> {
    let (grid, tabline) = grid_layout(app, area);

    if let Some(t) = tabline {
        if y == t.y && x >= t.x && x < t.x + t.width {
            return tab_at(app, t, x);
        }
    }

    // Outside the grid rect → no hit (formula bar above, status / tabline below).
    if y < grid.y || y >= grid.y + grid.height || x < grid.x || x >= grid.x + grid.width {
        return None;
    }
    // Borders of the grid `Block`.
    if y == grid.y || y == grid.y + grid.height - 1 {
        return None;
    }
    if x == grid.x || x == grid.x + grid.width - 1 {
        return None;
    }

    let inside_y = y - grid.y - 1;
    let inside_x = x - grid.x - 1;

    if inside_y == 0 {
        // Column-header row. The top-left corner (gutter + first
        // column_spacing cell) doesn't belong to any column.
        if inside_x < ROW_HEADER_WIDTH + COL_SPACING {
            return None;
        }
        let col = column_at(app, inside_x)?;
        return Some(HitTarget::ColumnHeader(col));
    }

    let visible_row = (inside_y - 1) as u32;
    if visible_row >= app.visible_rows {
        return None;
    }
    let row = app.scroll_row + visible_row;

    // Lump the column_spacing cell after the gutter into the gutter so the
    // 1-cell separator isn't a dead zone.
    if inside_x < ROW_HEADER_WIDTH + COL_SPACING {
        return Some(HitTarget::RowHeader(row));
    }

    let col = column_at(app, inside_x)?;
    Some(HitTarget::Cell { row, col })
}

/// Walk visible columns to map an `inside_x` (offset past the grid's left
/// border) to a column index. The trailing `COL_SPACING` after each column
/// is claimed by that column so the 1-cell separator isn't a dead zone.
/// Returns `None` if `inside_x` falls past the last visible data column.
fn column_at(app: &App, inside_x: u16) -> Option<u32> {
    let mut start = ROW_HEADER_WIDTH + COL_SPACING;
    for c in 0..app.visible_cols {
        let col_idx = app.scroll_col + c;
        let width = app.column_width(col_idx);
        let end = start.saturating_add(width).saturating_add(COL_SPACING);
        if inside_x >= start && inside_x < end {
            return Some(col_idx);
        }
        start = end;
    }
    None
}

/// Inverse of [`tabline_window`]: given an `x` within the tabline rect,
/// find which tab segment was hit. Returns `None` for clicks on the
/// leading/trailing `…` ellipsis, the `│` separators, or padding past
/// the last visible tab.
fn tab_at(app: &App, area: Rect, x: u16) -> Option<HitTarget> {
    let labels: Vec<String> = app.sheets.iter().map(|s| segment_label(&s.name)).collect();
    let widths: Vec<usize> = labels.iter().map(|l| l.chars().count()).collect();
    let win = tabline_window(&widths, app.active_sheet, area.width as usize);

    let mut cursor = area.x;
    if win.leading {
        cursor = cursor.saturating_add(TABLINE_ELLIPSIS.chars().count() as u16);
    }
    for (rel, idx) in (win.start..win.end).enumerate() {
        if rel > 0 {
            cursor = cursor.saturating_add(TAB_SEP.chars().count() as u16);
        }
        let w = widths[idx] as u16;
        if x >= cursor && x < cursor + w {
            return Some(HitTarget::Tab(idx));
        }
        cursor = cursor.saturating_add(w);
    }
    None
}

/// Compute the screen rect of the active cell within `grid_area`. Returns
/// `None` if the cell is scrolled out of the visible viewport.
fn active_cell_rect(app: &App, grid_area: Rect) -> Option<Rect> {
    let visible_row = app.cursor_row.checked_sub(app.scroll_row)?;
    let visible_col = app.cursor_col.checked_sub(app.scroll_col)?;
    if visible_row >= app.visible_rows || visible_col >= app.visible_cols {
        return None;
    }
    // grid is rendered inside a Block::default().borders(ALL), so content
    // starts at (x+1, y+1). The first row inside is the column header,
    // then data rows.
    // Inner_x lands at the first cell of the first data column: past the
    // left border, the row-header gutter, and the column_spacing cell that
    // separates the gutter from col[scroll_col].
    let inner_x = grid_area.x + 1 + ROW_HEADER_WIDTH + COL_SPACING;
    let inner_y = grid_area.y + 1 + 1; // border + header row
    // Walk widths + spacings of every column from `scroll_col` to `cursor_col`.
    let mut x_offset: u16 = 0;
    for c in 0..visible_col {
        x_offset = x_offset.saturating_add(app.column_width(app.scroll_col + c));
        x_offset = x_offset.saturating_add(COL_SPACING);
    }
    let cursor_width = app.column_width(app.cursor_col);
    Some(Rect {
        x: inner_x + x_offset,
        y: inner_y + visible_row as u16,
        width: cursor_width,
        height: 1,
    })
}

/// Modal swatch-grid picker. Activates via `fF` / `fB`. Row layout:
/// 4 swatches per row, each `█████` painted in the preset's color.
/// Selected swatch is marked with `>` ... `<` brackets. Bottom line
/// shows the highlighted preset's name + hex (or the in-progress
/// hex buffer when `?`-toggled into hex-entry mode).
fn draw_color_picker(f: &mut Frame, app: &App, grid_area: Rect) {
    use crate::app::ColorPickerKind;
    use crate::format::COLOR_PRESETS;
    let Some(state) = app.color_picker.as_ref() else {
        return;
    };
    const COLS: usize = 4;
    const SWATCH_W: usize = 5; // chars per swatch glyph
    const SEP_W: usize = 2; // ` ` between cells, plus selection brackets
    // Row width = COLS swatches × (sep + glyph)
    let inner_w = COLS * (SWATCH_W + SEP_W) + 1;
    let rows_total = COLOR_PRESETS.len().div_ceil(COLS);
    let height = (rows_total + 4) as u16; // grid rows + 2 status + borders

    let width = (inner_w as u16 + 4).min(grid_area.width.saturating_sub(2));
    let x = grid_area.x + (grid_area.width.saturating_sub(width)) / 2;
    let y = grid_area.y + (grid_area.height.saturating_sub(height)) / 2;
    let popup_area = Rect { x, y, width, height };

    let mut lines: Vec<Line<'_>> = Vec::new();
    for r in 0..rows_total {
        let mut spans: Vec<Span<'_>> = Vec::new();
        spans.push(Span::raw(" "));
        for c in 0..COLS {
            let idx = r * COLS + c;
            if idx >= COLOR_PRESETS.len() {
                break;
            }
            let (_, color) = COLOR_PRESETS[idx];
            let selected = idx == state.cursor && state.hex_input.is_none();
            let marker_l = if selected { ">" } else { " " };
            let marker_r = if selected { "<" } else { " " };
            spans.push(Span::raw(marker_l));
            spans.push(Span::styled(
                "█".repeat(SWATCH_W),
                Style::default().fg(ratatui::style::Color::Rgb(color.r, color.g, color.b)),
            ));
            spans.push(Span::raw(marker_r));
        }
        lines.push(Line::from(spans));
    }

    // Status line: highlighted preset name + hex, or the hex-input buffer.
    let status = if let Some(hex) = state.hex_input.as_deref() {
        format!(" hex: {hex:_<7}")
    } else {
        let (name, c) = COLOR_PRESETS[state.cursor];
        format!(" {name}  #{:02x}{:02x}{:02x}", c.r, c.g, c.b)
    };
    lines.push(Line::from(Span::styled(
        status,
        Style::default().fg(theme::FG_DEFAULT),
    )));
    let hint = if state.hex_input.is_some() {
        " 0-9a-f: hex | Enter: apply | ?: swatches | Esc: cancel"
    } else {
        " hjkl: nav | Enter: apply | ?: hex | Esc: cancel"
    };
    lines.push(Line::from(Span::styled(
        hint,
        Style::default().fg(theme::FG_MUTED),
    )));

    let title = match state.kind {
        ColorPickerKind::Fg => " Foreground color ",
        ColorPickerKind::Bg => " Background color ",
    };
    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(theme::BORDER_ACCENT))
        .title(title);
    let para = Paragraph::new(lines)
        .block(block)
        .style(Style::default().bg(theme::POPUP_BG));

    f.render_widget(Clear, popup_area);
    f.render_widget(para, popup_area);
}

/// Patch-show modal popup. Renders the rendered changeset lines from
/// `App::patch_show.lines`, scrollable via j/k. Esc / q closes (key
/// handler in `main.rs::handle_patch_show_key`).
fn draw_patch_show(f: &mut Frame, app: &App, grid_area: Rect) {
    let Some(state) = app.patch_show.as_ref() else {
        return;
    };
    // Sized to fill ~80% of the grid area, capped at the longest
    // line. Scroll offset truncates from the top.
    let max_line = state
        .lines
        .iter()
        .map(|s| s.chars().count())
        .max()
        .unwrap_or(0);
    let inner_w = max_line.min(grid_area.width.saturating_sub(4) as usize);
    let inner_h = state.lines.len().min(grid_area.height.saturating_sub(4) as usize);
    let width = (inner_w as u16 + 4).max(40);
    let height = (inner_h as u16 + 4).max(6);
    let width = width.min(grid_area.width.saturating_sub(2));
    let height = height.min(grid_area.height.saturating_sub(2));
    let x = grid_area.x + (grid_area.width.saturating_sub(width)) / 2;
    let y = grid_area.y + (grid_area.height.saturating_sub(height)) / 2;
    let popup_area = Rect { x, y, width, height };

    let visible_rows = (height as usize).saturating_sub(3);
    let lines: Vec<Line<'_>> = state
        .lines
        .iter()
        .skip(state.scroll)
        .take(visible_rows)
        .map(|line| Line::from(Span::styled(line.clone(), Style::default().fg(theme::FG_DEFAULT))))
        .collect();
    let title = format!(
        " patch ({} change{}, ↕ j/k  q to close) ",
        state.lines.len(),
        if state.lines.len() == 1 { "" } else { "s" }
    );
    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(theme::BORDER_ACCENT))
        .title(title);
    let para = Paragraph::new(lines)
        .block(block)
        .style(Style::default().bg(theme::POPUP_BG));
    f.render_widget(Clear, popup_area);
    f.render_widget(para, popup_area);
}

// [sheet.editing.formula-autocomplete]
fn draw_autocomplete_popup(f: &mut Frame, app: &App, grid_area: Rect) {
    let state = match &app.autocomplete {
        Some(s) => s,
        None => return,
    };
    if state.list.items.is_empty() {
        return;
    }

    const MAX_VISIBLE: usize = 8;
    let total = state.list.items.len();
    let height = (total.min(MAX_VISIBLE) as u16) + 2; // borders

    // Width: longest label + a 4-char kind suffix (" fn", " name") + borders.
    let max_label = state
        .list
        .items
        .iter()
        .map(|it| it.label.chars().count())
        .max()
        .unwrap_or(0) as u16;
    let width = (max_label + 8).max(20).min(grid_area.width.saturating_sub(2));

    // Anchor below the active cell. Fall back to the grid's top-left when
    // the cell is off-screen.
    let cell = active_cell_rect(app, grid_area).unwrap_or(Rect {
        x: grid_area.x + 1,
        y: grid_area.y + 1,
        width: 1,
        height: 1,
    });
    let mut x = cell.x;
    let mut y = cell.y + 1;
    if x + width > grid_area.x + grid_area.width {
        x = (grid_area.x + grid_area.width).saturating_sub(width);
    }
    if y + height > grid_area.y + grid_area.height {
        // Not enough room below — render above the cell instead.
        y = cell.y.saturating_sub(height);
    }

    let popup_area = Rect {
        x,
        y,
        width,
        height,
    };

    // Window of items around the selected entry so it stays visible.
    let visible = MAX_VISIBLE.min(total);
    let window_start = state
        .selected
        .saturating_sub(visible.saturating_sub(1))
        .min(total.saturating_sub(visible));

    let lines: Vec<Line<'_>> = state
        .list
        .items
        .iter()
        .enumerate()
        .skip(window_start)
        .take(visible)
        .map(|(i, item)| {
            let kind_tag = match item.kind {
                CompletionKind::Function => "fn",
                CompletionKind::Name => "name",
            };
            let text = format!("  {:<width$} {}", item.label, kind_tag, width = max_label as usize);
            let style = if i == state.selected {
                Style::default()
                    .bg(theme::PICKER_SELECTED_BG)
                    .fg(theme::FG_DEFAULT)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(theme::FG_DEFAULT)
            };
            Line::from(Span::styled(text, style))
        })
        .collect();

    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(theme::BORDER_ACCENT))
        .title(format!(" {} match{} ", total, if total == 1 { "" } else { "es" }));
    let para = Paragraph::new(lines).block(block).style(Style::default().bg(theme::POPUP_BG));

    f.render_widget(Clear, popup_area);
    f.render_widget(para, popup_area);
}

// [sheet.editing.formula-signature-help]
fn draw_signature_help(f: &mut Frame, sig: &SignatureHelp, grid_area: Rect) {
    let line = signature_help_line(&sig.function, sig.active_param);
    let total_width = line
        .spans
        .iter()
        .map(|s| s.content.chars().count() as u16)
        .sum::<u16>()
        + 4; // borders + padding
    let width = total_width.min(grid_area.width.saturating_sub(2));

    // Pin the tooltip to the top-left of the grid area so it doesn't
    // collide with the autocomplete popup (which anchors near the cell).
    let area = Rect {
        x: grid_area.x + 1,
        y: grid_area.y,
        width,
        height: 3,
    };

    let block = Block::default()
        .borders(Borders::ALL)
        .border_style(Style::default().fg(theme::BORDER_INFO))
        .title(" signature ");
    let para = Paragraph::new(line).block(block);

    f.render_widget(Clear, area);
    f.render_widget(para, area);
}

/// Build a styled signature line like `SUM(`<b>`number1`</b>`, [number2, …])`,
/// with the active parameter rendered bold + underlined.
fn signature_help_line(info: &FunctionInfo, active: usize) -> Line<'static> {
    let mut spans: Vec<Span<'static>> = vec![Span::styled(
        format!("{}(", info.name),
        Style::default()
            .fg(theme::TOKEN_FUNCTION)
            .add_modifier(Modifier::BOLD),
    )];
    let active_param_style = Style::default()
        .fg(theme::MODE_EDIT)
        .add_modifier(Modifier::BOLD | Modifier::UNDERLINED);
    let inactive_param_style = Style::default().fg(theme::FG_MUTED);
    let mut first = true;
    for (i, p) in info.params.iter().enumerate() {
        if !first {
            spans.push(Span::raw(", "));
        }
        first = false;
        let label = if p.optional {
            format!("[{}]", p.name)
        } else {
            p.name.clone()
        };
        let style = if i == active {
            active_param_style
        } else {
            inactive_param_style
        };
        spans.push(Span::styled(label, style));
    }
    if let Some(v) = &info.variadic {
        if !first {
            spans.push(Span::raw(", "));
        }
        let label = format!("[{}, …]", v.name);
        let style = if active >= info.params.len() {
            active_param_style
        } else {
            inactive_param_style
        };
        spans.push(Span::styled(label, style));
    }
    spans.push(Span::raw(")"));
    Line::from(spans)
}

/// Classification of a cell's COMPUTED display value, used to pick the
/// right alignment + accent. Driven entirely by the display string —
/// the TUI doesn't see the engine's internal `CellValue` type tag.
///
/// See [sheet.cell.boolean], [sheet.format.numeric-align-right],
/// [sheet.format.error-color].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum DisplayKind {
    Number,
    Boolean,
    Error,
    Text,
}

fn classify_display(text: &str) -> DisplayKind {
    if text.is_empty() {
        return DisplayKind::Text;
    }
    // Engine error strings start with `#` followed by an UPPERCASE letter
    // (`#REF!`, `#DIV/0!`, `#NAME?`, `#N/A`, `#CIRCULAR! ...`, `#SPILL!`,
    // `#VALUE!`, `#SIZE!`, `#LOADING!`). Lowercase-prefixed user strings
    // like `#hashtag` stay Text.
    let mut chars = text.chars();
    if chars.next() == Some('#') && chars.next().map(|c| c.is_ascii_uppercase()).unwrap_or(false) {
        return DisplayKind::Error;
    }
    if text == "TRUE" || text == "FALSE" {
        return DisplayKind::Boolean;
    }
    if text.parse::<f64>().is_ok() {
        return DisplayKind::Number;
    }
    if crate::format::looks_numeric(text) {
        return DisplayKind::Number;
    }
    DisplayKind::Text
}

fn align_text(kind: DisplayKind, text: &str, width: usize) -> String {
    match kind {
        // [sheet.format.numeric-align-right]
        DisplayKind::Number => format!("{text:>width$}"),
        // [sheet.cell.boolean]
        DisplayKind::Boolean => format!("{text:^width$}"),
        DisplayKind::Error | DisplayKind::Text => format!("{text:<width$}"),
    }
}

/// Align with an explicit override that wins over the classify-based
/// default. Used when a cell has `format.align = Some(_)`.
fn align_text_override(align: crate::format::Align, text: &str, width: usize) -> String {
    use crate::format::Align;
    match align {
        Align::Left => format!("{text:<width$}"),
        Align::Center => format!("{text:^width$}"),
        Align::Right => format!("{text:>width$}"),
    }
}

fn type_style_for(kind: DisplayKind) -> Style {
    match kind {
        DisplayKind::Number => Style::default().fg(theme::CELL_NUMBER),
        // [sheet.cell.boolean] accent + bold
        DisplayKind::Boolean => Style::default()
            .fg(theme::CELL_BOOLEAN)
            .add_modifier(Modifier::BOLD),
        // [sheet.format.error-color]
        DisplayKind::Error => Style::default()
            .fg(theme::CELL_ERROR)
            .add_modifier(Modifier::BOLD),
        DisplayKind::Text => Style::default().fg(theme::CELL_TEXT),
    }
}

fn style_for(kind: TokenKind) -> Style {
    match kind {
        // [sheet.editing.formula-string-coloring]
        TokenKind::String => Style::default().fg(theme::TOKEN_STRING),
        // [sheet.editing.formula-name-coloring]
        TokenKind::Name => Style::default().fg(theme::TOKEN_NAME),
        TokenKind::CellRef | TokenKind::Range => Style::default().fg(theme::TOKEN_CELLREF),
        TokenKind::Function => Style::default()
            .fg(theme::TOKEN_FUNCTION)
            .add_modifier(Modifier::BOLD),
        TokenKind::Number => Style::default().fg(theme::TOKEN_NUMBER),
        TokenKind::Boolean => Style::default()
            .fg(theme::TOKEN_BOOLEAN)
            .add_modifier(Modifier::BOLD),
        _ => Style::default(),
    }
}

fn draw_grid(f: &mut Frame, app: &mut App, area: Rect, highlights: &RefHighlights) {
    // Calculate how many columns fit by walking per-column widths.
    // Stops as soon as the next column wouldn't fit. This is the layout
    // that `active_cell_rect` and the scroll math agree on — uniform
    // `area.width / COL_WIDTH` would produce a different visible_cols
    // for non-uniform widths.
    let available_width = area.width.saturating_sub(ROW_HEADER_WIDTH + 2); // borders
    // Each visible data column costs its width + 1 column_spacing cell that
    // separates it from the previous constraint (row_header for the first
    // col, the previous data col otherwise). Without this accounting the
    // cumulative spacings push the constraint sum past the Table's inner
    // width and ratatui's solver shrinks the widest column — clipping
    // right-aligned numeric values like "100" → "10".
    let mut visible_cols: u32 = 0;
    let mut consumed: u16 = 0;
    loop {
        let w = app.column_width(app.scroll_col + visible_cols);
        let cost = w.saturating_add(COL_SPACING);
        if consumed.saturating_add(cost) > available_width {
            break;
        }
        consumed = consumed.saturating_add(cost);
        visible_cols += 1;
    }
    let visible_cols = visible_cols.max(1);
    let visible_rows = (area.height.saturating_sub(3)) as u32; // header row + borders
    app.visible_rows = visible_rows;
    app.visible_cols = visible_cols;

    // Per-column widths in display order (used by both header + data).
    let col_widths: Vec<u16> = (0..visible_cols)
        .map(|c| app.column_width(app.scroll_col + c))
        .collect();

    // Column headers
    let mut header_cells = vec![Cell::from(" ")];
    for (c, &width) in col_widths.iter().enumerate() {
        let col_idx = app.scroll_col + c as u32;
        let letter = to_cell_id(0, col_idx);
        // Strip the row number to get just the column letter
        let col_letter: String = letter.chars().take_while(|ch| ch.is_ascii_alphabetic()).collect();
        let style = if col_idx == app.cursor_col {
            Style::default()
                .fg(theme::MODE_EDIT)
                .add_modifier(Modifier::BOLD)
        } else {
            Style::default().fg(theme::FG_MUTED)
        };
        header_cells.push(Cell::from(Span::styled(
            format!("{col_letter:^width$}", width = width as usize),
            style,
        )));
    }
    let header = Row::new(header_cells)
        .style(Style::default())
        .height(1);

    // [sheet.cell.spill] pre-compute the set of cells that *anchor* a spill
    // (any other cell points back to them via spill_anchor_*). One pass per
    // frame; the cells vec is small.
    let spill_anchors: std::collections::HashSet<(u32, u32)> = app
        .cells
        .iter()
        .filter_map(|c| match (c.spill_anchor_row, c.spill_anchor_col) {
            (Some(r), Some(c)) => Some((r, c)),
            _ => None,
        })
        .collect();

    // Data rows
    let mut rows = Vec::new();
    for r in 0..visible_rows {
        let row_idx = app.scroll_row + r;
        let row_num = row_idx + 1; // 1-based display
        let row_style = if row_idx == app.cursor_row {
            Style::default()
                .fg(theme::MODE_EDIT)
                .add_modifier(Modifier::BOLD)
        } else {
            Style::default().fg(theme::FG_MUTED)
        };
        let mut cells = vec![Cell::from(Span::styled(
            format!("{row_num:>4}"),
            row_style,
        ))];

        for (c, &col_w) in col_widths.iter().enumerate() {
            let col_idx = app.scroll_col + c as u32;
            let is_cursor = row_idx == app.cursor_row && col_idx == app.cursor_col;
            let is_selected = app.is_selected(row_idx, col_idx);

            let display = app.displayed_for(row_idx, col_idx);
            let truncated = if col_w >= 2 && display.chars().count() > col_w as usize - 1 {
                let take = (col_w as usize).saturating_sub(2);
                let mut t: String = display.chars().take(take).collect();
                t.push('…');
                t
            } else if col_w < 2 {
                // Width 1 has no room for an ellipsis. Just take the
                // first char.
                display.chars().next().map(String::from).unwrap_or_default()
            } else {
                display.clone()
            };

            let mark_perimeter = app
                .clipboard_mark
                .as_ref()
                .map(|m| m.on_perimeter(row_idx, col_idx))
                .unwrap_or(false);
            // Only Cell-kind pointing has a meaningful single target;
            // Column / Row pointing tracks an axis range, not a cell.
            let pointing_target = app
                .pointing
                .as_ref()
                .map(|p| {
                    matches!(p.kind, crate::app::PointingKind::Cell)
                        && p.target_row == row_idx
                        && p.target_col == col_idx
                })
                .unwrap_or(false);

            // [sheet.cell.spill] member is owned by another anchor's spill.
            let spill_member = app
                .cells
                .iter()
                .find(|sc| sc.row_idx == row_idx && sc.col_idx == col_idx)
                .map(|sc| sc.spill_anchor_row.is_some())
                .unwrap_or(false);
            let spill_anchor = spill_anchors.contains(&(row_idx, col_idx));

            let cell_fmt = app.get_format(row_idx, col_idx);
            let kind = classify_display(&display);
            let mut base = type_style_for(kind);
            // [sheet.cell.spill] members render italic + dim (TUI fallback
            // for the spec's "italic, secondary text colour"). Anchors get
            // bold + the cell-ref accent — the closest TUI equivalent of
            // the spec's left-edge accent without sacrificing display width.
            if spill_member {
                base = base.add_modifier(Modifier::ITALIC | Modifier::DIM);
            }
            if spill_anchor {
                base = base.fg(theme::SPILL_ANCHOR_FG).add_modifier(Modifier::BOLD);
            }

            // V5: tint cells that match the active search pattern. Lower
            // priority than cursor / selection / mark, but above the
            // auto-classification.
            let search_match = app
                .search
                .as_ref()
                .map(|s| s.matches(&display))
                .unwrap_or(false);

            // [sheet.clipboard.mark-visual]
            // [sheet.editing.formula-ref-pointing]
            // [sheet.editing.formula-ref-highlight]
            // State-driven highlights win over the auto-classification.
            // Pointing target outranks ref-highlight: the cell being
            // actively pointed at is the *new* ref the user is inserting.
            let ref_color = highlights.by_cell.get(&(row_idx, col_idx)).copied();
            let mut style = if is_cursor && app.mode == Mode::Edit {
                Style::default()
                    .bg(theme::CURSOR_EDIT_BG)
                    .fg(theme::FG_ON_HIGHLIGHT)
                    .add_modifier(Modifier::BOLD)
            } else if pointing_target {
                Style::default()
                    .bg(theme::POINTING_BG)
                    .fg(theme::FG_ON_HIGHLIGHT)
                    .add_modifier(Modifier::BOLD)
            } else if let Some(c) = ref_color {
                Style::default()
                    .bg(c)
                    .fg(theme::FG_ON_HIGHLIGHT)
                    .add_modifier(Modifier::BOLD)
            } else if is_cursor {
                Style::default()
                    .bg(theme::CURSOR_NAV_BG)
                    .fg(theme::FG_ON_HIGHLIGHT)
                    .add_modifier(Modifier::BOLD)
            } else if is_selected {
                Style::default().bg(theme::SELECTION_BG).fg(theme::FG_DEFAULT)
            } else if mark_perimeter {
                Style::default()
                    .fg(theme::MARK_FG)
                    .add_modifier(Modifier::BOLD | Modifier::UNDERLINED)
            } else if search_match {
                base.bg(theme::SEARCH_MATCH_BG).fg(theme::FG_ON_HIGHLIGHT)
            } else {
                base
            };

            // Layer the cell's per-format style modifiers + colors on
            // top of the state-driven style. Applies even under cursor
            // / selection highlights — Sheets/Excel parity (a bold
            // cell stays bold when selected). Most state branches
            // reset modifiers to BOLD only, so adding format modifiers
            // here is additive. Cursor / selection / search bgs win
            // over format.bg via the layering order: state branch
            // sets bg first, then we conditionally override below.
            // For the un-highlighted ("else { base }") branch, format
            // colors fully apply.
            let highlighted = is_cursor || pointing_target || ref_color.is_some()
                || is_selected || mark_perimeter || search_match;
            if let Some(ref fmt) = cell_fmt {
                if fmt.bold {
                    style = style.add_modifier(Modifier::BOLD);
                }
                if fmt.italic {
                    style = style.add_modifier(Modifier::ITALIC);
                }
                if fmt.underline {
                    style = style.add_modifier(Modifier::UNDERLINED);
                }
                if fmt.strike {
                    style = style.add_modifier(Modifier::CROSSED_OUT);
                }
                // Foreground always applies — keeps a red-text cell
                // readable under selection bg. Background only when
                // there's no state-driven highlight that needs to win.
                if let Some(c) = fmt.fg {
                    style = style.fg(ratatui::style::Color::Rgb(c.r, c.g, c.b));
                }
                if !highlighted {
                    if let Some(c) = fmt.bg {
                        style = style.bg(ratatui::style::Color::Rgb(c.r, c.g, c.b));
                    }
                }
            }

            // Hyperlink auto-styling: cells that resolve to a URL get
            // underline + cyan fg, layered after explicit cell-format
            // fg so a user-set red URL stays red. Underline always
            // applies (reads against any bg); fg only when no format
            // fg and no state-driven highlight is in play, otherwise
            // it'd fight the highlight contrast. `url_for_cell`
            // returns Some both for plain-text URL cells and for
            // HYPERLINK custom values, so this single branch handles
            // both.
            if app.url_for_cell(row_idx, col_idx).is_some() {
                style = style.add_modifier(Modifier::UNDERLINED);
                if cell_fmt.as_ref().and_then(|f| f.fg).is_none() && !highlighted {
                    style = style.fg(theme::HYPERLINK_FG);
                }
            }

            // Datetime auto-styling: cells whose computed value is a
            // `lotus-datetime` type (jdate, jspan, …) get peach fg.
            // No modifier — these aren't clickable like hyperlinks, so
            // colour alone is the affordance. Same fg-precedence rule
            // as hyperlinks: format fg wins, state highlights win.
            #[cfg(feature = "datetime")]
            if app.datetime_tag_for_cell(row_idx, col_idx).is_some()
                && cell_fmt.as_ref().and_then(|f| f.fg).is_none()
                && !highlighted
            {
                style = style.fg(theme::DATETIME_FG);
            }

            // In edit mode, show the edit buffer in the cell with box borders.
            // Otherwise align based on the value's type.
            let cell_text = if is_cursor && app.mode == Mode::Edit {
                let w = col_w as usize;
                if w < 3 {
                    // Too narrow to draw │ borders; just clip the buffer.
                    let take = w.max(1);
                    app.edit_buf.chars().take(take).collect::<String>()
                } else if app.edit_buf.len() > w - 2 {
                    format!("│{}│", &app.edit_buf[app.edit_buf.len() - (w - 2)..])
                } else {
                    format!("│{:<pad$}│", app.edit_buf, pad = w - 2)
                }
            } else {
                match cell_fmt.as_ref().and_then(|f| f.align) {
                    Some(a) => align_text_override(a, &truncated, col_w as usize),
                    None => align_text(kind, &truncated, col_w as usize),
                }
            };

            cells.push(Cell::from(Span::styled(cell_text, style)));
        }

        rows.push(Row::new(cells));
    }

    // Column widths
    let mut widths = vec![Constraint::Length(ROW_HEADER_WIDTH)];
    for &w in &col_widths {
        widths.push(Constraint::Length(w));
    }

    // The tabline (rendered above the grid when sheets > 1) already shows
    // every sheet by index, so the grid title can stay terse.
    let title = format!(" {} | {} ", app.db_label, app.active_sheet_name());
    let table = Table::new(rows, &widths)
        .header(header)
        // Pin column_spacing explicitly. `visible_cols` and
        // `active_cell_rect` both account for it (`COL_SPACING`).
        .column_spacing(COL_SPACING)
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title(title),
        );

    f.render_widget(table, area);
}

fn draw_status(f: &mut Frame, app: &App, area: Rect) {
    // Command / Search both turn the status bar into a prompt with an
    // inline cursor.
    let prompt_prefix: Option<&str> = match app.mode {
        Mode::Command => Some(":"),
        Mode::Shell => Some("!"),
        Mode::Search(SearchDir::Forward) => Some("/"),
        Mode::Search(SearchDir::Backward) => Some("?"),
        _ => None,
    };
    if let Some(prefix) = prompt_prefix {
        let prompt = format!("{prefix}{}", app.edit_buf);
        f.render_widget(
            Paragraph::new(prompt).style(Style::default().fg(theme::MODE_COMMAND)),
            area,
        );
        let chars_before: usize = app.edit_buf[..app.edit_cursor].chars().count();
        let cursor_x = area.x + prefix.len() as u16 + chars_before as u16;
        f.set_cursor_position((cursor_x, area.y));
        return;
    }

    let (mode_label, mode_color) = match app.mode {
        Mode::Nav => (" -- NORMAL -- ", theme::MODE_NAV),
        // [sheet.editing.formula-ref-pointing] vim-style sub-mode label
        // surfaces "arrow keys are extending a ref, not moving the caret".
        Mode::Edit if app.pointing.is_some() => (" -- INSERT (point) -- ", theme::MODE_EDIT),
        Mode::Edit => (" -- INSERT -- ", theme::MODE_EDIT),
        Mode::Command => (" -- COMMAND -- ", theme::MODE_COMMAND),
        Mode::Shell => (" -- SHELL -- ", theme::MODE_COMMAND),
        Mode::Search(_) => (" -- SEARCH -- ", theme::MODE_SEARCH),
        Mode::Visual(VisualKind::Cell) => (" -- VISUAL -- ", theme::MODE_VISUAL),
        Mode::Visual(VisualKind::Row) => (" -- V-LINE -- ", theme::MODE_VISUAL),
        Mode::Visual(VisualKind::Column) => (" -- V-COLUMN -- ", theme::MODE_VISUAL),
        Mode::ColorPicker => (" -- PICKER -- ", theme::MODE_COMMAND),
        Mode::PatchShow => (" -- PATCH -- ", theme::MODE_VISUAL),
    };
    let cell_id = to_cell_id(app.cursor_row, app.cursor_col);

    let left = format!("{mode_label}{cell_id}");
    // V10 showcmd: pending count / operator / mark / prefix flag, shown
    // between the mode label and the right-aligned status text.
    let mut showcmd = String::new();
    if let Some(n) = app.pending_count {
        showcmd.push_str(&n.to_string());
    }
    if let Some(op) = app.pending_operator {
        showcmd.push(match op {
            Operator::Delete => 'd',
            Operator::Change => 'c',
            Operator::Yank => 'y',
        });
    }
    if app.pending_g {
        showcmd.push('g');
    }
    if app.pending_z {
        showcmd.push('z');
    }
    if let Some(mark) = app.pending_mark {
        showcmd.push(match mark {
            MarkAction::Set => 'm',
            MarkAction::JumpExact => '`',
            MarkAction::JumpRow => '\'',
        });
    }

    let right = if let Some(stats) = app.selection_stats() {
        stats
    } else if app.status.is_empty() {
        ":q quit  i/a:insert  v:visual  /:search  hjkl:move".to_string()
    } else {
        app.status.clone()
    };

    // Patch indicator: shown between the showcmd and the right-aligned
    // status when a `:patch new` is active. Distinct color so the user
    // can tell at a glance that edits are being recorded.
    let patch_indicator: Option<String> = app.store.patch_status().map(|s| {
        let basename = s
            .path
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("<patch>");
        if s.paused {
            format!(" [patch: {basename} paused] ")
        } else {
            format!(" [patch: {basename}] ")
        }
    });
    let patch_len = patch_indicator.as_ref().map(|s| s.len()).unwrap_or(0);

    let pad = 1 + showcmd.len() + 1; // showcmd + separators
    let gap = area
        .width
        .saturating_sub(
            left.len() as u16 + pad as u16 + patch_len as u16 + right.len() as u16 + 1,
        ) as usize;

    let mut spans = vec![Span::styled(
        left,
        Style::default()
            .fg(theme::FG_ON_HIGHLIGHT)
            .bg(mode_color)
            .add_modifier(Modifier::BOLD),
    )];
    if !showcmd.is_empty() {
        spans.push(Span::raw(" "));
        spans.push(Span::styled(
            showcmd.clone(),
            Style::default()
                .fg(theme::MODE_EDIT)
                .add_modifier(Modifier::BOLD),
        ));
    }
    spans.push(Span::raw(" ".repeat(gap)));
    if let Some(label) = patch_indicator {
        spans.push(Span::styled(
            label,
            Style::default()
                .fg(theme::FG_ON_HIGHLIGHT)
                .bg(theme::MODE_VISUAL)
                .add_modifier(Modifier::BOLD),
        ));
    }
    spans.push(Span::styled(right, Style::default().fg(theme::FG_MUTED)));

    f.render_widget(Paragraph::new(Line::from(spans)), area);
}

// ── Tabline ──────────────────────────────────────────────────────────────────
//
// Vim-style tabline shown above the grid when the workbook has >1 sheet.
// Each segment is " {1-indexed N} {name} " (with the name truncated at
// MAX_TAB_NAME_CHARS); segments are separated by `│`. The active tab is
// rendered in reverse video. When the segments overflow the available
// width, `tabline_window` selects a contiguous slice that includes the
// active tab, with `…` markers indicating elided tabs on either side.

const MAX_TAB_NAME_CHARS: usize = 16;
const TAB_SEP: &str = "│";
const TABLINE_ELLIPSIS: &str = "…";

/// `(start, end_exclusive, leading_ellipsis, trailing_ellipsis)`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct TablineWindow {
    start: usize,
    end: usize,
    leading: bool,
    trailing: bool,
}

/// Truncate `name` to at most `MAX_TAB_NAME_CHARS` characters; if truncated,
/// the last visible char is replaced with `…`.
fn truncate_tab_name(name: &str) -> String {
    let mut chars: Vec<char> = name.chars().collect();
    if chars.len() <= MAX_TAB_NAME_CHARS {
        return chars.into_iter().collect();
    }
    chars.truncate(MAX_TAB_NAME_CHARS - 1);
    let mut out: String = chars.into_iter().collect();
    out.push('…');
    out
}

fn segment_label(name: &str) -> String {
    format!(" {} ", truncate_tab_name(name))
}

/// Pick a contiguous window of tabs that fits in `total_width`, always
/// containing the active tab. `seg_widths[i]` is the rendered width (in
/// terminal cells) of segment `i`; segments are joined by a 1-cell `│`
/// separator, and elided runs on either side are signaled by a 1-cell `…`.
///
/// If even the active tab alone doesn't fit, returns the active tab anyway —
/// the caller is responsible for clipping. This shouldn't happen in practice
/// (16-char cap → max segment width ~20 cells, dwarfed by any reasonable
/// terminal).
fn tabline_window(seg_widths: &[usize], active: usize, total_width: usize) -> TablineWindow {
    let n = seg_widths.len();
    let sep_w = TAB_SEP.chars().count();
    let ellipsis_w = TABLINE_ELLIPSIS.chars().count();

    // Width if we render the whole list with no ellipsis.
    let all_width = seg_widths.iter().sum::<usize>() + sep_w * n.saturating_sub(1);
    if all_width <= total_width {
        return TablineWindow {
            start: 0,
            end: n,
            leading: false,
            trailing: false,
        };
    }

    let width_of = |start: usize, end: usize| -> usize {
        let segs = &seg_widths[start..end];
        let core = segs.iter().sum::<usize>() + sep_w * segs.len().saturating_sub(1);
        let leading = start > 0;
        let trailing = end < n;
        core + (leading as usize + trailing as usize) * ellipsis_w
    };

    // Start centered on the active tab; greedily expand outward, preferring
    // left first so reading order is preserved when scrolled.
    let mut start = active;
    let mut end = active + 1;
    if width_of(start, end) > total_width {
        return TablineWindow {
            start,
            end,
            leading: start > 0,
            trailing: end < n,
        };
    }
    loop {
        let mut grew = false;
        if start > 0 && width_of(start - 1, end) <= total_width {
            start -= 1;
            grew = true;
        }
        if end < n && width_of(start, end + 1) <= total_width {
            end += 1;
            grew = true;
        }
        if !grew {
            break;
        }
    }
    TablineWindow {
        start,
        end,
        leading: start > 0,
        trailing: end < n,
    }
}

fn draw_tabline(f: &mut Frame, app: &App, area: Rect) {
    let labels: Vec<String> = app
        .sheets
        .iter()
        .map(|s| segment_label(&s.name))
        .collect();
    let widths: Vec<usize> = labels.iter().map(|l| l.chars().count()).collect();
    let win = tabline_window(&widths, app.active_sheet, area.width as usize);

    let active_style = Style::default().add_modifier(Modifier::REVERSED);
    let inactive_style = Style::default().fg(theme::FG_MUTED);
    let dim = Style::default().fg(theme::FG_MUTED);

    let mut spans: Vec<Span<'_>> = Vec::with_capacity(win.end - win.start + 4);
    if win.leading {
        spans.push(Span::styled(TABLINE_ELLIPSIS, dim));
    }
    for (rel, idx) in (win.start..win.end).enumerate() {
        if rel > 0 {
            spans.push(Span::styled(TAB_SEP, dim));
        }
        let style = if idx == app.active_sheet {
            active_style
        } else {
            inactive_style
        };
        spans.push(Span::styled(labels[idx].clone(), style));
    }
    if win.trailing {
        spans.push(Span::styled(TABLINE_ELLIPSIS, dim));
    }

    f.render_widget(Paragraph::new(Line::from(spans)), area);
}

#[cfg(test)]
mod tests {
    use super::*;

    fn span_text(s: &Span<'_>) -> String {
        s.content.to_string()
    }

    /// Concatenated text of every span — must round-trip the input.
    fn joined(spans: &[Span<'_>]) -> String {
        spans.iter().map(span_text).collect()
    }

    #[test]
    fn non_formula_returns_single_default_span() {
        let spans = colorize_formula("hello world", &RefHighlights::default());
        assert_eq!(spans.len(), 1);
        assert_eq!(span_text(&spans[0]), "hello world");
        assert_eq!(spans[0].style, Style::default());
    }

    #[test]
    fn formula_round_trips_through_spans() {
        let formula = "=SUM(A1:A3, \"hi\", 42)";
        let spans = colorize_formula(formula, &RefHighlights::default());
        assert_eq!(joined(&spans), formula);
    }

    #[test]
    fn string_literal_renders_with_token_string_color() {
        let spans = colorize_formula("=\"hi\"+1", &RefHighlights::default());
        let str_span = spans
            .iter()
            .find(|s| s.content == "\"hi\"")
            .expect("string span");
        assert_eq!(str_span.style.fg, Some(theme::TOKEN_STRING));
    }

    #[test]
    fn cell_ref_and_function_get_distinct_colors() {
        let spans = colorize_formula("=SUM(A1)", &RefHighlights::default());
        let fn_span = spans
            .iter()
            .find(|s| s.content == "SUM")
            .expect("function span");
        let ref_span = spans
            .iter()
            .find(|s| s.content == "A1")
            .expect("cell ref span");
        assert_ne!(fn_span.style.fg, ref_span.style.fg);
        assert_eq!(ref_span.style.fg, Some(theme::TOKEN_CELLREF));
    }

    #[test]
    fn classify_recognises_numbers_booleans_and_errors() {
        assert_eq!(classify_display("3.14"), DisplayKind::Number);
        assert_eq!(classify_display("-42"), DisplayKind::Number);
        assert_eq!(classify_display("0"), DisplayKind::Number);
        assert_eq!(classify_display("TRUE"), DisplayKind::Boolean);
        assert_eq!(classify_display("FALSE"), DisplayKind::Boolean);
        assert_eq!(classify_display("#REF!"), DisplayKind::Error);
        assert_eq!(classify_display("#DIV/0!"), DisplayKind::Error);
        assert_eq!(classify_display("#NAME?"), DisplayKind::Error);
        assert_eq!(classify_display("#N/A"), DisplayKind::Error);
        assert_eq!(classify_display("hello"), DisplayKind::Text);
        assert_eq!(classify_display(""), DisplayKind::Text);
        // String that *starts* with `#` but isn't an error (rare; engine
        // never emits this, but be defensive).
        assert_eq!(classify_display("#chan"), DisplayKind::Text);
    }

    #[test]
    fn classify_recognises_currency_as_number() {
        // Formatted currency renders right-aligned like a number.
        assert_eq!(classify_display("$1.25"), DisplayKind::Number);
        assert_eq!(classify_display("$1,234.56"), DisplayKind::Number);
        assert_eq!(classify_display("-$3.00"), DisplayKind::Number);
        // Bare dollar is text.
        assert_eq!(classify_display("$"), DisplayKind::Text);
        assert_eq!(classify_display("$abc"), DisplayKind::Text);
    }

    #[test]
    fn align_text_picks_alignment_per_kind() {
        assert_eq!(align_text(DisplayKind::Number, "3.14", 8), "    3.14");
        assert_eq!(align_text(DisplayKind::Boolean, "TRUE", 8), "  TRUE  ");
        assert_eq!(align_text(DisplayKind::Error, "#REF!", 8), "#REF!   ");
        assert_eq!(align_text(DisplayKind::Text, "hi", 8), "hi      ");
    }

    #[test]
    fn type_style_distinct_per_kind() {
        // Each kind picks a different fg colour — eyes-off check that we
        // didn't collapse two kinds onto the same accent.
        let n = type_style_for(DisplayKind::Number).fg;
        let b = type_style_for(DisplayKind::Boolean).fg;
        let e = type_style_for(DisplayKind::Error).fg;
        let t = type_style_for(DisplayKind::Text).fg;
        assert_ne!(n, b);
        assert_ne!(n, e);
        assert_ne!(b, e);
        assert_ne!(t, e);
    }

    #[test]
    fn ref_highlights_assign_distinct_palette_colors_per_unique_ref() {
        // =A1*A4 → A1 and A4 are distinct refs, get different palette colors.
        let h = compute_ref_highlights("=A1*A4");
        let a1 = h.by_cell.get(&(0, 0)).expect("A1 highlighted");
        let a4 = h.by_cell.get(&(3, 0)).expect("A4 highlighted");
        assert_ne!(a1, a4);
        // First ref claims palette[0]; second claims palette[1].
        assert_eq!(*a1, theme::REF_PALETTE[0]);
        assert_eq!(*a4, theme::REF_PALETTE[1]);
    }

    #[test]
    fn ref_highlights_repeat_text_reuses_color() {
        // =A1+A1 → both occurrences map to the same color.
        let h = compute_ref_highlights("=A1+A1");
        // A1 lives at byte offset 1 and offset 4 in `=A1+A1`.
        let first = h.by_start.get(&1).expect("first A1");
        let second = h.by_start.get(&4).expect("second A1");
        assert_eq!(first, second);
    }

    #[test]
    fn ref_highlights_range_covers_every_member_cell() {
        // =SUM(A1:A3) — A1, A2, A3 all share the range's color.
        let h = compute_ref_highlights("=SUM(A1:A3)");
        let a1 = h.by_cell.get(&(0, 0)).expect("A1");
        let a2 = h.by_cell.get(&(1, 0)).expect("A2");
        let a3 = h.by_cell.get(&(2, 0)).expect("A3");
        assert_eq!(a1, a2);
        assert_eq!(a2, a3);
    }

    #[test]
    fn ref_highlights_empty_for_non_formula() {
        let h = compute_ref_highlights("hello");
        assert!(h.by_start.is_empty());
        assert!(h.by_cell.is_empty());
    }

    #[test]
    fn colorize_formula_overrides_ref_color_when_highlighted() {
        // With highlights, the A1 token should adopt the palette color
        // instead of the default cell-ref color.
        let h = compute_ref_highlights("=A1+1");
        let spans = colorize_formula("=A1+1", &h);
        let a1 = spans
            .iter()
            .find(|s| s.content == "A1")
            .expect("A1 span");
        assert_eq!(a1.style.fg, Some(theme::REF_PALETTE[0]));
        // Falls back to default when highlights are empty.
        let spans = colorize_formula("=A1+1", &RefHighlights::default());
        let a1 = spans
            .iter()
            .find(|s| s.content == "A1")
            .expect("A1 span");
        assert_eq!(a1.style.fg, Some(theme::TOKEN_CELLREF));
    }

    #[test]
    fn bare_equals_is_handled() {
        // Just `=` with nothing after it — formula_tokens returns empty;
        // we should still emit the `=` as a default span.
        let spans = colorize_formula("=", &RefHighlights::default());
        assert_eq!(joined(&spans), "=");
    }

    #[test]
    fn truncate_tab_name_caps_at_max() {
        assert_eq!(truncate_tab_name("short"), "short");
        // 16 chars exactly — stays as-is.
        assert_eq!(truncate_tab_name("0123456789abcdef"), "0123456789abcdef");
        // 17 chars — truncated to 15 chars + ellipsis.
        let long = "0123456789abcdefg";
        let out = truncate_tab_name(long);
        assert_eq!(out.chars().count(), MAX_TAB_NAME_CHARS);
        assert!(out.ends_with('…'));
    }

    #[test]
    fn tabline_window_fits_all_when_room_for_everything() {
        // Three 10-wide segments + 2 separators = 32. Width 80 fits all.
        let win = tabline_window(&[10, 10, 10], 1, 80);
        assert_eq!(
            win,
            TablineWindow {
                start: 0,
                end: 3,
                leading: false,
                trailing: false
            }
        );
    }

    #[test]
    fn tabline_window_centers_on_active_with_overflow() {
        // Five 10-wide segments → all-width = 5*10 + 4 = 54. Width 30
        // forces overflow. Active = middle (idx=2). Expect a window that
        // grows out of [2..3) until the budget runs out.
        let win = tabline_window(&[10, 10, 10, 10, 10], 2, 30);
        assert!(win.start <= 2 && 2 < win.end);
        // Ellipsis flags reflect what was cut.
        assert_eq!(win.leading, win.start > 0);
        assert_eq!(win.trailing, win.end < 5);
    }

    #[test]
    fn tabline_window_keeps_active_visible_at_far_right() {
        // Active is the last tab; algorithm must include it (no expansion to
        // the right is possible) and then add neighbors leftward.
        let win = tabline_window(&[10, 10, 10, 10, 10], 4, 30);
        assert!(win.end == 5);
        assert!(win.start >= 1); // some leading tabs cut
        assert!(win.leading);
        assert!(!win.trailing);
    }

    #[test]
    fn tabline_window_active_only_when_budget_too_tight() {
        // Width below even one segment + 2 ellipses → still returns just
        // the active tab (rendered clipped by the terminal).
        let win = tabline_window(&[10, 10, 10], 1, 5);
        assert_eq!(win.start, 1);
        assert_eq!(win.end, 2);
    }

    // ── cell_at hit-testing ──────────────────────────────────────────
    //
    // Exercises the inverse of the layout walk in `draw_grid` /
    // `active_cell_rect`. visible_rows / visible_cols are normally set
    // by `draw_grid` each frame; the tests render once to a TestBackend
    // so the values reflect the actual layout.

    use ratatui::{backend::TestBackend, Terminal};

    fn make_app() -> App {
        let store = crate::store::Store::open_in_memory().unwrap();
        App::new(store, "test")
    }

    fn render(app: &mut App, width: u16, height: u16) -> ratatui::layout::Rect {
        let backend = TestBackend::new(width, height);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| crate::ui::draw(f, app)).unwrap();
        ratatui::layout::Rect::new(0, 0, width, height)
    }

    #[test]
    fn cell_at_resolves_grid_interior_to_cell() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Grid starts at y=3 (after 3-line formula bar). Header row is
        // y=4; data rows start at y=5. First data column starts at
        // x = 1(border) + 5(gutter) + 1(spacing) = 7. Default col width
        // is 12 — so x=10 lands inside col 0, y=5 lands on row 0.
        let hit = cell_at(&app, area, 10, 5);
        assert_eq!(hit, Some(HitTarget::Cell { row: 0, col: 0 }));
    }

    #[test]
    fn cell_at_resolves_column_header_row() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // y=4 is the column-header row inside the grid block.
        let hit = cell_at(&app, area, 10, 4);
        assert_eq!(hit, Some(HitTarget::ColumnHeader(0)));
    }

    #[test]
    fn cell_at_resolves_row_gutter() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // x=2 is inside the row-number gutter (5 chars wide), y=6 → row 1.
        let hit = cell_at(&app, area, 2, 6);
        assert_eq!(hit, Some(HitTarget::RowHeader(1)));
    }

    #[test]
    fn cell_at_top_left_corner_is_none() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // y=4 (header row), x=2 (gutter axis) — the corner where the
        // row-gutter and column-header axes meet belongs to neither.
        let hit = cell_at(&app, area, 2, 4);
        assert_eq!(hit, None);
    }

    #[test]
    fn cell_at_grid_borders_are_none() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Top border of the grid block.
        assert_eq!(cell_at(&app, area, 10, 3), None);
        // Left border.
        assert_eq!(cell_at(&app, area, 0, 5), None);
    }

    #[test]
    fn cell_at_formula_bar_and_status_are_none() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Formula bar occupies y=0..3.
        assert_eq!(cell_at(&app, area, 10, 1), None);
        // Status bar is the last row (y=23 with 24-row terminal, single sheet).
        assert_eq!(cell_at(&app, area, 10, 23), None);
    }

    #[test]
    fn cell_at_past_last_visible_column_is_none() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Width 80 fits ~5-6 default-12 columns past gutter+spacing.
        // x=79 is the right border; well past the last column.
        assert_eq!(cell_at(&app, area, 79, 5), None);
    }

    #[test]
    fn cell_at_last_data_row_resolves_bottom_border_does_not() {
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        // Layout in 80x24: formula 0..3, grid 3..23 (height 20), status 23.
        // Inside grid: top border y=3, header y=4, data y=5..22, bottom
        // border y=22. So y=21 is the last data row; y=22 is the border.
        let last_data = cell_at(&app, area, 10, 21);
        assert!(matches!(last_data, Some(HitTarget::Cell { .. })));
        let bottom_border = cell_at(&app, area, 10, 22);
        assert_eq!(bottom_border, None);
    }

    #[test]
    fn cell_at_resolves_tab_in_multi_sheet_workbook() {
        let mut app = make_app();
        app.add_sheet("Sheet2").unwrap();
        let area = render(&mut app, 80, 24);
        // With 2 sheets, layout is formula(3) + grid(Min) + tabline(1)
        // + status(1) → tabline at y=22, status at y=23.
        // First tab " test " starts at x=0; click at x=2 lands inside it.
        let hit = cell_at(&app, area, 2, 22);
        assert_eq!(hit, Some(HitTarget::Tab(0)));
    }

    #[test]
    fn cell_at_no_tabline_in_single_sheet_workbook() {
        // The y-coordinate that would be the tabline in a multi-sheet
        // workbook (y=22 in 80x24) must NOT resolve to `Tab(_)` when
        // there's only one sheet — there's no tabline rendered, so it's
        // either a data row (slightly different layout because the grid
        // grows by 1 without the tabline) or the grid's bottom border.
        let mut app = make_app();
        let area = render(&mut app, 80, 24);
        let hit = cell_at(&app, area, 10, 22);
        assert!(!matches!(hit, Some(HitTarget::Tab(_))));
    }

    #[test]
    fn cell_at_respects_horizontal_scroll() {
        // After scrolling right by 2 columns, a click on the first
        // visible data column should resolve to col 2, not col 0.
        let mut app = make_app();
        app.scroll_col = 2;
        let area = render(&mut app, 80, 24);
        let hit = cell_at(&app, area, 10, 5);
        assert_eq!(hit, Some(HitTarget::Cell { row: 0, col: 2 }));
    }

    #[test]
    fn cell_at_respects_vertical_scroll() {
        let mut app = make_app();
        app.scroll_row = 100;
        let area = render(&mut app, 80, 24);
        // First visible data row after scroll should resolve to row 100.
        let hit = cell_at(&app, area, 10, 5);
        assert_eq!(hit, Some(HitTarget::Cell { row: 100, col: 0 }));
    }
}
