//! UI snapshot tests using `ratatui::backend::TestBackend` + `insta`.
//!
//! Each test builds an `App` in a known state, draws one frame to a
//! fixed-size TestBackend, stringifies the resulting buffer (visible
//! glyphs only, no styles), and compares against a checked-in
//! snapshot. Run with `cargo insta review` to accept intentional
//! changes; new snapshots first appear as `.snap.new` files.
//!
//! Plain-text snapshots are deliberate — content captures most of what
//! we care about (mode banner has the mode label spelled out, etc.)
//! and styled-byte diffs would be too noisy.

use crate::store::Store;
use ratatui::{
    backend::TestBackend,
    buffer::Buffer,
    style::{Color, Modifier},
    Terminal,
};

use crate::app::{App, ColorPickerKind, MarkAction, Operator, SearchDir, SearchState, VisualKind};
use crate::theme;
use crate::tutor;
use crate::ui;

const W: u16 = 80;
const H: u16 = 24;

fn make_app() -> App {
    let store = Store::open_in_memory().unwrap();
    App::new(store, "test")
}

fn type_into(app: &mut App, row: u32, col: u32, value: &str) {
    app.cursor_row = row;
    app.cursor_col = col;
    app.start_edit_blank();
    for c in value.chars() {
        app.edit_insert(c);
    }
    app.confirm_edit();
}

fn render(app: &mut App, width: u16, height: u16) -> String {
    let backend = TestBackend::new(width, height);
    let mut terminal = Terminal::new(backend).unwrap();
    terminal.draw(|f| ui::draw(f, app)).unwrap();
    buffer_to_string(terminal.backend().buffer())
}

fn buffer_to_string(buf: &Buffer) -> String {
    let mut out = String::with_capacity((buf.area.width as usize + 1) * buf.area.height as usize);
    for y in 0..buf.area.height {
        for x in 0..buf.area.width {
            let cell = buf.cell((x, y)).expect("in bounds");
            out.push_str(cell.symbol());
        }
        out.push('\n');
    }
    out
}

/// Plain text + a "highlights" section listing cell ranges that carry a
/// non-default highlight (cursor / selection / search match / clipboard
/// mark / pointing target). Used for scenes whose feature is purely
/// styling — plain text alone can't tell them apart from the unstyled
/// case.
///
/// Highlight categories are detected by their (bg, fg, modifier) tuple
/// matching the styles `ui.rs` uses; runs of consecutive cells with the
/// same category on the same row collapse into `(y, x_lo..=x_hi)` for
/// readability.
fn render_with_highlights(app: &mut App, width: u16, height: u16) -> String {
    let backend = TestBackend::new(width, height);
    let mut terminal = Terminal::new(backend).unwrap();
    terminal.draw(|f| ui::draw(f, app)).unwrap();
    let buf = terminal.backend().buffer();

    let text = buffer_to_string(buf);
    let mut categorised: Vec<(&'static str, Vec<(u16, u16)>)> = vec![
        ("cursor", Vec::new()),
        ("edit cursor", Vec::new()),
        ("selection", Vec::new()),
        ("search match", Vec::new()),
        ("mark perimeter", Vec::new()),
        ("pointing target", Vec::new()),
        ("ref highlight", Vec::new()),
        ("hyperlink", Vec::new()),
        #[cfg(feature = "datetime")]
        ("datetime", Vec::new()),
    ];

    // Only classify cells inside the data grid. The formula bar (y=0..3),
    // the grid block borders + column header (y=3..=4), and the status
    // bar (y=h-1) all carry mode-color paint that would false-match the
    // pointing-target / mark / etc. categories.
    let data_y_lo = 5;
    let data_y_hi = buf.area.height.saturating_sub(2); // exclude bottom border + status bar
    for y in data_y_lo..data_y_hi {
        for x in 0..buf.area.width {
            let cell = buf.cell((x, y)).expect("in bounds");
            let cat = classify(cell.bg, cell.fg, cell.modifier);
            if let Some(name) = cat {
                if let Some(slot) = categorised.iter_mut().find(|(n, _)| *n == name) {
                    slot.1.push((x, y));
                }
            }
        }
    }

    let mut hlights = String::new();
    for (name, cells) in &categorised {
        if cells.is_empty() {
            continue;
        }
        let runs = collapse_runs(cells);
        hlights.push_str(name);
        hlights.push_str(": ");
        for (i, (y, x_lo, x_hi)) in runs.iter().enumerate() {
            if i > 0 {
                hlights.push_str(", ");
            }
            if x_lo == x_hi {
                hlights.push_str(&format!("(x={x_lo}, y={y})"));
            } else {
                hlights.push_str(&format!("(x={x_lo}..={x_hi}, y={y})"));
            }
        }
        hlights.push('\n');
    }

    if hlights.is_empty() {
        text
    } else {
        format!("{text}─── highlights ───\n{hlights}")
    }
}

/// Return the highlight category name for a cell's style, or `None` if
/// the cell is unstyled / part of the chrome (mode banner, headers). All
/// concrete colors come from `theme.rs`; `classify` doesn't compare hex
/// values so the assertions don't have to be re-baselined when the theme
/// rebinds a role to a different palette entry.
fn classify(bg: Color, fg: Color, modifier: Modifier) -> Option<&'static str> {
    let bold = modifier.contains(Modifier::BOLD);
    if bg == theme::CURSOR_NAV_BG && fg == theme::FG_ON_HIGHLIGHT && bold {
        return Some("cursor");
    }
    if bg == theme::CURSOR_EDIT_BG && fg == theme::FG_ON_HIGHLIGHT && bold {
        return Some("edit cursor");
    }
    if bg == theme::POINTING_BG && fg == theme::FG_ON_HIGHLIGHT && bold {
        return Some("pointing target");
    }
    if bg == theme::SELECTION_BG {
        return Some("selection");
    }
    if bg == theme::SEARCH_MATCH_BG && fg == theme::FG_ON_HIGHLIGHT {
        return Some("search match");
    }
    // Clipboard mark perimeter: foreground accent + bold + underlined, no
    // bg change — the underlying value still has to read.
    if fg == theme::MARK_FG && modifier.contains(Modifier::UNDERLINED) && bold {
        return Some("mark perimeter");
    }
    // [sheet.editing.formula-ref-highlight]: any palette bg + on-highlight fg + bold.
    if theme::REF_PALETTE.contains(&bg) && fg == theme::FG_ON_HIGHLIGHT && bold {
        return Some("ref highlight");
    }
    // Hyperlink auto-styling: cyan fg + underline, no special bg. Skip
    // cells that are part of a state highlight (those return earlier).
    if fg == theme::HYPERLINK_FG && modifier.contains(Modifier::UNDERLINED) {
        return Some("hyperlink");
    }
    // Datetime auto-styling: peach fg, no modifier (dates aren't
    // clickable). The 'no modifier' guard keeps it from collecting
    // the bold-pointing-target cells that share a similar palette.
    #[cfg(feature = "datetime")]
    if fg == theme::DATETIME_FG && modifier.is_empty() {
        return Some("datetime");
    }
    None
}

/// Collapse a flat list of (x, y) into per-row runs `(y, x_lo, x_hi)`.
/// Assumes input is naturally sorted by y then x (the loop order above).
fn collapse_runs(cells: &[(u16, u16)]) -> Vec<(u16, u16, u16)> {
    let mut out: Vec<(u16, u16, u16)> = Vec::new();
    for &(x, y) in cells {
        match out.last_mut() {
            Some(run) if run.0 == y && run.2 + 1 == x => {
                run.2 = x;
            }
            _ => out.push((y, x, x)),
        }
    }
    out
}

#[test]
fn snapshot_normal_mode_with_filled_cells() {
    let mut app = make_app();
    type_into(&mut app, 0, 0, "alpha");
    type_into(&mut app, 0, 1, "beta");
    type_into(&mut app, 1, 0, "1");
    type_into(&mut app, 1, 1, "2");
    type_into(&mut app, 1, 2, "=A2+B2");
    app.cursor_row = 0;
    app.cursor_col = 0;
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_insert_mode_mid_edit() {
    let mut app = make_app();
    app.cursor_row = 1;
    app.cursor_col = 1;
    app.start_edit_blank();
    for c in "hello".chars() {
        app.edit_insert(c);
    }
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_visual_cell_selection() {
    let mut app = make_app();
    type_into(&mut app, 0, 0, "a");
    type_into(&mut app, 0, 1, "b");
    type_into(&mut app, 1, 0, "c");
    type_into(&mut app, 1, 1, "d");
    type_into(&mut app, 2, 2, "e");
    app.cursor_row = 0;
    app.cursor_col = 0;
    app.enter_visual(VisualKind::Cell);
    app.move_cursor(2, 2);
    insta::assert_snapshot!(render_with_highlights(&mut app, W, H));
}

#[test]
fn snapshot_v_line_selection() {
    let mut app = make_app();
    type_into(&mut app, 0, 0, "row0");
    type_into(&mut app, 1, 0, "row1");
    type_into(&mut app, 2, 0, "row2");
    type_into(&mut app, 3, 0, "row3");
    app.cursor_row = 1;
    app.cursor_col = 0;
    app.enter_visual(VisualKind::Row);
    app.move_cursor(1, 0);
    insta::assert_snapshot!(render_with_highlights(&mut app, W, H));
}

#[test]
fn snapshot_command_prompt() {
    let mut app = make_app();
    app.start_command();
    for c in "tabnew Foo".chars() {
        app.edit_insert(c);
    }
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_search_prompt() {
    let mut app = make_app();
    type_into(&mut app, 0, 0, "alpha");
    type_into(&mut app, 1, 1, "alphabet");
    app.start_search(SearchDir::Forward);
    for c in "alpha".chars() {
        app.edit_insert(c);
    }
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_formula_ref_highlight() {
    // [sheet.editing.formula-ref-highlight] — typing =A1*A4 in C1 should
    // highlight A1 and A4 in distinct palette colors, both visible in the
    // grid as bg-tinted cells.
    let mut app = make_app();
    type_into(&mut app, 0, 0, "200");
    type_into(&mut app, 1, 0, "TRUE");
    type_into(&mut app, 2, 0, "FALSE");
    type_into(&mut app, 3, 0, "1.1");
    app.cursor_row = 0;
    app.cursor_col = 2; // C1
    app.start_edit_blank();
    for c in "=A1*A4".chars() {
        app.edit_insert(c);
    }
    insta::assert_snapshot!(render_with_highlights(&mut app, W, H));
}

#[test]
fn snapshot_search_active_with_matches() {
    let mut app = make_app();
    type_into(&mut app, 0, 0, "alpha");
    type_into(&mut app, 1, 1, "alphabet");
    type_into(&mut app, 3, 2, "beta");
    app.search = Some(SearchState {
        pattern: "alpha".into(),
        direction: SearchDir::Forward,
        case_insensitive: true,
    });
    // Move cursor away from the matches so cursor highlight doesn't
    // overdraw a match cell.
    app.cursor_row = 5;
    app.cursor_col = 5;
    insta::assert_snapshot!(render_with_highlights(&mut app, W, H));
}

#[test]
fn snapshot_showcmd_pending_5d() {
    let mut app = make_app();
    type_into(&mut app, 0, 0, "x");
    app.cursor_row = 0;
    app.cursor_col = 0;
    app.pending_count = Some(5);
    app.pending_operator = Some(Operator::Delete);
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_showcmd_pending_g_with_count() {
    let mut app = make_app();
    app.pending_count = Some(12);
    app.pending_g = true;
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_showcmd_pending_mark() {
    let mut app = make_app();
    app.pending_mark = Some(MarkAction::JumpExact);
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_hyperlink_autostyle() {
    // Three URL cells (https / mailto / file) styled as hyperlinks,
    // one plain text cell that must NOT pick up the hyperlink fg, and
    // a URL cell under the cursor to confirm cursor highlight wins
    // over hyperlink fg (underline still layered on top).
    let mut app = make_app();
    type_into(&mut app, 0, 0, "https://example.com");
    type_into(&mut app, 1, 0, "plain text");
    type_into(&mut app, 2, 0, "mailto:user@host");
    type_into(&mut app, 3, 0, "file:///tmp/x");
    // Wider column A so the URLs aren't clipped.
    app.set_column_width(0, 22).unwrap();
    // Park cursor on a non-URL cell so the link cells render with their
    // own fg (default cursor on A1 would otherwise repaint the first
    // run in cursor bg/fg).
    app.cursor_row = 5;
    app.cursor_col = 0;
    insta::assert_snapshot!(render_with_highlights(&mut app, W, H));
}

#[cfg(feature = "datetime")]
#[test]
fn snapshot_datetime_autostyle() {
    // Four datetime cells (date / datetime / span / time) styled
    // peach, plus a plain text cell that must NOT pick up the
    // datetime fg. Cursor parked off the column so the date cells
    // render with their own fg.
    let mut app = make_app();
    type_into(&mut app, 0, 0, "2025-04-27"); // jdate
    type_into(&mut app, 1, 0, "plain text");
    type_into(&mut app, 2, 0, "2025-04-27T12:30:45"); // jdatetime
    type_into(&mut app, 3, 0, "=DATE(2025, 2, 1) - DATE(2025, 1, 1)"); // jspan
    type_into(&mut app, 4, 0, "12:30:45"); // jtime
    app.set_column_width(0, 22).unwrap();
    app.cursor_row = 6;
    app.cursor_col = 0;
    insta::assert_snapshot!(render_with_highlights(&mut app, W, H));
}

#[cfg(feature = "datetime")]
#[test]
fn snapshot_datetime_friendly_display() {
    // Pinned-time arithmetic so the snapshot doesn't drift: 2026-05-01
    // minus 2025-01-01 is 486 days. The friendly form (whatever jiff
    // emits for that span) must render in the cell — not the canonical
    // P*D form.
    let mut app = make_app();
    type_into(&mut app, 0, 0, "=DATE(2026, 5, 1) - DATE(2025, 1, 1)");
    app.set_column_width(0, 22).unwrap();
    app.cursor_row = 5;
    app.cursor_col = 0;
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_hyperlink_function_renders_label() {
    // HYPERLINK custom-type cells display their label in the grid
    // (not the encoded url\u{1F}label payload), and pick up the same
    // underline + cyan styling as auto-detected URL strings.
    let mut app = make_app();
    type_into(
        &mut app,
        0,
        0,
        "=HYPERLINK(\"https://example.com\", \"click here\")",
    );
    type_into(&mut app, 1, 0, "https://plain.example.com");
    type_into(&mut app, 2, 0, "regular cell");
    app.set_column_width(0, 22).unwrap();
    app.cursor_row = 5;
    app.cursor_col = 0;
    insta::assert_snapshot!(render_with_highlights(&mut app, W, H));
}

#[test]
fn snapshot_clipboard_mark_perimeter() {
    let mut app = make_app();
    type_into(&mut app, 1, 1, "a");
    type_into(&mut app, 1, 2, "b");
    type_into(&mut app, 2, 1, "c");
    type_into(&mut app, 2, 2, "d");
    // Yank a 2x2 selection so the mark perimeter shows.
    app.cursor_row = 1;
    app.cursor_col = 1;
    app.enter_visual(VisualKind::Cell);
    app.move_cursor(1, 1);
    app.set_clipboard_mark(crate::app::ClipMarkMode::Copy);
    app.exit_visual();
    // Move cursor away so the cursor highlight doesn't paint over the mark.
    app.cursor_row = 5;
    app.cursor_col = 5;
    insta::assert_snapshot!(render_with_highlights(&mut app, W, H));
}

#[test]
fn snapshot_tutor_l1_at_startup() {
    let mut store = Store::open_in_memory().unwrap();
    tutor::seed_tutor_db(&mut store).unwrap();
    let mut app = App::new(store, "tutor");
    // Wider terminal so the formula bar shows the full first
    // instruction; the 80-col default truncates lesson copy.
    insta::assert_snapshot!(render(&mut app, 160, 30));
}

#[test]
fn snapshot_pointing_active_shows_point_submode_in_banner() {
    // [sheet.editing.formula-ref-pointing] while pointing is active the
    // mode banner reads " -- INSERT (point) -- " instead of the plain
    // " -- INSERT -- " — surfaces "arrow keys are extending a ref now".
    let mut app = make_app();
    app.cursor_row = 0;
    app.cursor_col = 0;
    app.start_edit();
    app.edit_insert('=');
    // Right-arrow → enter pointing on B1.
    assert!(app.try_pointing_arrow(0, 1, false));
    assert!(app.pointing.is_some());
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_variable_column_widths() {
    let mut app = make_app();
    // Column A wider for long text, B narrower for numbers, C autofit.
    app.set_column_width(0, 22).unwrap();
    app.set_column_width(1, 6).unwrap();
    type_into(&mut app, 0, 0, "Long descriptive label");
    type_into(&mut app, 1, 0, "Another row");
    type_into(&mut app, 0, 1, "1");
    type_into(&mut app, 1, 1, "42");
    type_into(&mut app, 0, 2, "x");
    type_into(&mut app, 1, 2, "yz");
    app.autofit_column(2).unwrap();
    app.cursor_row = 0;
    app.cursor_col = 0;
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_multi_tab_workbook_second_active() {
    let mut app = make_app();
    app.add_sheet("Two").unwrap();
    app.add_sheet("Three").unwrap();
    // add_sheet switches to the new sheet; bounce to the second one.
    app.switch_sheet(1);
    type_into(&mut app, 0, 0, "on Two");
    app.cursor_row = 0;
    app.cursor_col = 0;
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn right_aligned_value_renders_full_at_every_column_width() {
    // Regression: with column_spacing=1 the visible_cols loop didn't include
    // the spacing in its budget, so the layout overflowed by `visible_cols`
    // cells and ratatui's solver shrank the widest column. A right-aligned
    // numeric value like "100" then dropped its trailing digit (showed as
    // "10"). Reproduces with col A wide and many trailing default-width
    // columns competing for room. See `COL_SPACING` in `ui.rs`.
    let mut app = make_app();
    type_into(&mut app, 0, 0, "100");
    app.cursor_row = 0;
    app.cursor_col = 0;
    for w in [10u16, 30, 50, 75, 76, 80] {
        app.set_column_width(0, w).unwrap();
        let out = render(&mut app, 200, 30);
        // Pull out non-space, non-border chars from the cursor row.
        // Row index: formula bar (3) + grid border (1) + header (1) = 5.
        let row = out.lines().nth(5).unwrap();
        let digits: String = row
            .chars()
            .filter(|&c| c.is_ascii_digit())
            .collect();
        // The cursor row's only digits should be the row number (1) plus
        // every digit of "100".
        assert_eq!(
            digits, "1100",
            "col_w={w} row chars: {row:?}"
        );
    }
}

#[test]
fn snapshot_full_format_grid() {
    // Comprehensive single-frame check for the format epic: every axis
    // appears at least once across the cells. Plain-text snapshot
    // captures alignment + formatted display; styling axes (bold /
    // italic / underline / strike / fg / bg) don't change the text
    // bytes but participate in the same render path that produces
    // this string.
    let mut app = make_app();
    // Number formats:
    type_into(&mut app, 0, 0, "$1.25");      // USD/2 via auto-detect
    type_into(&mut app, 1, 0, "4.5%");        // Percent/1 via auto-detect
    type_into(&mut app, 2, 0, "100");          // plain number
    type_into(&mut app, 3, 0, "hello");        // plain text
    // Apply explicit alignment to row 4 to demonstrate override:
    type_into(&mut app, 4, 0, "centered");
    app.cursor_row = 4;
    app.cursor_col = 0;
    app.apply_format_update(|f| f.align = Some(crate::format::Align::Center));
    app.cursor_row = 0;
    app.cursor_col = 0;
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_color_picker_open() {
    // gf-opened picker overlaying the grid. Snapshot captures the
    // 4×6 swatch grid (selection on swatch 0 = "red"), the status
    // line with name + hex, and the hint footer.
    let mut app = make_app();
    type_into(&mut app, 0, 0, "x");
    app.cursor_row = 0;
    app.cursor_col = 0;
    app.open_color_picker(ColorPickerKind::Fg);
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_color_picker_hex_mode() {
    // Same picker after toggling into hex-input mode (`?`) and typing.
    let mut app = make_app();
    type_into(&mut app, 0, 0, "x");
    app.cursor_row = 0;
    app.cursor_col = 0;
    app.open_color_picker(ColorPickerKind::Bg);
    app.color_picker_toggle_hex();
    for c in "fa3".chars() {
        app.color_picker_hex_input(c);
    }
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_currency_formatted_grid() {
    // Typing `$1.25` style input goes through auto-detect: stored raw
    // is `1.25`, format is USD/2, the grid shows `$1.25` right-aligned.
    // Mixes positive, negative, thousands-grouped, and a non-currency
    // text fallthrough (`$abc` stays as a left-aligned string) so the
    // snapshot covers all alignment cases in one frame.
    let mut app = make_app();
    type_into(&mut app, 0, 0, "$1.25");
    type_into(&mut app, 1, 0, "$1,234.56");
    type_into(&mut app, 2, 0, "-$3");
    type_into(&mut app, 3, 0, "$abc");
    type_into(&mut app, 4, 0, "42"); // unformatted number for comparison
    app.cursor_row = 0;
    app.cursor_col = 0;
    insta::assert_snapshot!(render(&mut app, W, H));
}

#[test]
fn snapshot_tabline_overflow_with_active_in_view() {
    // Many sheets at a narrow width force the tabline to elide tabs on
    // either side of the active one. Active is in the middle so we
    // expect both leading and trailing `…` markers.
    let mut app = make_app();
    for name in &["Two", "Three", "Four", "Five", "Six", "Seven", "Eight"] {
        app.add_sheet(name).unwrap();
    }
    // Active = "Four" (idx 3 of 8).
    app.switch_sheet(3);
    insta::assert_snapshot!(render(&mut app, 50, 12));
}
