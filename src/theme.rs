//! Color palette and semantic role mappings for the vlotus UI.
//!
//! All ratatui [`Color`] choices in `ui.rs` and `snapshots.rs` route through
//! this module. The two layers are separate on purpose:
//!
//! - [`mocha`] — the raw Catppuccin Mocha palette (rosewater through crust).
//!   Don't reference these directly from render code; pick a semantic role
//!   from the second half of this file instead.
//! - Semantic constants — `MODE_EDIT`, `TOKEN_CELLREF`, `CURSOR_NAV_BG`,
//!   `REF_PALETTE`, etc. These describe *intent* (what's being colored),
//!   so when a future PR swaps Catppuccin for another palette or wires
//!   up runtime theming, the call sites don't change.
//!
//! Hex values come from <https://catppuccin.com/palette> (Mocha flavour).

use ratatui::style::Color;

/// Catppuccin Mocha palette — raw colors. Prefer the semantic constants
/// below for everything except the [`REF_PALETTE`] (which is intentionally
/// drawn from the bright accents). The full palette is exposed so a future
/// theme/role binding can pick up an entry without re-deriving the hex.
#[allow(dead_code)]
pub mod mocha {
    use ratatui::style::Color;

    pub const ROSEWATER: Color = Color::Rgb(0xf5, 0xe0, 0xdc);
    pub const FLAMINGO: Color = Color::Rgb(0xf2, 0xcd, 0xcd);
    pub const PINK: Color = Color::Rgb(0xf5, 0xc2, 0xe7);
    pub const MAUVE: Color = Color::Rgb(0xcb, 0xa6, 0xf7);
    pub const RED: Color = Color::Rgb(0xf3, 0x8b, 0xa8);
    pub const MAROON: Color = Color::Rgb(0xeb, 0xa0, 0xac);
    pub const PEACH: Color = Color::Rgb(0xfa, 0xb3, 0x87);
    pub const YELLOW: Color = Color::Rgb(0xf9, 0xe2, 0xaf);
    pub const GREEN: Color = Color::Rgb(0xa6, 0xe3, 0xa1);
    pub const TEAL: Color = Color::Rgb(0x94, 0xe2, 0xd5);
    pub const SKY: Color = Color::Rgb(0x89, 0xdc, 0xeb);
    pub const SAPPHIRE: Color = Color::Rgb(0x74, 0xc7, 0xec);
    pub const BLUE: Color = Color::Rgb(0x89, 0xb4, 0xfa);
    pub const LAVENDER: Color = Color::Rgb(0xb4, 0xbe, 0xfe);

    pub const TEXT: Color = Color::Rgb(0xcd, 0xd6, 0xf4);
    pub const SUBTEXT1: Color = Color::Rgb(0xba, 0xc2, 0xde);
    pub const SUBTEXT0: Color = Color::Rgb(0xa6, 0xad, 0xc8);
    pub const OVERLAY2: Color = Color::Rgb(0x93, 0x99, 0xb2);
    pub const OVERLAY1: Color = Color::Rgb(0x7f, 0x84, 0x9c);
    pub const OVERLAY0: Color = Color::Rgb(0x6c, 0x70, 0x86);
    pub const SURFACE2: Color = Color::Rgb(0x58, 0x5b, 0x70);
    pub const SURFACE1: Color = Color::Rgb(0x45, 0x47, 0x5a);
    pub const SURFACE0: Color = Color::Rgb(0x31, 0x32, 0x44);
    pub const BASE: Color = Color::Rgb(0x1e, 0x1e, 0x2e);
    pub const MANTLE: Color = Color::Rgb(0x18, 0x18, 0x25);
    pub const CRUST: Color = Color::Rgb(0x11, 0x11, 0x1b);
}

// ── Foreground / chrome ──────────────────────────────────────────────────

/// Default foreground for ordinary cell text.
pub const FG_DEFAULT: Color = mocha::TEXT;
/// Muted foreground for inactive headers, separators, hint text.
pub const FG_MUTED: Color = mocha::OVERLAY1;
/// Foreground used on top of any saturated highlight bg (cursor, ref tint,
/// search match) — always the darkest base color so contrast holds.
pub const FG_ON_HIGHLIGHT: Color = mocha::BASE;

/// Default block border (formula bar in Nav mode, grid block).
pub const BORDER_DEFAULT: Color = mocha::OVERLAY0;
/// Accent border for the autocomplete popup.
pub const BORDER_ACCENT: Color = mocha::YELLOW;
/// Accent border for the signature-help tooltip.
pub const BORDER_INFO: Color = mocha::SAPPHIRE;
/// Background fill for floating popups (autocomplete list).
pub const POPUP_BG: Color = mocha::MANTLE;

// ── Modes (used by the formula-bar block title + the status banner) ─────

pub const MODE_NAV: Color = mocha::SAPPHIRE;
pub const MODE_EDIT: Color = mocha::YELLOW;
pub const MODE_VISUAL: Color = mocha::MAUVE;
pub const MODE_COMMAND: Color = mocha::PEACH;
pub const MODE_SEARCH: Color = mocha::PEACH;

// ── Cell display kinds (computed-value rendering) ───────────────────────

pub const CELL_NUMBER: Color = mocha::PEACH;
pub const CELL_BOOLEAN: Color = mocha::SKY;
pub const CELL_ERROR: Color = mocha::RED;
pub const CELL_TEXT: Color = mocha::TEXT;

// ── Formula tokens (formula bar syntax-highlight) ───────────────────────

pub const TOKEN_STRING: Color = mocha::GREEN;
pub const TOKEN_NAME: Color = mocha::MAUVE;
pub const TOKEN_CELLREF: Color = mocha::SKY;
pub const TOKEN_FUNCTION: Color = mocha::YELLOW;
pub const TOKEN_NUMBER: Color = mocha::PEACH;
pub const TOKEN_BOOLEAN: Color = mocha::SKY;

// ── Cell state highlights ───────────────────────────────────────────────

/// Cursor cell while the user is typing into it.
pub const CURSOR_EDIT_BG: Color = mocha::YELLOW;
/// Cursor cell in Nav (and any non-edit) mode.
pub const CURSOR_NAV_BG: Color = mocha::SUBTEXT0;
/// Cell currently being inserted as a ref via arrow-key pointing.
pub const POINTING_BG: Color = mocha::SAPPHIRE;
/// Visual-mode rectangular selection.
pub const SELECTION_BG: Color = mocha::SURFACE1;
/// Active search match tint. Distinct from the edit cursor so a cell that
/// is both edited and matches the active pattern reads unambiguously.
pub const SEARCH_MATCH_BG: Color = mocha::FLAMINGO;
/// Clipboard-mark perimeter accent (fg only — the cell's bg stays put so
/// the underlying value remains legible).
pub const MARK_FG: Color = mocha::YELLOW;
/// Anchor cell of a spill region.
pub const SPILL_ANCHOR_FG: Color = mocha::TEAL;
/// Foreground for cells whose text is auto-detected as a URL (or, in
/// later tickets, an explicit hyperlink custom value). Picked so the
/// "links are blue" affordance reads against the dark base bg.
pub const HYPERLINK_FG: Color = mocha::SAPPHIRE;
/// Foreground for cells whose computed value is a `lotus-datetime`
/// type (`jdate`, `jspan`, etc.). Peach is the analog of the
/// hyperlink sapphire — distinct enough to read at a glance while
/// staying within the Mocha palette.
#[cfg(feature = "datetime")]
pub const DATETIME_FG: Color = mocha::PEACH;
/// Interactive picker selection (autocomplete list, etc.).
pub const PICKER_SELECTED_BG: Color = mocha::SURFACE2;

// ── Formula ref highlights ──────────────────────────────────────────────

/// Colors handed out to unique cell refs in the formula being typed.
/// Six bright Catppuccin accents that stay distinct on `BASE` bg and
/// remain legible with `FG_ON_HIGHLIGHT` on top — see
/// [sheet.editing.formula-ref-highlight].
pub const REF_PALETTE: [Color; 6] = [
    mocha::BLUE,
    mocha::GREEN,
    mocha::PEACH,
    mocha::MAUVE,
    mocha::TEAL,
    mocha::RED,
];
