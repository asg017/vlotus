# vlotus

A standalone vim-style terminal spreadsheet on top of
[`lotus-core`](https://github.com/asg017/liblotus/tree/main/crates/lotus-core),
built with [`ratatui`](https://ratatui.rs) and `crossterm`. SQLite
storage is owned by `src/store/`. Reference implementation of an
interactive UI backed by the engine.

## Run

```bash
# interactive TUI, creates the db file if it doesn't exist
cargo run -p vlotus [db_path]

# one-shot cell evaluation (no UI)
cargo run -p vlotus -- eval <db_path> <cell>

# bundled vimtutor-style lesson workbook (in-memory; mutations vanish on quit)
cargo run -p vlotus -- tutor

# work with portable `.lpatch` diff files
cargo run -p vlotus -- patch apply <db> <patch> [--invert] [--on-conflict omit|replace|abort]
cargo run -p vlotus -- patch show <patch>
cargo run -p vlotus -- patch invert <input> <output>
cargo run -p vlotus -- patch combine <output> <patch>...
```

Default `db_path` is `sheet.db` in the current directory.

## Documentation

- [USER-GUIDE.md](./USER-GUIDE.md) — keys, commands, mouse gestures, the bundled tutor, the `datetime` feature.
- [KEYMAP.md](./KEYMAP.md) — exhaustive reference: every key + the App method + backing test.

## Build

```bash
cargo build -p vlotus --release
./target/release/vlotus sheet.db
```

## UI snapshot tests

`src/snapshots.rs` covers ~20 styled UI states (Normal, Insert, Visual,
V-LINE, Command/Search prompts, search highlight, showcmd, clipboard mark,
tutor L1, multi-tab, formatted cells, color picker) using
`ratatui::backend::TestBackend` + `insta`.

When an intentional UI change lands, expect those tests to fail. Update them
with:

```bash
INSTA_UPDATE=new cargo test -p vlotus       # regenerate .snap.new files
cargo insta review                           # interactive accept (if cargo-insta installed)
# or just: rename .snap.new → .snap after eyeballing
```

Snapshots live alongside the test source under `src/snapshots/`.

## Dependencies

lotus-core + lotus-datetime (git deps from
[`asg017/liblotus`](https://github.com/asg017/liblotus)), ratatui,
crossterm, rusqlite, clap, arboard, thiserror,
fallible-streaming-iterator, open, csv, serde_json. Dev: insta.
