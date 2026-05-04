# vlotus development

Architectural notes, dependencies, and test loops. For user-facing
docs see [USER-GUIDE.md](./USER-GUIDE.md); for the exhaustive keymap
+ backing tests see [KEYMAP.md](./KEYMAP.md); for agent-facing
guidance see [CLAUDE.md](./CLAUDE.md).

## Architecture

vlotus is a reference implementation of an interactive UI on top of
[`lotus-core`](https://github.com/asg017/liblotus/tree/main/crates/lotus-core).

- `src/main.rs` / `src/app.rs` — `App` struct + run loop, keymap
  dispatch, mode machine.
- `src/store/` — SQLite-backed persistence (long-lived dirty-buffer
  txn; commit on `:w`, rollback on `:q!`). Schema, undo log, and
  patch sessions live here.
- `src/format.rs`, `src/datetime.rs`, `src/hyperlink.rs` — composable
  cell-format axes, datetime extension wiring, and the hyperlink
  custom type / function.
- `src/shell.rs` — the `!`-prompt subprocess paste pipeline (sniffs
  JSON / NDJSON / TSV / CSV / plain).
- `src/snapshots.rs` — UI snapshot tests via `ratatui::backend::TestBackend`
  + `insta`.
- `src/patch_cli.rs` — `vlotus patch …` subcommand surface
  (apply / show / invert / combine / diff for `.lpatch` files).

`CLAUDE.md` has the most detailed cross-module notes (when each
binding/command goes where, render-layer ordering, format-axis
mutation path, undo / redo / session interaction, etc.).

## Build

```bash
cargo build -p vlotus --release
./target/release/vlotus sheet.db
```

The `datetime` Cargo feature is on by default. Build a smaller binary
without it via `cargo build -p vlotus --no-default-features`.

## Dependencies

lotus-core + lotus-datetime (git deps from
[`asg017/liblotus`](https://github.com/asg017/liblotus)), ratatui,
crossterm, rusqlite, clap, arboard, thiserror,
fallible-streaming-iterator, open, csv, serde_json. Dev: insta.

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

## Release

CI is wired through [`cargo-dist`](https://opensource.axo.dev/cargo-dist/).
Tag a version (`vX.Y.Z`) on `main` and `.github/workflows/release.yml`
will build per-platform archives, build PyPI wheels via `uvx maturin
build -b bin` (subworkflow `.github/workflows/build-pypi.yml`),
publish them to PyPI via `uv publish` (subworkflow
`.github/workflows/publish-pypi.yml`), and create a GitHub Release.

Configuration:

- `dist-workspace.toml` — cargo-dist config (targets, installers,
  custom local + publish jobs).
- `pyproject.toml` — maturin build-system at the repo root.
- PyPI publishing uses trusted publishing (`id-token: write`,
  `environment: release`), so a `release` GitHub Environment must
  exist on the repo and the PyPI project must be configured with
  `asg017/vlotus` → `release.yml` → `release` as a trusted publisher.

`.github/workflows/test.yml` runs `cargo test` across the same target
matrix on every push to `main` and on PRs.
