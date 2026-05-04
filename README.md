# vlotus

A vim-like tool for spreadsheets in the terminal, powered by [`liblotus`](https://github.com/asg017/liblotus). Built with heavy usage of LLMs, so use at your own risk. Pre-alpha software, use wisely!

```bash
# Learn how to use vlotus with brief lessons
vlotus tutor

# Create an in-memory, ephermal spreadsheet for quick calculations
vlotus

# Persist spreadsheet edits to a file
vlotus my-spreadsheet.db

# Evaluate cell calculations
vlotus eval my-spreadsheet.db A1
```

Navigate and edit with vim-like keybindings, see [`USER-GUIDE.md`](./USER-GUIDE.md) for more info.

## Documentation

- [USER-GUIDE.md](./USER-GUIDE.md) — keys, commands, mouse gestures, the bundled tutor, the `datetime` feature.
- [KEYMAP.md](./KEYMAP.md) — exhaustive reference: every key + the App method + backing test.
- [DEVELOPMENT.md](./DEVELOPMENT.md) — architecture notes, dependencies, snapshot tests, release flow.

## License

Dual-licensed under [MIT](./LICENSE-MIT) or [Apache-2.0](./LICENSE-APACHE), at your option.
