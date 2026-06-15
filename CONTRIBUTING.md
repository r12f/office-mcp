# Contributing

Thanks for your interest! `office-mcp` is in an implementation phase with the
design contract kept under [docs/spec/](docs/spec/).

## How to review the spec

1. Read [docs/spec/00-overview.md](docs/spec/00-overview.md) first.
2. Open issues against specific files — quote the section you're commenting on.
3. For meaningful changes, open a PR against the spec file directly.

## Repository layout

- `mcp-server/` is the long-running daemon package. Run server build, test,
  CLI, smoke, and runtime-evidence commands from this directory.
- `addin/` is the Word add-in package. Run manifest validation and task pane
  checks from this directory.
- `packaging/` contains cross-component installers and release packaging.
- Generated artifacts and historical feasibility evidence belong under
  `artifacts/`, not as additional root-level source packages.

## What we want feedback on (M0)

- Anything missing in the architecture that you've been burned by in similar
  systems (especially: COM, Office.js, IRM, multi-instance reverse-IPC).
- Holes in the error model that would prevent agents from recovering well.
- Tools you'd add to the v1 Word catalog ([04-word-capabilities.md](docs/spec/04-word-capabilities.md))
  or, more importantly, tools you'd DROP.
- Deployment friction we haven't accounted for.

## What we don't want yet

- New root-level application packages. Keep long-lived runtime code in
  `mcp-server/`, add-in runtime code in `addin/`, and installer glue in
  `packaging/`.
- Excel / PowerPoint / Outlook tool catalogs. Those are M7+.

## Discussion

GitHub issues for now. A more interactive channel will follow once M1 lands.
