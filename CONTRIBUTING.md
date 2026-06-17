# Contributing

Thanks for your interest! `office-mcp` is in an implementation phase with the
design contract kept under [doc/spec/](doc/spec/).

## How to review the spec

1. Read [doc/spec/00-overview.md](doc/spec/00-overview.md) first.
2. Open issues against specific files — quote the section you're commenting on.
3. For meaningful changes, open a PR against the spec file directly.

## Repository layout

- `src/office-mcp/daemon/` is the native Rust daemon package. Run Rust build,
  test, and lint commands from the repository root.
- `src/office-mcp/daemon/evidence/` contains runtime, UI, smoke, and validation
  evidence harnesses. Run evidence commands from this directory.

- `src/office-mcp/ui/` contains the daemon web console assets.
- `src/office-ctl/common/` contains shared TypeScript add-in utilities.
- `src/office-ctl/word/` is the Word add-in package. Run manifest validation and task pane
  checks from this directory.
- `src/office-ctl/excel/` is reserved for the Excel add-in entry point and
  host-specific command implementations.
- `packaging/` contains cross-component installers and release packaging.
- Generated artifacts and historical feasibility evidence belong under
  `artifacts/`, not as additional root-level source packages.

## What we want feedback on (M0)

- Anything missing in the architecture that you've been burned by in similar
  systems (especially: COM, Office.js, IRM, multi-instance reverse-IPC).
- Holes in the error model that would prevent agents from recovering well.
- Tools you'd add to the v1 Word catalog ([04-word-capabilities.md](doc/spec/04-word-capabilities.md))
  or, more importantly, tools you'd DROP.
- Deployment friction we haven't accounted for.

## What we don't want yet

- New root-level application packages. Keep daemon runtime code in
  `src/office-mcp/daemon/`, daemon UI code in `src/office-mcp/ui/`, add-in
  runtime code in `src/office-ctl/`, and installer glue in `packaging/`.
- Unplanned host tool catalogs. Word and Excel v1 have specs; PowerPoint and
  Outlook remain future milestones until their capability docs are written.

## Discussion

GitHub issues for now. A more interactive channel will follow once M1 lands.
