# Contributing

Thanks for your interest! `office-mcp` is in the spec phase. Most useful
contributions right now are reviews of the documents under [docs/spec/](docs/spec/).

## How to review the spec

1. Read [docs/spec/00-overview.md](docs/spec/00-overview.md) first.
2. Open issues against specific files — quote the section you're commenting on.
3. For meaningful changes, open a PR against the spec file directly.

## What we want feedback on (M0)

- Anything missing in the architecture that you've been burned by in similar
  systems (especially: COM, Office.js, IRM, multi-instance reverse-IPC).
- Holes in the error model that would prevent agents from recovering well.
- Tools you'd add to the v1 Word catalog ([04-word-capabilities.md](docs/spec/04-word-capabilities.md))
  or, more importantly, tools you'd DROP.
- Deployment friction we haven't accounted for.

## What we don't want yet

- Code PRs implementing the spec. Hold off until M0 ships; we want to nail
  the contract first.
- Excel / PowerPoint / Outlook tool catalogs. Those are M7+.

## Discussion

GitHub issues for now. A more interactive channel will follow once M1 lands.
