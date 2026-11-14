# Changelog

All notable changes to Sheet Sentry are documented here. The project follows [Semantic Versioning](https://semver.org/).

## 1.0.0 — 2026-11-14

First public release.

### Added

- `.xlsx`/`.xlsm` reader with shared strings, rich text, number formats, defined names, hidden rows and sheets, shared-formula expansion, macro and external-link detection, and size limits.
- Bounded row alignment over typed position-independent (R1C1) signatures; inferred insertion reference remapping and conservative reference review after deletions.
- 29 change kinds in five severities, including hardcoded formulas, shrunken ranges, new errors (also inside inserted rows), `IFERROR`/`IFNA` masking, changed constants, moved references, broken formula patterns, redefined names, hidden rows and sheets.
- Verdict: Risky edit, Needs review, Minor changes, No meaningful changes.
- Browser app with sheet tabs, severity-coloured cells, an inspector with the expected formula after row changes, severity filters and recalculated results on demand.
- Markdown with escaped workbook text, CSV with formula-like fields neutralized, and JSON reports.
- `npm run review` command line with `--fail-on` exit codes for CI.
- Synthetic Q3 budget demo, downloadable as two real `.xlsx` files.
- 64 unit/CLI tests, 11 Playwright browser tests and CI.
- Strict Content-Security-Policy with `connect-src 'none'`.

### Hardened before publication

- Resolve shared-formula followers even when their master occurs later; reject missing definitions and missing worksheet parts.
- Validate worksheet bounds, XML character references, numeric values, style indices and duplicate cells; reject XML entity declarations and bound aggregate row spans.
- Index sparse columns for comparison, preserve value types and meaningful formula whitespace, and flag errors in newly added cells/sheets.
- Prevent stale file reads and failed replacements from comparing old inputs.
- Protect input workbooks from CLI report overwrites, including file aliases, and return documented option/output errors.
- Quote carriage returns in CSV and escape untrusted workbook markup and filenames in Markdown.
