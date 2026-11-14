# Changelog

All notable changes to Sheet Sentry are documented here. The project follows [Semantic Versioning](https://semver.org/).

## 1.0.0 — 2026-11-14

First public release.

### Added

- `.xlsx`/`.xlsm` reader with shared strings, rich text, number formats, defined names, hidden rows and sheets, shared-formula expansion, macro and external-link detection, and size limits.
- Row alignment with a Myers diff over position-independent (R1C1) row signatures, and reference remapping through the row map, so references Excel moved by itself are not reported.
- 29 change kinds in five severities, including hardcoded formulas, shrunken ranges, new errors (also inside inserted rows), `IFERROR`/`IFNA` masking, changed constants, moved references, broken formula patterns, redefined names, hidden rows and sheets.
- Verdict: Risky edit, Needs review, Minor changes, No meaningful changes.
- Browser app with sheet tabs, severity-coloured cells, an inspector with the expected formula after row changes, severity filters and recalculated results on demand.
- Markdown, CSV (formula-injection safe) and JSON reports.
- `npm run review` command line with `--fail-on` exit codes for CI.
- Synthetic Q3 budget demo, downloadable as two real `.xlsx` files.
- 38 unit tests, 8 Playwright browser tests and CI.
- Strict Content-Security-Policy with `connect-src 'none'`.
