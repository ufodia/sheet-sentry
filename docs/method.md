# Method

Sheet Sentry answers one question: **what changed between these two saved workbooks, after accounting for supported row insertions?** It does not open Excel, recalculate formulas or prove that an edit is correct. Alignment and risk labels are review heuristics.

## 1. Reading a workbook

An `.xlsx` or `.xlsm` file is a ZIP archive of XML parts (SpreadsheetML). Sheet Sentry unzips it with fflate and reads:

| Part | What is used |
| --- | --- |
| `xl/workbook.xml` + relationships | sheet names, order and state (`visible`, `hidden`, `veryHidden`), defined names (built-in `_xlnm.*` names such as print areas are skipped) |
| `xl/worksheets/sheet*.xml` | every cell's formula, cached value and type; hidden rows; the used size |
| `xl/sharedStrings.xml` | text values, including rich text runs |
| `xl/styles.xml` | number-format codes (built-in and custom); common numeric and percentage formats are displayed, while date serials remain numeric |
| `xl/vbaProject.bin`, `xl/externalLinks/*` | whether macros and links to other workbooks exist |

**Shared formulas.** A shared master formula is expanded to its followers by shifting supported relative A1 references. Masters may appear later in the XML. Missing definitions are rejected; an out-of-bounds shifted reference becomes `#REF!`. Structured table labels and external workbook references are preserved, not interpreted as local A1 cells.

**Input limits.** Each workbook is limited to 50 MB of compressed input, 300 MB of selected uncompressed XML, 3 million populated cells and 1.5 million combined worksheet row positions. Only relevant XML parts are decompressed. Coordinates must fit Excel's 1,048,576 rows and 16,384 columns. Legacy `.xls` files are recognised and refused with a suggestion to save as `.xlsx`. Missing worksheet parts, duplicate cells, unresolved shared formulas, invalid numeric values and XML document types/entity declarations are rejected. The scanner does not load external entities and is not a general XML validator; unusual serializations may be unsupported.

## 2. Aligning rows

Comparing `A1` with `A1`, `A2` with `A2` and so on fails as soon as a row is inserted: everything below it would look changed. So rows are aligned first.

Each row gets a **signature**: its values and its formulas written in **R1C1 form**, which is independent of position (`=B4*2` in row 4 and `=B5*2` in row 5 are both `=R[0]C[-1]*2`). A row that moved but did not change keeps its signature.

A bounded **Myers diff** pairs identical signatures between the two versions. Rows left unpaired in the same gap on both sides are paired in order as *modified*; the rest are *inserted* or *deleted*. After 1,500 edit steps the search falls back to equal signatures at the same positions. Repeated labels and extensive edits can produce ambiguous alignment: inspect the before/after addresses.

Signatures encode typed values and column positions, avoiding delimiter collisions. Populated columns are indexed once, so a sheet with distant cells does not scan every empty cell in its bounding rectangle. Row-position arrays remain bounded by the input limits. Comparison runs synchronously and large workbooks can pause the interface.

Columns are not aligned: an inserted column shifts every cell to its right.

## 3. Remapping references

For supported A1 formulas, Sheet Sentry computes an **expected formula** using the inferred row insertion map, including case-insensitive cross-sheet references. This is an estimate from saved content, not a record of Excel's editing operations.

- `=Costs!B6` that is now `=Costs!B7` because a row was inserted above it: expected, **not reported**.
- `=SUM(B2:B5)` that should have become `=SUM(B2:B6)` but is still `=SUM(B2:B5)`: **range no longer covers all rows**.

**Deleted rows are conservative.** A reference to a deleted cell should not silently become a reference to the next surviving cell, while ranges have different endpoint rules. Sheets containing inferred deletions therefore do not receive insertion remapping. Unchanged formulas whose references reach a deleted row or later receive a high-severity reference-review finding. Legitimate deletion adjustments can produce extra findings.

Formula comparison preserves string contents, spaces in quoted sheet names and possible whitespace intersection operators. It is not a complete Excel grammar or semantic equivalence engine. Structured, external, 3D, dynamic-array and data-table formulas need manual verification.

## 4. Classifying differences

Each pair of cells is compared after alignment. The first matching rule decides the kind; for formulas every aspect is recorded as a note and the most severe one wins.

| Situation | Kind | Severity |
| --- | --- | --- |
| Cell now shows an error value, or the formula now contains `#REF!` outside a string | Error introduced | critical |
| Formula became a constant equal to its last result | Formula replaced by its own result | critical |
| Formula became a different constant | Formula replaced by a constant | critical |
| A range in the formula is strictly inside the expected range | Range no longer covers all rows | critical |
| More `IFERROR` / `IFNA` / `ISERROR` than expected | Errors now hidden | high |
| The set of numeric constants in the formula changed | Number inside a formula changed | high |
| A range moved to different cells or another sheet | Formula points somewhere else | high |
| Neighbours on both sides share an R1C1 formula and this cell does not (newly) | Formula breaks its pattern | high |
| A range contains the expected range and more | Range grew | medium |
| Any other formula rewrite | Formula edited | medium |
| A typed number changed | Input value changed | medium |
| Constant became a formula | Constant replaced by a formula | low |
| Text or number format changed; cell added | Text changed · Number format changed · Cell added | low |
| Cell emptied | Cell cleared | medium |
| Same supported formula, different cached result | Result recalculated (cause not established) | info |

When a formula is newly wrapped in `IFERROR(x, fallback)`, the wrapper is reported once and `x` is compared with the expected formula on its own — so the fallback `0` is not also reported as a changed constant.

**Rows and sheets.** Inserted rows are reported once per block (info); any error cell inside them is still reported as *error introduced*. Errors in new cells and new sheets are also critical. Deleted rows (medium), newly hidden rows (medium), hidden sheets (medium), unhidden sheets (low), added sheets (low) and removed sheets (high) are reported separately. Sheets are matched by name.

**Values.** Stored value types are compared, so numeric `1` and text `"1"` differ. Numeric equality uses a relative tolerance of `1e-9 × max(1, |before|, |after|)`. Smaller differences may not produce a finding. Formulas are compared as text and supported syntax; cached outputs are not recalculated or attributed to particular inputs.

**Workbook.** Defined names that were added (low), removed (medium) or now point elsewhere (high, compared after normalisation); a VBA project that was not there before (high); more external workbook links (high).

## 5. Verdict

| Verdict | Condition |
| --- | --- |
| **Risky edit** | at least one critical change |
| **Needs review** | at least one high change |
| **Minor changes** | at least one medium or low change |
| **No meaningful changes** | only info, or nothing |

## What the method cannot tell you

- **Whether the new numbers are right.** A changed input can be a correction. Sheet Sentry shows *what* changed; you decide whether it should have.
- **Values Excel never cached.** Without recalculation, outputs are only as current as the file's cached values.
- **Column insertions, renamed sheets, moved blocks.** These are reported as many cell changes, or as a removed plus an added sheet.
- **Charts, pivots, conditional formats, validation, comments and VBA code** are not compared.
- **Intent.** A “clean” verdict is not an audit opinion. It means none of the patterns above were found.
