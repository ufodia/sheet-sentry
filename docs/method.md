# Method

Sheet Sentry answers one question: **what did the edit change in this workbook, beyond what Excel would have changed by itself?** It works on the saved files only — it does not open Excel and does not recalculate.

## 1. Reading a workbook

An `.xlsx` or `.xlsm` file is a ZIP archive of XML parts (SpreadsheetML). Sheet Sentry unzips it with fflate and reads:

| Part | What is used |
| --- | --- |
| `xl/workbook.xml` + relationships | sheet names, order and state (`visible`, `hidden`, `veryHidden`), defined names (built-in `_xlnm.*` names such as print areas are skipped) |
| `xl/worksheets/sheet*.xml` | every cell's formula, cached value and type; hidden rows; the used size |
| `xl/sharedStrings.xml` | text values, including rich text runs |
| `xl/styles.xml` | number formats (built-in and custom), so percentages and dates display sensibly |
| `xl/vbaProject.bin`, `xl/externalLinks/*` | whether macros and links to other workbooks exist |

**Shared formulas.** Excel stores a formula filled down a column once and marks the other cells as sharing it. Sheet Sentry expands each shared formula to every cell, shifting relative references exactly as Excel does, so every cell has its own formula.

**Safety limits.** Files over 50 MB, archives that expand to more than 300 MB of XML, and sheets with more than 3 million cells are refused with a clear message. Legacy `.xls` files are recognised by their signature and refused with a suggestion to save as `.xlsx`.

## 2. Aligning rows

Comparing `A1` with `A1`, `A2` with `A2` and so on fails as soon as a row is inserted: everything below it would look changed. So rows are aligned first.

Each row gets a **signature**: its values and its formulas written in **R1C1 form**, which is independent of position (`=B4*2` in row 4 and `=B5*2` in row 5 are both `=R[0]C[-1]*2`). A row that moved but did not change keeps its signature.

A **Myers diff** (the longest-common-subsequence algorithm behind `git diff`) pairs identical signatures between the two versions. Rows left unpaired in the same gap on both sides — typically rows whose values or formulas were edited — are paired up in order as *modified* and compared cell by cell; the rest are *inserted* or *deleted*. The result is a **row map** from every original row to its row in the edited version. A deleted row maps to the next surviving row, as Excel's reference adjustment does.

Columns are not aligned: an inserted column shifts every cell to its right.

## 3. Remapping references

For every formula in the original, Sheet Sentry computes the **expected formula**: the original with each row reference passed through the row map of the sheet it points at (including references to other sheets). This is what Excel itself would have produced while the rows were inserted or deleted.

- `=Costs!B6` that is now `=Costs!B7` because a row was inserted above it: expected, **not reported**.
- `=SUM(B2:B5)` that should have become `=SUM(B2:B6)` but is still `=SUM(B2:B5)`: **range no longer covers all rows**.

## 4. Classifying differences

Each pair of cells is compared after alignment. The first matching rule decides the kind; for formulas every aspect is recorded as a note and the most severe one wins.

| Situation | Kind | Severity |
| --- | --- | --- |
| Cell now shows an error value, or the formula now contains `#REF!` | Error introduced | critical |
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
| Same formula, different cached result | Result recalculated | info |

When a formula is newly wrapped in `IFERROR(x, fallback)`, the wrapper is reported once and `x` is compared with the expected formula on its own — so the fallback `0` is not also reported as a changed constant.

**Rows and sheets.** Inserted rows are reported once per block (info); any error cell inside them is still reported as *error introduced*. Deleted rows (medium), newly hidden rows (medium), hidden sheets (medium), unhidden sheets (low), added sheets (low) and removed sheets (high) are reported separately. Sheets are matched by name.

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
