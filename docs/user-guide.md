# User guide

## 1. Load two workbooks

| Slot | What to put there |
| --- | --- |
| **1 · Original workbook** | the version you trust — before you gave it to the assistant |
| **2 · AI-edited workbook** | the version the assistant, agent or colleague returned |

Click a slot to choose a file, drop a file on it, or drop **both files at once** anywhere on the page (the first becomes the original). The comparison starts as soon as both slots are filled. Each slot shows the file name, its number of sheets and cells.

Supported: `.xlsx` and `.xlsm` up to 50 MB. For `.xls`, `.xlsb`, `.ods` or Google Sheets, save or download a copy as `.xlsx` first.

**No files at hand?** Click **Review the demo budget**. **Download the demo files** saves the same two workbooks so you can open them in Excel and compare with what Sheet Sentry reports.

## 2. Read the verdict

The bar at the top names both files and gives one of four verdicts:

| Verdict | Means |
| --- | --- |
| **Risky edit** | at least one critical change — hardcoded formulas, shrunken ranges or new errors |
| **Needs review** | formulas, names or structure changed in ways that deserve a second look |
| **Minor changes** | values or presentation changed; no formula risks found |
| **No meaningful changes** | no review findings within the supported scope; a correctness guarantee is not implied |

The five counters show how many changes each severity has. **Click a counter** to list only that severity; click it again to return to the full list.

## 3. Review changes

**Changes to review** (right, or below on a phone) lists up to 400 matching changes from most to least severe. Export CSV or JSON for the complete list. Each entry shows the kind, sheet and cell, and an explanation. Click an entry to open its sheet and select its cell.

Cached-result differences under an unchanged or insertion-remapped formula are info findings, hidden by default. The app does not recalculate or establish their cause. Tick **Show recalculated results** to see them; they then also get a thin grey outline on the sheet.

### The sheet view

The edited version of each sheet is shown with the cells that changed outlined in their severity colour:

| Colour | Severity |
| --- | --- |
| red | critical |
| orange | high |
| yellow | medium |
| blue | low |
| grey | info |

Sheet tabs carry a dot in the colour of their most severe change; hidden and removed sheets are labelled. Hidden rows are striped. Formulas are shown in green; hover a cell to see its formula. Large sheets are shown 120 rows at a time (use **Previous rows** / **Next rows**) and up to column AN; every cell is compared regardless.

### The inspector

Select an outlined cell to see:

- **Before** — the original formula (or *(constant)*) and its value. When rows were inserted or deleted, the label shows where the cell was in the original, e.g. *Before · B6*.
- **After** — the edited formula or constant and its value.
- **Expected after the row changes** — a formula inferred from supported row insertions. Check the row alignment before treating it as the intended formula. Deleted-row references are flagged conservatively for manual review.
- **Notes** — every aspect detected, e.g. *B2:B6 became B2:B5: 1 cell(s) dropped*.

## 4. Export

| Button | File | Use it for |
| --- | --- | --- |
| **Review summary (Markdown)** | `sheet-sentry-review.md` | a pull request, ticket, chat message or email |
| **All changes (CSV)** | `sheet-sentry-changes.csv` | filtering and sign-off in a spreadsheet |
| **Full report (JSON)** | `sheet-sentry-report.json` | archiving, scripts and dashboards |

Files are generated inside the tab and saved by your browser. **New comparison** clears everything.

## 5. Command line

```bash
npm run review -- original.xlsx edited.xlsx --fail-on high --csv changes.csv --json report.json
```

The Markdown review goes to standard output. Exit code `0` means nothing reached `--fail-on` (default `critical`), `1` means something did, `2` means a usage or file error. See the README for a GitHub Actions example.

## Troubleshooting

| Message or symptom | What to do |
| --- | --- |
| *…legacy .xls file or a password-protected workbook…* | save an unprotected copy as `.xlsx` in Excel or LibreOffice |
| *could not be unzipped* / *no workbook part* | the file is not an Excel workbook, or it is damaged |
| *larger than 50 MB* / *more than 3 million cells* | split the workbook or remove unused sheets |
| *combined worksheet row spans exceed 1.5 million rows* | remove distant unused rows or split the file; the budget includes each sheet's highest populated or hidden row |
| *missing or unsupported* / *shared formula … has no definition* | resave an unprotected copy in Excel; the reader refuses to compare incomplete or unsupported parts |
| Values are empty but formulas are shown | the file was saved without calculated values (common for files written by scripts); open and save it in Excel |
| Almost every cell in a sheet is reported | a column was inserted or deleted, or the sheet was rebuilt; columns are compared by position |
| A sheet shows as removed and another as added | the sheet was renamed |
| References are flagged after row deletion | verify them manually; the tool intentionally avoids equating a deleted cell with its next surviving neighbour |
