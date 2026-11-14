# Architecture

Sheet Sentry is a static Vite app plus a small Node command line. Both use the same engine in `src/core`, which has no DOM dependencies and one runtime dependency (fflate).

```
src/
├── core/
│   ├── refs.ts      A1/R1C1 references: parse, find, shift, remap, normalise, ranges, literals, functions
│   ├── xlsx.ts      read a workbook (and write minimal ones for tests and the demo)
│   ├── align.ts     Myers LCS over row signatures, row operations, row map
│   ├── compare.ts   compare two workbooks → changes, per-sheet marks, counts, verdict
│   └── report.ts    Markdown, CSV and JSON reports
├── demo.ts          builds the Q3 budget and its AI edit as real .xlsx bytes
├── main.ts          browser UI: loading, verdict, sheet view, inspector, list, exports
├── dom.ts           text-only element helper (never parses HTML)
└── cli.ts           `npm run review` command line
```

## Data flow

```
File / bytes ──► readWorkbook ──► Workbook ─┐
                                            ├─► compareWorkbooks ──► Comparison ──► UI, reports, exit code
File / bytes ──► readWorkbook ──► Workbook ─┘
```

### `Workbook`

```ts
interface Workbook {
  sheets: Sheet[];                 // in tab order
  names: Map<string, string>;      // defined names → formula text
  hasMacros: boolean;              // xl/vbaProject.bin present
  externalLinks: number;           // xl/externalLinks/externalLink*.xml parts
}
interface Sheet {
  name: string;
  state: 'visible' | 'hidden' | 'veryHidden';
  cells: Map<string, Cell>;        // "B7" → cell
  rows: number;                    // used rows
  cols: number;                    // used columns
  hiddenRows: Set<number>;         // zero-based
}
interface Cell {
  formula?: string;                // without "=", shared formulas expanded
  value?: string | number | boolean;
  type: 'n' | 's' | 'b' | 'e';
  format?: string;                 // number format code
}
```

The reader uses a small, purpose-built XML scanner rather than a DOM parser, so it runs unchanged in Node and in the browser and never instantiates external entities.

### `Comparison`

```ts
interface Comparison {
  changes: Change[];               // most severe first
  sheets: SheetResult[];           // edited sheets in order, then removed sheets
  counts: Record<Severity, number>;
  compared: { cells: number; formulas: number };
  verdict: 'clean' | 'minor' | 'review' | 'risky';
}
interface Change {
  id: number;
  kind: Kind;                      // e.g. 'range-shrunk'
  severity: 'critical' | 'high' | 'medium' | 'low' | 'info';
  sheet: string;
  cell?: string;                   // edited address
  originalCell?: string;           // original address, when rows moved
  before?: CellState;
  after?: CellState;
  expected?: string;               // original formula through inferred insertions
  detail: string;
  notes: string[];
}
```

`SheetResult.marks` maps each edited cell to its most severe change; the UI uses it to colour the grid.

## The browser app

`main.ts` keeps a handful of module-level variables (the two loaded workbooks, the comparison, the selected sheet, change, filter and row window) and re-renders the workspace from them with `h()`. Every string from a workbook becomes a text node, so a cell containing markup is displayed, never interpreted.

Comparison runs on the main thread. Populated columns are indexed by row to avoid scanning sparse empty rectangles. The reader bounds file bytes, selected XML bytes, populated cells and aggregate row spans; these are resource limits, not a latency guarantee. Large workbooks can pause the tab. Per-slot request generations prevent late file reads from replacing a newer selection or a reset/demo, and failed replacements clear their prior workbook.

## The command line

`cli.ts` runs under `tsx`, reads both files from disk, prints `markdownReport` to standard output, optionally writes `--json` and `--csv`, and sets the exit code from `--fail-on`. Unknown, duplicate or missing options are rejected. Before writing, resolved paths and existing file identities prevent report outputs from overwriting inputs or one another, including symlink/hard-link aliases. Usage/read/write errors return `2`; `--help` returns `0`.

## Report formats

| Report | Contents |
| --- | --- |
| Markdown | heading, files and date, verdict, counts table, compared cells and formulas, a table of every non-info change |
| CSV | `severity,kind,sheet,cell,original_cell,before,after,detail` for every change; formula-like fields, including leading whitespace/control characters, are prefixed with `'`; quotes, newlines and carriage returns are escaped |
| JSON | `tool`, `generatedAt`, `files`, `verdict`, `counts`, `compared`, `sheets[] {name,status,rowsInserted,rowsDeleted}`, `changes[] {severity,kind,title,sheet,cell,originalCell,before,after,expectedFormula,detail,notes}` |

## Tests

- `tests/core.test.ts` — references, alignment, workbook round trips, shared formulas, unsupported files, every planted change in the demo, row-moved references, recalculated outputs, identical workbooks, errors in inserted rows, reports and CSV safety.
- `tests/hosting.test.mjs` — the production hosting script.
- `tests/regressions.test.ts` — malformed XML and missing parts, shared-master ordering, sparse sheets, reference semantics, type changes and safe serialization.
- `tests/cli.test.ts` — option validation, thresholds, report output and input overwrite protection including hard links.
- `e2e/app.spec.ts` — the real app in Chromium: demo, navigation between sheets, filters, file inputs and unsupported files, failed replacements, pending-read races, markup-as-text, identical files, every export, no off-origin requests during review, and phone layout.
