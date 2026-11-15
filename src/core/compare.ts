import { alignRows, rowMapper, type RowOp } from './align.ts';
import {
  address,
  cellsIn,
  functions,
  indexToCol,
  literals,
  mapRefs,
  normalize,
  parseAddress,
  rangeText,
  ranges,
  toR1C1,
} from './refs.ts';
import type { Cell, Sheet, Workbook } from './xlsx.ts';

export type Severity = 'critical' | 'high' | 'medium' | 'low' | 'info';

export type Kind =
  | 'formula-hardcoded'
  | 'formula-to-value'
  | 'error-introduced'
  | 'range-shrunk'
  | 'error-masked'
  | 'literal-changed'
  | 'references-changed'
  | 'pattern-break'
  | 'formula-edited'
  | 'range-grew'
  | 'input-changed'
  | 'text-changed'
  | 'value-to-formula'
  | 'cell-added'
  | 'cell-removed'
  | 'format-changed'
  | 'output-changed'
  | 'row-hidden'
  | 'rows-inserted'
  | 'rows-deleted'
  | 'sheet-added'
  | 'sheet-removed'
  | 'sheet-hidden'
  | 'sheet-shown'
  | 'name-changed'
  | 'name-added'
  | 'name-removed'
  | 'macros-added'
  | 'external-link-added';

export const SEVERITY: Record<Kind, Severity> = {
  'formula-hardcoded': 'critical',
  'formula-to-value': 'critical',
  'error-introduced': 'critical',
  'range-shrunk': 'critical',
  'error-masked': 'high',
  'literal-changed': 'high',
  'references-changed': 'high',
  'pattern-break': 'high',
  'sheet-removed': 'high',
  'name-changed': 'high',
  'macros-added': 'high',
  'external-link-added': 'high',
  'formula-edited': 'medium',
  'range-grew': 'medium',
  'input-changed': 'medium',
  'cell-removed': 'medium',
  'rows-deleted': 'medium',
  'row-hidden': 'medium',
  'sheet-hidden': 'medium',
  'name-removed': 'medium',
  'text-changed': 'low',
  'value-to-formula': 'low',
  'cell-added': 'low',
  'format-changed': 'low',
  'sheet-added': 'low',
  'sheet-shown': 'low',
  'name-added': 'low',
  'output-changed': 'info',
  'rows-inserted': 'info',
};

export const TITLE: Record<Kind, string> = {
  'formula-hardcoded': 'Formula replaced by its own result',
  'formula-to-value': 'Formula replaced by a constant',
  'error-introduced': 'Error introduced',
  'range-shrunk': 'Range no longer covers all rows',
  'error-masked': 'Errors now hidden (IFERROR)',
  'literal-changed': 'Number inside a formula changed',
  'references-changed': 'Formula points somewhere else',
  'pattern-break': 'Formula breaks its row or column pattern',
  'formula-edited': 'Formula edited',
  'range-grew': 'Range grew',
  'input-changed': 'Input value changed',
  'text-changed': 'Text changed',
  'value-to-formula': 'Constant replaced by a formula',
  'cell-added': 'Cell added',
  'cell-removed': 'Cell cleared',
  'format-changed': 'Number format changed',
  'output-changed': 'Result recalculated',
  'row-hidden': 'Row hidden',
  'rows-inserted': 'Rows inserted',
  'rows-deleted': 'Rows deleted',
  'sheet-added': 'Sheet added',
  'sheet-removed': 'Sheet removed',
  'sheet-hidden': 'Sheet hidden',
  'sheet-shown': 'Sheet unhidden',
  'name-changed': 'Named range redefined',
  'name-added': 'Named range added',
  'name-removed': 'Named range removed',
  'macros-added': 'Macros added',
  'external-link-added': 'External workbook link added',
};

export interface CellState {
  formula?: string;
  value?: string | number | boolean;
  format?: string;
}

export interface Change {
  id: number;
  kind: Kind;
  severity: Severity;
  sheet: string;
  /** Address in the edited workbook (or the original for removed content). */
  cell?: string;
  originalCell?: string;
  before?: CellState;
  after?: CellState;
  /** Original formula remapped through inferred row insertions, when that differs. */
  expected?: string;
  detail: string;
  /** Every aspect detected for this cell, most severe first. */
  notes: string[];
}

export interface SheetResult {
  name: string;
  status: 'matched' | 'added' | 'removed';
  inserted: number;
  deleted: number;
  /** Edited-sheet addresses → most severe change on that cell. */
  marks: Map<string, Change>;
  sheet?: Sheet;
}

export interface Comparison {
  changes: Change[];
  sheets: SheetResult[];
  counts: Record<Severity, number>;
  compared: { cells: number; formulas: number };
  verdict: 'clean' | 'minor' | 'review' | 'risky';
}

const ORDER: Severity[] = ['critical', 'high', 'medium', 'low', 'info'];
const rank = (s: Severity) => ORDER.indexOf(s);

const valueText = (v: Cell['value']) =>
  v === undefined ? '' : typeof v === 'number' ? String(+v.toPrecision(12)) : String(v);

/** Index populated columns once; sparse sheets must not scan their entire bounding rectangle. */
function rowColumns(sheet: Sheet): Map<number, number[]> {
  const rows = new Map<number, number[]>();
  for (const ref of sheet.cells.keys()) {
    const { row, col } = parseAddress(ref);
    const columns = rows.get(row) ?? [];
    columns.push(col);
    rows.set(row, columns);
  }
  for (const columns of rows.values()) columns.sort((a, b) => a - b);
  return rows;
}

function textAnchors(sheet: Sheet): Map<string, number> {
  const counts = new Map<string, number>();
  for (const [ref, cell] of sheet.cells) {
    if (
      cell.formula !== undefined ||
      cell.type !== 's' ||
      typeof cell.value !== 'string' ||
      !cell.value.trim()
    )
      continue;
    const key = JSON.stringify([parseAddress(ref).col, cell.value]);
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  return counts;
}

function rowSignatures(
  sheet: Sheet,
  columns: Map<number, number[]>,
  ownAnchors: Map<string, number>,
  otherAnchors: Map<string, number>,
): string[] {
  const signatures = Array<string>(sheet.rows).fill('');
  for (const [row, cols] of columns) {
    const anchor = cols.find((col) => {
      const cell = sheet.cells.get(address(col, row))!;
      if (
        cell.formula !== undefined ||
        cell.type !== 's' ||
        typeof cell.value !== 'string' ||
        !cell.value.trim()
      )
        return false;
      const key = JSON.stringify([col, cell.value]);
      return ownAnchors.get(key) === 1 && otherAnchors.get(key) === 1;
    });
    if (anchor !== undefined) {
      signatures[row] = JSON.stringify(['label', anchor, sheet.cells.get(address(anchor, row))!.value]);
      continue;
    }
    signatures[row] = JSON.stringify(
      cols.map((col) => {
        const cell = sheet.cells.get(address(col, row))!;
        return cell.formula !== undefined
          ? [col, 'f', toR1C1(cell.formula, row, col)]
          : [col, cell.type, cell.value];
      }),
    );
  }
  return signatures;
}

function sameNumber(a: unknown, b: unknown): boolean {
  if (typeof a === 'number' && typeof b === 'number')
    return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));
  return a === b;
}

const sameValue = (a: Cell, b: Cell) => a.type === b.type && sameNumber(a.value, b.value);
const referenceError = (cell: Cell) =>
  cell.type === 'e' || /#REF!/i.test((cell.formula ?? '').replace(/"(?:[^"]|"")*"/g, ''));

const describeValue = (v: Cell['value']) =>
  v === undefined ? '(no cached value)' : typeof v === 'string' ? `"${v}"` : valueText(v);

/** Formula cells whose R1C1 form differs from both neighbours that agree with each other. */
function patternBreaks(sheet: Sheet): Map<string, string> {
  const r1c1 = new Map<string, string>();
  for (const [ref, cell] of sheet.cells)
    if (cell.formula !== undefined) {
      const { row, col } = parseAddress(ref);
      r1c1.set(ref, toR1C1(cell.formula, row, col));
    }
  const breaks = new Map<string, string>();
  for (const [ref, form] of r1c1) {
    const { row, col } = parseAddress(ref);
    for (const [d1, d2, axis] of [
      [address(col - 1, row), address(col + 1, row), 'row'],
      [address(col, row - 1), address(col, row + 1), 'column'],
    ] as const) {
      if (col === 0 && axis === 'row') continue;
      if (row === 0 && axis === 'column') continue;
      const a = r1c1.get(d1);
      const b = r1c1.get(d2);
      if (a && a === b && a !== form)
        breaks.set(ref, `${d1} and ${d2} share one formula pattern; this cell does not`);
    }
  }
  return breaks;
}

interface FormulaVerdict {
  kind: Kind;
  notes: string[];
}

/** The first argument of an outermost IFERROR/IFNA wrapper, or the whole formula when there is none. */
export function unwrapMask(formula: string): string {
  const f = normalize(formula);
  const head = /^(IFERROR|IFNA)\(/.exec(f);
  if (!head) return f;
  let depth = 1;
  let comma = -1;
  let quoted = false;
  for (let i = head[0].length; i < f.length; i++) {
    const ch = f[i];
    if (ch === '"') quoted = !quoted;
    if (quoted) continue;
    if (ch === '(') depth++;
    else if (ch === ')' && --depth === 0)
      return i === f.length - 1 && comma > 0 ? f.slice(head[0].length, comma) : f;
    else if (ch === ',' && depth === 1 && comma < 0) comma = i;
  }
  return f;
}

/** Compare the formula Excel would have after the row changes with the formula actually found. */
function formulaDifference(expected: string, found: string): FormulaVerdict | null {
  if (normalize(expected) === normalize(found)) return null;
  const notes: { kind: Kind; text: string }[] = [];
  const masking = (f: string) =>
    functions(f).filter((x) => x === 'IFERROR' || x === 'IFNA' || x === 'ISERROR').length;
  // A new IFERROR(x, fallback) is reported once; x is then compared on its own so the fallback is not a "changed number".
  const masked = masking(found) > masking(expected);
  const actual = masked ? unwrapMask(found) : found;
  if (masked)
    notes.push({ kind: 'error-masked', text: 'Wrapped in IFERROR/IFNA: errors now show as a normal value' });
  const le = literals(expected).sort((a, b) => a - b);
  const la = literals(actual).sort((a, b) => a - b);
  if (le.join(',') !== la.join(','))
    notes.push({
      kind: 'literal-changed',
      text: `Numbers in the formula changed: ${le.join(', ') || 'none'} → ${la.join(', ') || 'none'}`,
    });
  const re = ranges(expected);
  const ra = ranges(actual);
  if (re.length === ra.length) {
    re.forEach((e, i) => {
      const a = ra[i];
      if ((e.sheet ?? '') !== (a.sheet ?? '')) {
        notes.push({ kind: 'references-changed', text: `${rangeText(e)} → ${rangeText(a)}` });
        return;
      }
      if (e.r0 === a.r0 && e.r1 === a.r1 && e.c0 === a.c0 && e.c1 === a.c1) return;
      const inside = a.r0 >= e.r0 && a.r1 <= e.r1 && a.c0 >= e.c0 && a.c1 <= e.c1;
      const around = a.r0 <= e.r0 && a.r1 >= e.r1 && a.c0 <= e.c0 && a.c1 >= e.c1;
      if (inside && cellsIn(e) > 1)
        notes.push({
          kind: 'range-shrunk',
          text: `${rangeText(e)} became ${rangeText(a)}: ${cellsIn(e) - cellsIn(a)} cell(s) dropped`,
        });
      else if (around) notes.push({ kind: 'range-grew', text: `${rangeText(e)} became ${rangeText(a)}` });
      else notes.push({ kind: 'references-changed', text: `${rangeText(e)} → ${rangeText(a)}` });
    });
  } else notes.push({ kind: 'references-changed', text: `${re.length} reference(s) → ${ra.length}` });
  if (!notes.length && normalize(expected) !== normalize(actual))
    notes.push({ kind: 'formula-edited', text: 'The formula was rewritten' });
  notes.sort((x, y) => rank(SEVERITY[x.kind]) - rank(SEVERITY[y.kind]));
  return { kind: notes[0].kind, notes: notes.map((n) => n.text) };
}

export function compareWorkbooks(original: Workbook, edited: Workbook): Comparison {
  const changes: Change[] = [];
  let id = 0;
  const add = (c: Omit<Change, 'id' | 'severity' | 'notes'> & { notes?: string[] }) => {
    const change: Change = { id: ++id, severity: SEVERITY[c.kind], notes: c.notes ?? [c.detail], ...c };
    changes.push(change);
    return change;
  };

  // Workbook structure.
  if (edited.hasMacros && !original.hasMacros)
    add({
      kind: 'macros-added',
      sheet: '(workbook)',
      detail: 'The edited file contains a VBA project; the original did not.',
    });
  if (edited.externalLinks > original.externalLinks)
    add({
      kind: 'external-link-added',
      sheet: '(workbook)',
      detail: `${edited.externalLinks - original.externalLinks} link(s) to other workbooks were added.`,
    });
  for (const [name, f] of edited.names) {
    const before = original.names.get(name);
    if (before === undefined) add({ kind: 'name-added', sheet: '(names)', detail: `${name} = ${f}` });
    else if (normalize(before) !== normalize(f))
      add({ kind: 'name-changed', sheet: '(names)', detail: `${name}: ${before} → ${f}` });
  }
  for (const [name, f] of original.names)
    if (!edited.names.has(name))
      add({ kind: 'name-removed', sheet: '(names)', detail: `${name} (was ${f})` });

  const originalByName = new Map(original.sheets.map((s) => [s.name, s]));
  const editedByName = new Map(edited.sheets.map((s) => [s.name, s]));

  // Row alignment for every matched sheet, so cross-sheet references can be remapped too.
  const ops = new Map<string, RowOp[]>();
  const mappers = new Map<string, (row: number) => number>();
  const deletedRows = new Map<string, Set<number>>();
  const survivingRows = new Map<string, number[]>();
  const insertedRows = new Map<string, Set<number>>();
  const columns = new Map<Sheet, Map<number, number[]>>();
  for (const sheet of [...original.sheets, ...edited.sheets]) columns.set(sheet, rowColumns(sheet));
  for (const sheet of original.sheets) {
    const other = editedByName.get(sheet.name);
    if (!other) continue;
    const anchorsA = textAnchors(sheet);
    const anchorsB = textAnchors(other);
    const sa = rowSignatures(sheet, columns.get(sheet)!, anchorsA, anchorsB);
    const sb = rowSignatures(other, columns.get(other)!, anchorsB, anchorsA);
    const o = alignRows(sa, sb);
    ops.set(sheet.name, o);
    const key = sheet.name.toLowerCase();
    deletedRows.set(key, new Set(o.flatMap((op) => (op.type === 'deleted' ? [op.a] : []))));
    survivingRows.set(
      key,
      o.flatMap((op) => (op.type === 'same' || op.type === 'modified' ? [op.a] : [])),
    );
    insertedRows.set(key, new Set(o.flatMap((op) => (op.type === 'inserted' ? [op.b] : []))));
    mappers.set(key, rowMapper(o, sheet.rows));
  }
  const identity = (r: number) => r;

  const results: SheetResult[] = [];
  let cellsCompared = 0;
  let formulasCompared = 0;

  for (const sheet of edited.sheets) {
    const before = originalByName.get(sheet.name);
    const result: SheetResult = {
      name: sheet.name,
      status: before ? 'matched' : 'added',
      inserted: 0,
      deleted: 0,
      marks: new Map(),
      sheet,
    };
    results.push(result);
    const mark = (c: Change) => {
      if (!c.cell) return;
      const prev = result.marks.get(c.cell);
      if (!prev || rank(c.severity) < rank(prev.severity)) result.marks.set(c.cell, c);
    };
    if (!before) {
      add({ kind: 'sheet-added', sheet: sheet.name, detail: `${sheet.cells.size} cell(s)` });
      for (const [ref, cell] of sheet.cells)
        if (referenceError(cell))
          mark(
            add({
              kind: 'error-introduced',
              sheet: sheet.name,
              cell: ref,
              after: state(cell),
              detail: `A new sheet contains an error at ${ref}.`,
            }),
          );
      continue;
    }
    if (before.state === 'visible' && sheet.state !== 'visible')
      add({
        kind: 'sheet-hidden',
        sheet: sheet.name,
        detail:
          sheet.state === 'veryHidden'
            ? 'Set to very hidden: it does not appear in the Unhide list.'
            : 'The sheet tab is now hidden.',
      });
    if (before.state !== 'visible' && sheet.state === 'visible')
      add({ kind: 'sheet-shown', sheet: sheet.name, detail: 'A hidden sheet is now visible.' });

    const rowOps = ops.get(sheet.name)!;
    const remap = (row: number, sheetName?: string) =>
      (mappers.get((sheetName ?? sheet.name).toLowerCase()) ?? identity)(row);
    const breaksBefore = patternBreaks(before);
    const breaksAfter = patternBreaks(sheet);

    // Inserted and deleted row blocks.
    const blocks = (type: 'inserted' | 'deleted') => {
      const rows = rowOps
        .filter((o) => o.type === type)
        .map((o) => (o.type === 'inserted' ? o.b : (o as { a: number }).a));
      const out: [number, number][] = [];
      for (const r of rows) {
        const last = out[out.length - 1];
        if (last && r === last[1] + 1) last[1] = r;
        else out.push([r, r]);
      }
      return out;
    };
    for (const [r0, r1] of blocks('inserted')) {
      result.inserted += r1 - r0 + 1;
      add({
        kind: 'rows-inserted',
        sheet: sheet.name,
        cell: address(0, r0),
        detail: r0 === r1 ? `Row ${r0 + 1}` : `Rows ${r0 + 1}–${r1 + 1}`,
      });
      // New rows have no "before", but an error in them is still new.
      for (let r = r0; r <= r1; r++)
        for (const c of columns.get(sheet)!.get(r) ?? []) {
          const ref = address(c, r);
          const cell = sheet.cells.get(ref);
          if (!cell) continue;
          if (referenceError(cell))
            mark(
              add({
                kind: 'error-introduced',
                sheet: sheet.name,
                cell: ref,
                after: { formula: cell.formula, value: cell.value, format: cell.format },
                detail: `An inserted row shows ${cell.type === 'e' ? cell.value : '#REF!'}${cell.formula ? ` from =${cell.formula}` : ''}`,
              }),
            );
        }
    }
    for (const [r0, r1] of blocks('deleted')) {
      result.deleted += r1 - r0 + 1;
      add({
        kind: 'rows-deleted',
        sheet: sheet.name,
        originalCell: address(0, r0),
        detail: r0 === r1 ? `Original row ${r0 + 1}` : `Original rows ${r0 + 1}–${r1 + 1}`,
      });
    }

    for (const op of rowOps) {
      if (op.type !== 'same' && op.type !== 'modified') continue;
      if (sheet.hiddenRows.has(op.b) && !before.hiddenRows.has(op.a)) {
        mark(
          add({
            kind: 'row-hidden',
            sheet: sheet.name,
            cell: address(0, op.b),
            detail: `Row ${op.b + 1} is now hidden; its contents are invisible on screen and in most printouts.`,
          }),
        );
      }
      const cols = new Set([
        ...(columns.get(before)!.get(op.a) ?? []),
        ...(columns.get(sheet)!.get(op.b) ?? []),
      ]);
      for (const c of [...cols].sort((a, b) => a - b)) {
        const ref = address(c, op.b);
        const oref = address(c, op.a);
        const a = before.cells.get(oref);
        const b = sheet.cells.get(ref);
        if (!a && !b) continue;
        cellsCompared++;
        const base = {
          sheet: sheet.name,
          cell: ref,
          originalCell: oref,
          before: a && state(a),
          after: b && state(b),
        };
        if (a && !b) {
          add({
            kind: 'cell-removed',
            ...base,
            cell: undefined,
            detail: `${oref} held ${a.formula !== undefined ? '=' + a.formula : describeValue(a.value)}`,
          });
          continue;
        }
        if (!a && b) {
          mark(
            add({
              kind: referenceError(b) ? 'error-introduced' : 'cell-added',
              ...base,
              detail:
                b.formula !== undefined ? `New formula =${b.formula}` : `New value ${describeValue(b.value)}`,
            }),
          );
          continue;
        }
        const x = a!;
        const y = b!;
        if (y.type === 'e' && x.type !== 'e') {
          mark(
            add({
              kind: 'error-introduced',
              ...base,
              detail: `Now shows ${y.value}${y.formula ? ` from =${y.formula}` : ''}`,
            }),
          );
          continue;
        }
        if (referenceError(y) && !referenceError(x)) {
          mark(
            add({ kind: 'error-introduced', ...base, detail: `The formula lost a reference: =${y.formula}` }),
          );
          continue;
        }
        if (x.formula !== undefined && y.formula === undefined) {
          formulasCompared++;
          const own =
            x.value !== undefined &&
            y.value !== undefined &&
            (sameNumber(x.value, y.value) || valueText(x.value) === valueText(y.value));
          mark(
            add({
              kind: own ? 'formula-hardcoded' : 'formula-to-value',
              ...base,
              detail: own
                ? `=${x.formula} was replaced by its last result ${describeValue(y.value)}. It will no longer update.`
                : `=${x.formula} was replaced by the constant ${describeValue(y.value)} (it calculated ${describeValue(x.value)}).`,
            }),
          );
          continue;
        }
        if (x.formula === undefined && y.formula !== undefined) {
          mark(
            add({
              kind: 'value-to-formula',
              ...base,
              detail: `${describeValue(x.value)} is now =${y.formula}`,
            }),
          );
          continue;
        }
        if (x.formula !== undefined && y.formula !== undefined) {
          formulasCompared++;
          const simpleAggregate =
            /^(SUM|AVERAGE|COUNT|MIN|MAX)\(\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+\)$/.test(
              normalize(x.formula),
            );
          let ambiguousDeletion = false;
          const expected = mapRefs(x.formula, (r, token, endpoint) => {
            if (r.row < 0) return r;
            const key = (r.sheet ?? sheet.name).toLowerCase();
            const deleted = deletedRows.get(key);
            const surviving = survivingRows.get(key);
            let sourceRow = r.row;
            if (deleted?.has(sourceRow)) {
              if (!token.b || token.a.row < 0 || token.b.row < 0) {
                ambiguousDeletion = true;
                return r;
              }
              const low = Math.min(token.a.row, token.b.row);
              const high = Math.max(token.a.row, token.b.row);
              const rows = surviving ?? [];
              let left = 0;
              let right = rows.length;
              while (left < right) {
                const mid = (left + right) >>> 1;
                if (rows[mid] < sourceRow) left = mid + 1;
                else right = mid;
              }
              sourceRow = (sourceRow === low ? rows[left] : rows[left - 1]) ?? -1;
              if (sourceRow < low || sourceRow > high) {
                ambiguousDeletion = true;
                return r;
              }
            }
            let mapped = remap(sourceRow, r.sheet);
            // For a simple subtotal immediately below a one-column range, newly
            // inserted numeric rows adjoining either edge are part of that range.
            if (
              simpleAggregate &&
              token.b &&
              !token.sheet &&
              token.a.col === token.b.col &&
              token.a.row >= 0 &&
              token.b.row === op.a - 1 &&
              token.a.row <= token.b.row
            ) {
              const inserted = insertedRows.get(key);
              const direction = endpoint === 'a' ? -1 : 1;
              let adjacent = mapped + direction;
              while (inserted?.has(adjacent) && adjacent >= 0 && adjacent < op.b) {
                const cell = sheet.cells.get(address(r.col, adjacent));
                if (!cell || (typeof cell.value !== 'number' && cell.formula === undefined)) break;
                mapped = adjacent;
                adjacent += direction;
              }
            }
            return { ...r, row: mapped };
          });
          if (ambiguousDeletion) {
            mark(
              add({
                kind: 'references-changed',
                ...base,
                detail: 'A referenced row was deleted. Review this reference against the original.',
              }),
            );
            continue;
          }
          const diff = formulaDifference(expected, y.formula);
          if (diff) {
            mark(
              add({
                kind: diff.kind,
                ...base,
                expected: expected !== x.formula ? expected : undefined,
                detail: `=${expected} → =${y.formula}`,
                notes: diff.notes,
              }),
            );
            continue;
          }
          if (breaksAfter.has(ref) && !breaksBefore.has(oref)) {
            mark(add({ kind: 'pattern-break', ...base, detail: breaksAfter.get(ref)! }));
            continue;
          }
          if (x.value !== undefined && y.value !== undefined && !sameValue(x, y))
            mark(
              add({
                kind: 'output-changed',
                ...base,
                detail: `${describeValue(x.value)} → ${describeValue(y.value)}`,
              }),
            );
        } else if (!sameValue(x, y)) {
          const numeric = typeof x.value === 'number' && typeof y.value === 'number';
          const pct =
            numeric && x.value !== 0
              ? ` (${((((y.value as number) - (x.value as number)) / Math.abs(x.value as number)) * 100).toFixed(1)} %)`
              : '';
          mark(
            add({
              kind: numeric ? 'input-changed' : 'text-changed',
              ...base,
              detail: `${describeValue(x.value)} → ${describeValue(y.value)}${pct}`,
            }),
          );
          continue;
        }
        if ((x.format ?? 'General') !== (y.format ?? 'General') && !result.marks.has(ref))
          mark(
            add({
              kind: 'format-changed',
              ...base,
              detail: `${x.format ?? 'General'} → ${y.format ?? 'General'}`,
            }),
          );
      }
    }
  }
  for (const sheet of original.sheets)
    if (!editedByName.has(sheet.name)) {
      results.push({
        name: sheet.name,
        status: 'removed',
        inserted: 0,
        deleted: sheet.rows,
        marks: new Map(),
      });
      add({
        kind: 'sheet-removed',
        sheet: sheet.name,
        detail: `${sheet.cells.size} cell(s), ${[...sheet.cells.values()].filter((c) => c.formula).length} formula(s)`,
      });
    }

  changes.sort((a, b) => rank(a.severity) - rank(b.severity) || a.id - b.id);
  const counts = Object.fromEntries(
    ORDER.map((s) => [s, changes.filter((c) => c.severity === s).length]),
  ) as Record<Severity, number>;
  const verdict = counts.critical
    ? 'risky'
    : counts.high
      ? 'review'
      : counts.medium || counts.low
        ? 'minor'
        : 'clean';
  return {
    changes,
    sheets: results,
    counts,
    compared: { cells: cellsCompared, formulas: formulasCompared },
    verdict,
  };
}

function state(c: Cell): CellState {
  return { formula: c.formula, value: c.value, format: c.format };
}

export { indexToCol };
