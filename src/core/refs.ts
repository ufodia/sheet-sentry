/**
 * A1-reference handling for Excel formulas: parsing, shifting, row remapping,
 * R1C1 normalization and the extraction of ranges, functions and literals.
 * String literals inside formulas are never touched.
 */

export interface CellRef {
  sheet?: string;
  /** 0-based column, or -1 for a whole-row reference. */
  col: number;
  /** 0-based row, or -1 for a whole-column reference. */
  row: number;
  colAbs: boolean;
  rowAbs: boolean;
}

export function colToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function indexToCol(index: number): string {
  let s = '';
  let n = index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export const address = (col: number, row: number) => `${indexToCol(col)}${row + 1}`;

export function parseAddress(a1: string): { col: number; row: number } {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(a1);
  if (!m) throw new Error(`Invalid cell address ${a1}`);
  return { col: colToIndex(m[1]), row: Number(m[2]) - 1 };
}

/** One token of a reference expression: a single cell, a range, or a whole row/column range. */
export interface RefToken {
  start: number;
  end: number;
  sheet?: string;
  a: CellRef;
  b?: CellRef;
}

const SHEET = "(?:'(?:[^']|'')+'|[A-Za-z_][\\w.]*)!";
const CELL = '\\$?[A-Za-z]{1,3}\\$?\\d+';
const COLS = '\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}';
const ROWS = '\\$?\\d+:\\$?\\d+';
const REF = new RegExp(`(${SHEET})?(?:(${CELL})(?::(${CELL}))?|(${COLS})|(${ROWS}))`, 'y');

function parseCell(text: string, sheet?: string): CellRef {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(text)!;
  return { sheet, colAbs: m[1] === '$', col: colToIndex(m[2]), rowAbs: m[3] === '$', row: Number(m[4]) - 1 };
}

function unquoteSheet(raw?: string): string | undefined {
  if (!raw) return undefined;
  const name = raw.slice(0, -1);
  return name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
}

/** Find reference tokens, skipping string literals and function or name identifiers. */
export function findRefs(formula: string): RefToken[] {
  const tokens: RefToken[] = [];
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"') {
      i++;
      while (i < formula.length && !(formula[i] === '"' && formula[i + 1] !== '"'))
        i += formula[i] === '"' ? 2 : 1;
      i++;
      continue;
    }
    const prev = i > 0 ? formula[i - 1] : '';
    if (/[\w.$']/.test(prev) && ch !== "'") {
      i++;
      continue;
    }
    REF.lastIndex = i;
    const m = REF.exec(formula);
    if (m && m[0]) {
      const end = i + m[0].length;
      const next = formula[end] ?? '';
      // "LOG10(" is a function, "Revenue2026" a name: both end without a reference boundary.
      if (!/[\w(.]/.test(next)) {
        const sheet = unquoteSheet(m[1]);
        if (m[2])
          tokens.push({
            start: i,
            end,
            sheet,
            a: parseCell(m[2], sheet),
            b: m[3] ? parseCell(m[3], sheet) : undefined,
          });
        else if (m[4]) {
          const [x, y] = m[4].split(':');
          tokens.push({
            start: i,
            end,
            sheet,
            a: {
              sheet,
              col: colToIndex(x.replace('$', '')),
              row: -1,
              colAbs: x.startsWith('$'),
              rowAbs: false,
            },
            b: {
              sheet,
              col: colToIndex(y.replace('$', '')),
              row: -1,
              colAbs: y.startsWith('$'),
              rowAbs: false,
            },
          });
        } else if (m[5]) {
          const [x, y] = m[5].split(':');
          tokens.push({
            start: i,
            end,
            sheet,
            a: {
              sheet,
              col: -1,
              row: Number(x.replace('$', '')) - 1,
              colAbs: false,
              rowAbs: x.startsWith('$'),
            },
            b: {
              sheet,
              col: -1,
              row: Number(y.replace('$', '')) - 1,
              colAbs: false,
              rowAbs: y.startsWith('$'),
            },
          });
        }
        i = end;
        continue;
      }
    }
    i++;
  }
  return tokens;
}

function cellText(ref: CellRef): string {
  const col = ref.col >= 0 ? `${ref.colAbs ? '$' : ''}${indexToCol(ref.col)}` : '';
  const row = ref.row >= 0 ? `${ref.rowAbs ? '$' : ''}${ref.row + 1}` : '';
  return col + row;
}

function sheetText(sheet?: string): string {
  if (!sheet) return '';
  return /^[A-Za-z_][\w.]*$/.test(sheet) ? `${sheet}!` : `'${sheet.replace(/'/g, "''")}'!`;
}

/** Rewrite every reference through `map`. */
export function mapRefs(formula: string, map: (ref: CellRef) => CellRef): string {
  const tokens = findRefs(formula);
  let out = '';
  let pos = 0;
  for (const t of tokens) {
    out += formula.slice(pos, t.start);
    const a = map(t.a);
    out += sheetText(t.sheet) + cellText(a);
    if (t.b) out += ':' + cellText(map(t.b));
    pos = t.end;
  }
  return out + formula.slice(pos);
}

/** Move relative references by (dRow, dCol), as Excel does when copying a formula. */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
  return mapRefs(formula, (r) => ({
    ...r,
    row: r.row >= 0 && !r.rowAbs ? r.row + dRow : r.row,
    col: r.col >= 0 && !r.colAbs ? r.col + dCol : r.col,
  }));
}

/** Position-independent form: relative references as offsets from the formula's own cell. */
export function toR1C1(formula: string, row: number, col: number): string {
  const tokens = findRefs(formula);
  let out = '';
  let pos = 0;
  const part = (ref: CellRef) => {
    const r = ref.row < 0 ? '' : ref.rowAbs ? `R${ref.row + 1}` : `R[${ref.row - row}]`;
    const c = ref.col < 0 ? '' : ref.colAbs ? `C${ref.col + 1}` : `C[${ref.col - col}]`;
    return r + c;
  };
  for (const t of tokens) {
    out += formula.slice(pos, t.start) + sheetText(t.sheet) + part(t.a) + (t.b ? ':' + part(t.b) : '');
    pos = t.end;
  }
  return normalize(out + formula.slice(pos));
}

/** Canonical text for comparison: no whitespace outside strings, upper-case outside strings. */
export function normalize(formula: string): string {
  let out = '';
  let inString = false;
  for (const ch of formula.replace(/^=/, '')) {
    if (ch === '"') inString = !inString;
    out += inString || ch === '"' ? ch : /\s/.test(ch) ? '' : ch.toUpperCase();
  }
  return out;
}

/** Function names used in a formula, upper-case, with repeats. */
export function functions(formula: string): string[] {
  return [...stripStrings(formula).matchAll(/([A-Za-z_][\w.]*)\s*\(/g)].map((m) => m[1].toUpperCase());
}

/** Numeric literals outside references and strings. */
export function literals(formula: string): number[] {
  const masked = maskRefs(stripStrings(formula));
  return [...masked.matchAll(/(?<![\w.$])(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)(?![\w.])/g)].map((m) =>
    Number(m[1]),
  );
}

function stripStrings(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/g, (s) => ' '.repeat(s.length));
}

function maskRefs(formula: string): string {
  let out = formula;
  for (const t of findRefs(formula).reverse())
    out = out.slice(0, t.start) + ' '.repeat(t.end - t.start) + out.slice(t.end);
  return out;
}

export interface Range {
  sheet?: string;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** Ranges (and single cells as 1×1 ranges) in the formula, in order of appearance. */
export function ranges(formula: string): Range[] {
  return findRefs(formula).map((t) => {
    const b = t.b ?? t.a;
    return {
      sheet: t.sheet,
      r0: Math.min(t.a.row, b.row),
      c0: Math.min(t.a.col, b.col),
      r1: Math.max(t.a.row, b.row),
      c1: Math.max(t.a.col, b.col),
    };
  });
}

export const rangeText = (r: Range) =>
  `${sheetText(r.sheet)}${r.c0 >= 0 ? indexToCol(r.c0) : ''}${r.r0 >= 0 ? r.r0 + 1 : ''}${
    r.r1 !== r.r0 || r.c1 !== r.c0 ? `:${r.c1 >= 0 ? indexToCol(r.c1) : ''}${r.r1 >= 0 ? r.r1 + 1 : ''}` : ''
  }`;

export const cellsIn = (r: Range) => (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1);
