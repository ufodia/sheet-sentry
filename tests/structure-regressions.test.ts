import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { compareWorkbooks } from '../src/core/compare.ts';
import { readWorkbook, writeWorkbook } from '../src/core/xlsx.ts';

type Cells = Parameters<typeof writeWorkbook>[0][number]['cells'];
const originalRows = [
  ['Rent', 100],
  ['Wages', 80],
  ['Cloud', 60],
  ['Travel', 40],
  ['Misc', 20],
] as const;
function budget(rows: readonly (readonly [string, number])[], formula: string): Uint8Array {
  const cells: Cells = { A1: 'Item', B1: 'Amount' };
  rows.forEach(([label, value], i) => {
    cells[`A${i + 2}`] = label;
    cells[`B${i + 2}`] = value;
  });
  cells[`A${rows.length + 2}`] = 'Total';
  cells[`B${rows.length + 2}`] = { f: formula, v: rows.reduce((sum, [, value]) => sum + value, 0) };
  return writeWorkbook([{ name: 'Budget', cells }]);
}
const before = budget(originalRows, 'SUM(B2:B6)');
function compare(edited: Uint8Array) {
  return compareWorkbooks(readWorkbook(before), readWorkbook(edited));
}

describe('openpyxl and structural budget regressions', () => {
  it('reads an actual openpyxl resave with empty formula caches as clean', () => {
    const original = readWorkbook(readFileSync('examples/q3-budget.xlsx'));
    const edited = readWorkbook(readFileSync('tests/fixtures/q3-budget-openpyxl.xlsx'));
    expect(edited.sheets[0].cells.get('E2')?.value).toBeUndefined();
    expect(edited.sheets[0].cells.get('E2')?.formula).toBeDefined();
    expect(compareWorkbooks(original, edited).verdict).toBe('clean');
  });
  it('does not claim a replacement matches an unavailable formula result', () => {
    const noCache = readWorkbook(readFileSync('tests/fixtures/budget-before.xlsx'));
    const cells: Cells = { A1: 'Item', B1: 'Amount', A7: 'Total', B7: '' };
    originalRows.forEach(([label, value], i) => {
      cells[`A${i + 2}`] = label;
      cells[`B${i + 2}`] = value;
    });
    const blank = readWorkbook(writeWorkbook([{ name: 'Budget', cells }]));
    expect(compareWorkbooks(noCache, blank).changes.find((c) => c.cell === 'B7')?.kind).toBe(
      'formula-to-value',
    );
  });
  const fixture = (name: string) => readWorkbook(readFileSync(`tests/fixtures/${name}.xlsx`));
  const independentBefore = fixture('budget-before');
  for (const position of ['head', 'middle', 'tail']) {
    it(`aligns insertion at ${position} and reports an omitted row`, () => {
      const correct = compareWorkbooks(independentBefore, fixture(`insert-${position}-correct`));
      expect(
        correct.changes.filter((c) => ['critical', 'high', 'medium', 'low'].includes(c.severity)),
      ).toEqual([]);
      const bad = compareWorkbooks(independentBefore, fixture(`insert-${position}-wrong`));
      expect(bad.changes.some((c) => c.kind === 'range-shrunk' && c.cell === 'B8')).toBe(true);
    });
    it(`aligns deletion at ${position} and distinguishes an incorrect total`, () => {
      const correct = compareWorkbooks(independentBefore, fixture(`delete-${position}-correct`));
      expect(correct.changes.filter((c) => c.cell === 'B6' && c.severity !== 'info')).toEqual([]);
      expect(
        correct.changes.filter(
          (c) => c.kind === 'input-changed' || c.kind === 'text-changed' || c.kind === 'formula-to-value',
        ),
      ).toEqual([]);
      const bad = compareWorkbooks(independentBefore, fixture(`delete-${position}-wrong`));
      expect(
        bad.changes.some(
          (c) => c.cell === 'B6' && ['range-grew', 'references-changed', 'range-shrunk'].includes(c.kind),
        ),
      ).toBe(true);
    });
  }
  it('still detects changed inputs and a hardcoded total when labels are stable', () => {
    const modified = [...originalRows];
    modified[1] = ['Wages', 999] as never;
    expect(
      compare(budget(modified, 'SUM(B2:B6)')).changes.some(
        (c) => c.kind === 'input-changed' && c.cell === 'B3',
      ),
    ).toBe(true);
    const cells: Cells = { A1: 'Item', B1: 'Amount', A7: 'Total', B7: 300 };
    originalRows.forEach(([label, value], i) => {
      cells[`A${i + 2}`] = label;
      cells[`B${i + 2}`] = value;
    });
    expect(
      compare(writeWorkbook([{ name: 'Budget', cells }])).changes.some(
        (c) => c.kind === 'formula-hardcoded' && c.cell === 'B7',
      ),
    ).toBe(true);
  });
  it.each(['note', 'empty'])('does not expand an edge subtotal for an inserted %s row', (kind) => {
    const cells: Cells = { A1: 'Item', B1: 'Amount' };
    originalRows.forEach(([label, value], i) => {
      cells[`A${i + 2}`] = label;
      cells[`B${i + 2}`] = value;
    });
    cells.A7 = 'Comment';
    if (kind === 'note') cells.B7 = 'not an amount';
    cells.A8 = 'Total';
    cells.B8 = { f: 'SUM(B2:B6)', v: 300 };
    const result = compare(writeWorkbook([{ name: 'Budget', cells }]));
    expect(
      result.changes.some((c) => c.cell === 'B8' && ['range-shrunk', 'range-grew'].includes(c.kind)),
    ).toBe(false);
    expect(result.changes.some((c) => c.kind === 'rows-inserted')).toBe(true);
  });
});
