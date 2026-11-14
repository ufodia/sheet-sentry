import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { alignRows, lcsPairs, rowMapper } from '../src/core/align.ts';
import { compareWorkbooks, unwrapMask, type Kind } from '../src/core/compare.ts';
import { csvReport, jsonReport, markdownReport } from '../src/core/report.ts';
import {
  colToIndex,
  findRefs,
  functions,
  indexToCol,
  literals,
  mapRefs,
  normalize,
  ranges,
  shiftFormula,
  toR1C1,
} from '../src/core/refs.ts';
import { readWorkbook, writeWorkbook, WorkbookError } from '../src/core/xlsx.ts';
import { demoWorkbooks } from '../src/demo.ts';

describe('references', () => {
  it('converts column letters', () => {
    expect(colToIndex('A')).toBe(0);
    expect(colToIndex('Z')).toBe(25);
    expect(colToIndex('AA')).toBe(26);
    expect(indexToCol(16383)).toBe('XFD');
  });

  it('finds references but not functions, names or strings', () => {
    const f = 'SUM(B2:B10)+LOG10(A1)*TaxRate+"A1 in text"+\'My Sheet\'!$C$3+Sheet2!D:D';
    expect(findRefs(f).map((t) => f.slice(t.start, t.end))).toEqual([
      'B2:B10',
      'A1',
      "'My Sheet'!$C$3",
      'Sheet2!D:D',
    ]);
  });

  it('shifts relative references like Excel copy', () => {
    expect(shiftFormula('A1+$B$1+B$2+$C3', 2, 1)).toBe('B3+$B$1+C$2+$C5');
  });

  it('normalizes position-independent formulas', () => {
    expect(toR1C1('SUM(B2:B4)', 4, 1)).toBe(toR1C1('SUM(C2:C4)', 4, 2));
    expect(toR1C1('SUM(B2:B4)', 4, 1)).not.toBe(toR1C1('SUM(C2:C3)', 4, 2));
    expect(normalize('= sum( a1 , "x y" )')).toBe('SUM(A1,"x y")');
  });

  it('extracts functions, literals and ranges', () => {
    const f = 'IFERROR(SUM(B2:B10)*1.05+LOG10(A1)-"3",0)';
    expect(functions(f)).toEqual(['IFERROR', 'SUM', 'LOG10']);
    expect(literals(f)).toEqual([1.05, 0]);
    expect(ranges(f)).toEqual([
      { sheet: undefined, r0: 1, c0: 1, r1: 9, c1: 1 },
      { sheet: undefined, r0: 0, c0: 0, r1: 0, c1: 0 },
    ]);
  });

  it('remaps rows through a mapping', () => {
    expect(mapRefs('SUM(B2:B5)', (r) => ({ ...r, row: r.row >= 3 ? r.row + 1 : r.row }))).toBe('SUM(B2:B6)');
  });
});

describe('error wrappers', () => {
  it('unwraps an outermost IFERROR or IFNA', () => {
    expect(unwrapMask('IFERROR(B5/B1,0)')).toBe('B5/B1');
    expect(unwrapMask('=iferror(SUM(A1:A3)/2, "n/a")')).toBe('SUM(A1:A3)/2');
    expect(unwrapMask('IFERROR(A1,0)+1')).toBe('IFERROR(A1,0)+1');
    expect(unwrapMask('SUM(A1:A3)')).toBe('SUM(A1:A3)');
  });
});

describe('row alignment', () => {
  it('finds the longest common subsequence', () => {
    expect(lcsPairs(['a', 'b', 'c', 'd'], ['a', 'x', 'b', 'c', 'd'])).toEqual([
      [0, 0],
      [1, 2],
      [2, 3],
      [3, 4],
    ]);
  });

  it('reports inserted, deleted and modified rows and maps references', () => {
    const ops = alignRows(['h', 'a', 'b', 'c', 't1'], ['h', 'a', 'new', 'b', 'c', 't2']);
    expect(ops.map((o) => o.type)).toEqual(['same', 'same', 'inserted', 'same', 'same', 'modified']);
    const map = rowMapper(ops, 5);
    expect([0, 1, 2, 3, 4].map(map)).toEqual([0, 1, 3, 4, 5]);
    const del = rowMapper(alignRows(['a', 'b', 'c'], ['a', 'c']), 3);
    expect([0, 1, 2].map(del)).toEqual([0, 1, 1]);
  });
});

describe('workbook files', () => {
  it('round-trips values, formulas, formats, hidden rows, states and names', () => {
    const bytes = writeWorkbook(
      [
        {
          name: 'Data',
          cells: {
            A1: 'Label & <tag>',
            B1: 42,
            C1: true,
            D1: { f: 'B1*2', v: 84 },
            E1: { v: 0.25, fmt: 'percent' },
          },
          hiddenRows: [2],
        },
        { name: 'Secret sheet', state: 'hidden', cells: { A1: { f: "'Secret sheet'!A2", v: '#REF!' } } },
      ],
      { Rate: 'Data!$B$1' },
    );
    const wb = readWorkbook(bytes);
    const data = wb.sheets[0];
    expect(data.cells.get('A1')).toMatchObject({ value: 'Label & <tag>', type: 's' });
    expect(data.cells.get('B1')).toMatchObject({ value: 42, type: 'n' });
    expect(data.cells.get('C1')).toMatchObject({ value: true, type: 'b' });
    expect(data.cells.get('D1')).toMatchObject({ formula: 'B1*2', value: 84 });
    expect(data.cells.get('E1')?.format).toBe('0.0%');
    expect([...data.hiddenRows]).toEqual([2]);
    expect(wb.sheets[1]).toMatchObject({ name: 'Secret sheet', state: 'hidden' });
    expect(wb.sheets[1].cells.get('A1')).toMatchObject({ type: 'e', value: '#REF!' });
    expect(wb.names.get('Rate')).toBe('Data!$B$1');
  });

  it('expands shared formulas', () => {
    const sheet = `<?xml version="1.0"?><worksheet><sheetData>
      <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>
      <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>
      <row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row>
    </sheetData></worksheet>`;
    const bytes = zipSync({
      'xl/workbook.xml': strToU8(
        '<workbook><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
      ),
      'xl/worksheets/sheet1.xml': strToU8(sheet),
    });
    const s = readWorkbook(bytes).sheets[0];
    expect(['B1', 'B2', 'B3'].map((r) => s.cells.get(r)?.formula)).toEqual(['A1*2', 'A2*2', 'A3*2']);
  });

  it('explains unsupported files', () => {
    expect(() => readWorkbook(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]))).toThrow(/legacy \.xls/);
    expect(() => readWorkbook(strToU8('hello'))).toThrow(WorkbookError);
    expect(() => readWorkbook(zipSync({ 'a.txt': strToU8('x') }))).toThrow(/no workbook part/);
  });
});

describe('comparison', () => {
  const { original, edited } = demoWorkbooks();
  const result = compareWorkbooks(readWorkbook(original), readWorkbook(edited));
  const find = (kind: Kind, cell?: string, sheet?: string) =>
    result.changes.find(
      (c) => c.kind === kind && (!cell || c.cell === cell) && (!sheet || c.sheet === sheet),
    );

  it('finds every planted problem in the demo', () => {
    expect(find('range-shrunk', 'B7', 'Costs')?.notes[0]).toBe('B2:B6 became B2:B5: 1 cell(s) dropped');
    expect(find('range-shrunk', 'B7', 'Costs')).toMatchObject({ originalCell: 'B6', expected: 'SUM(B2:B6)' });
    expect(find('range-shrunk', 'D5', 'Revenue')?.expected).toBeUndefined();
    expect(find('formula-hardcoded', 'B3', 'Summary')).toBeDefined();
    expect(find('error-masked', 'B6', 'Summary')?.notes).toEqual([
      'Wrapped in IFERROR/IFNA: errors now show as a normal value',
    ]);
    expect(find('literal-changed', 'B7', 'Revenue')?.notes[0]).toContain('1.05 → 1.08');
    expect(find('range-shrunk', 'D5', 'Revenue')).toBeDefined();
    expect(find('input-changed', 'C4', 'Revenue')?.detail).toContain('54000 → 45000 (-16.7 %)');
    expect(find('row-hidden', 'A5', 'Costs')).toBeDefined();
    expect(find('sheet-hidden', undefined, 'Assumptions')).toBeDefined();
    expect(find('name-changed')?.detail).toBe('TaxRate: Assumptions!$B$1 → Assumptions!$B$2');
    expect(find('rows-inserted', 'A4', 'Costs')).toBeDefined();
    expect(result.verdict).toBe('risky');
  });

  it('accepts references that moved because of inserted rows', () => {
    // Summary!B2 changed from =Costs!B6 to =Costs!B7: that is Excel's own adjustment, not an edit.
    expect(
      result.changes.some((c) => c.sheet === 'Summary' && c.cell === 'B2' && c.severity !== 'info'),
    ).toBe(false);
    // Unchanged rows below the insertion are not reported.
    expect(result.changes.some((c) => c.sheet === 'Costs' && (c.cell === 'B5' || c.cell === 'B6'))).toBe(
      false,
    );
  });

  it('reports recalculated results as information only', () => {
    expect(find('output-changed', 'B1', 'Summary')?.severity).toBe('info');
  });

  it('is clean for identical workbooks', () => {
    const same = compareWorkbooks(readWorkbook(original), readWorkbook(original));
    expect(same.changes).toEqual([]);
    expect(same.verdict).toBe('clean');
  });

  it('flags new errors, formula-to-constant, added sheets and pattern breaks', () => {
    const a = writeWorkbook([
      {
        name: 'S',
        cells: {
          A1: 1,
          A2: 2,
          B1: { f: 'A1*2', v: 2 },
          B2: { f: 'A2*2', v: 4 },
          C1: { f: 'B1+1', v: 3 },
          D1: { f: 'C1+1', v: 4 },
          E1: { f: 'D1+1', v: 5 },
        },
      },
    ]);
    const b = writeWorkbook([
      {
        name: 'S',
        cells: {
          A1: 1,
          A2: 2,
          B1: { f: 'A1*2', v: '#REF!' },
          B2: 99,
          C1: { f: 'B1+1', v: 3 },
          D1: { f: 'C1+1', v: 4 },
          E1: { f: 'D1+1', v: 5 },
        },
      },
      { name: 'Extra', cells: { A1: 'x' } },
    ]);
    const r = compareWorkbooks(readWorkbook(a), readWorkbook(b));
    expect(r.changes.find((c) => c.cell === 'B1')?.kind).toBe('error-introduced');
    expect(r.changes.find((c) => c.cell === 'B2')?.kind).toBe('formula-to-value');
    expect(r.changes.some((c) => c.kind === 'sheet-added' && c.sheet === 'Extra')).toBe(true);
  });

  it('flags errors inside inserted rows', () => {
    const a = writeWorkbook([{ name: 'S', cells: { A1: 'a', B1: 1, A2: 'b', B2: 2 } }]);
    const b = writeWorkbook([
      { name: 'S', cells: { A1: 'a', B1: 1, A2: 'new', B2: { f: '#REF!*2', v: '#REF!' }, A3: 'b', B3: 2 } },
    ]);
    const r = compareWorkbooks(readWorkbook(a), readWorkbook(b));
    expect(r.changes.find((c) => c.kind === 'rows-inserted')?.cell).toBe('A2');
    expect(r.changes.find((c) => c.kind === 'error-introduced')?.cell).toBe('B2');
    expect(r.verdict).toBe('risky');
  });
});

describe('reports', () => {
  const { original, edited } = demoWorkbooks();
  const result = compareWorkbooks(readWorkbook(original), readWorkbook(edited));
  const files = { original: 'budget.xlsx', edited: 'budget-ai.xlsx' };

  it('writes JSON, CSV and Markdown', () => {
    const json = JSON.parse(jsonReport(result, files));
    expect(json.verdict).toBe('risky');
    expect(json.counts.critical).toBeGreaterThan(0);
    const csv = csvReport(result);
    expect(csv.split('\n')[0]).toBe('severity,kind,sheet,cell,original_cell,before,after,detail');
    expect(csv).toContain('formula-hardcoded');
    const md = markdownReport(result, files);
    expect(md).toContain('# Sheet Sentry review');
    expect(md).toContain('Formula replaced by its own result');
  });

  it('neutralizes spreadsheet formulas in CSV cells', () => {
    const csv = csvReport({ ...result, changes: [{ ...result.changes[0], detail: '=HYPERLINK("x")' }] });
    expect(csv).toContain("'=HYPERLINK");
  });
});
