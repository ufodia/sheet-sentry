import { strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { compareWorkbooks } from '../src/core/compare.ts';
import { csvReport, jsonReport, markdownReport } from '../src/core/report.ts';
import { normalize, parseAddress, shiftFormula } from '../src/core/refs.ts';
import { readWorkbook, writeWorkbook, WorkbookError } from '../src/core/xlsx.ts';

function fixture(
  xml: string,
  workbook = '<workbook><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>',
) {
  return zipSync({
    'xl/workbook.xml': strToU8(workbook),
    'xl/_rels/workbook.xml.rels': strToU8(
      '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    ),
    'xl/worksheets/sheet1.xml': strToU8(xml),
  });
}

const compare = (
  a: Record<string, Parameters<typeof writeWorkbook>[0][number]['cells'][string]>,
  b: typeof a,
) =>
  compareWorkbooks(
    readWorkbook(writeWorkbook([{ name: 'S', cells: a }])),
    readWorkbook(writeWorkbook([{ name: 'S', cells: b }])),
  );

describe('workbook parsing regressions', () => {
  it('distinguishes uncached formula results from false, errors and empty strings', () => {
    const sheet = readWorkbook(
      fixture(`<worksheet><sheetData><row r="1">
      <c r="A1" t="b"><f>1=1</f><v></v></c>
      <c r="B1" t="e"><f>1/0</f><v> </v></c>
      <c r="C1" t="s"><f>"test"</f><v></v></c>
      <c r="D1" t="b"><f>1=1</f><v>1</v></c>
      <c r="E1" t="b"><f>1=0</f><v>0</v></c>
      <c r="F1" t="str"><f>""</f><v></v></c>
      </row></sheetData></worksheet>`),
    ).sheets[0];
    for (const ref of ['A1', 'B1', 'C1']) expect(sheet.cells.get(ref)?.value).toBeUndefined();
    expect(sheet.cells.get('B1')?.type).not.toBe('e');
    expect(sheet.cells.get('D1')?.value).toBe(true);
    expect(sheet.cells.get('E1')?.value).toBe(false);
    expect(sheet.cells.get('F1')?.value).toBe('');
    for (const cell of [
      '<c r="A1" t="b"><v>2</v></c>',
      '<c r="A1" t="b"><v></v></c>',
      '<c r="A1" t="e"><v></v></c>',
    ])
      expect(() =>
        readWorkbook(fixture(`<worksheet><sheetData><row r="1">${cell}</row></sheetData></worksheet>`)),
      ).toThrow(WorkbookError);
  });
  it('accepts legal single-quoted XML attributes and normalized cell addresses', () => {
    const xml =
      "<worksheet><sheetData><row r='1'><c r='a1' t='inlineStr'><is><t>&lt;safe&gt;</t></is></c></row></sheetData></worksheet>";
    const workbook = "<workbook><sheets><sheet name='S' r:id='rId1'/></sheets></workbook>";
    expect(readWorkbook(fixture(xml, workbook)).sheets[0].cells.get('A1')?.value).toBe('<safe>');
  });

  it('resolves shared formulas whose master appears later and has no cached result', () => {
    const xml =
      '<worksheet><sheetData><row r="1"><c r="B1"><f t="shared" si="0"/></c></row><row r="2"><c r="B2"><f t="shared" si="0">A2*2</f><v>4</v></c></row></sheetData></worksheet>';
    const sheet = readWorkbook(fixture(xml)).sheets[0];
    expect(sheet.cells.get('B1')).toMatchObject({ formula: 'A1*2', value: undefined });
  });

  it('rejects missing shared formula definitions rather than treating them as constants', () => {
    expect(() =>
      readWorkbook(
        fixture(
          '<worksheet><sheetData><row r="1"><c r="A1"><f t="shared" si="3"/><v>42</v></c></row></sheetData></worksheet>',
        ),
      ),
    ).toThrow(/has no definition/);
  });

  it('rejects missing worksheet parts instead of reporting an empty sheet', () => {
    const files = unzipSync(writeWorkbook([{ name: 'S', cells: { A1: 42 } }]));
    delete files['xl/worksheets/sheet1.xml'];
    expect(() => readWorkbook(zipSync(files))).toThrow(/missing or unsupported/);
  });

  it('rejects entity declarations without resolving them', () => {
    expect(() =>
      readWorkbook(fixture('<!DOCTYPE worksheet [<!ENTITY secret SYSTEM "file:///private">]><worksheet/>')),
    ).toThrow(/entity declarations/);
  });

  it('rejects invalid XML code points, numeric values and shared-string indices', () => {
    for (const body of [
      '<c r="A1" t="inlineStr"><is><t>&#99999999;</t></is></c>',
      '<c r="A1"><v>NaN</v></c>',
      '<c r="A1"><v>  </v></c>',
      '<c r="A1" t="s"><v>99</v></c>',
    ])
      expect(() =>
        readWorkbook(fixture(`<worksheet><sheetData><row r="1">${body}</row></sheetData></worksheet>`)),
      ).toThrow(WorkbookError);
  });

  it('bounds addresses and aggregate sparse row spans before comparison', () => {
    for (const ref of ['A0', 'XFE1', 'A1048577', 'A999999999999999999999'])
      expect(() => parseAddress(ref)).toThrow();
    const large = writeWorkbook([
      { name: 'A', cells: { A800000: 1 } },
      { name: 'B', cells: { A800000: 1 } },
    ]);
    expect(() => readWorkbook(large)).toThrow(/1.5 million/);
  });

  it('preserves hidden empty rows in the worksheet span', () => {
    const sheet = readWorkbook(
      fixture('<worksheet><sheetData><row r="10" hidden="1"/></sheetData></worksheet>'),
    ).sheets[0];
    expect(sheet.rows).toBe(10);
    expect(sheet.hiddenRows.has(9)).toBe(true);
  });

  it('refuses duplicate cells instead of silently overwriting data', () => {
    const xml =
      '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c></row></sheetData></worksheet>';
    expect(() => readWorkbook(fixture(xml))).toThrow(/Duplicate cell/);
  });
});

describe('comparison regressions', () => {
  it('preserves spaces in sheet names and reference intersections when normalizing', () => {
    expect(normalize("'My Sheet'!A1")).not.toBe(normalize("'MySheet'!A1"));
    expect(normalize('SUM(A1 B1)')).not.toBe(normalize('SUM(A1B1)'));
  });
  it('remaps case-insensitive cross-sheet references through inferred insertions', () => {
    const a = readWorkbook(
      writeWorkbook([
        { name: 'Data', cells: { A1: 'one', A2: 'two' } },
        { name: 'Summary', cells: { A1: { f: 'data!A2', v: 'two' } } },
      ]),
    );
    const b = readWorkbook(
      writeWorkbook([
        { name: 'Data', cells: { A1: 'one', A2: 'new', A3: 'two' } },
        { name: 'Summary', cells: { A1: { f: 'data!A3', v: 'two' } } },
      ]),
    );
    expect(compareWorkbooks(a, b).changes.some((c) => c.sheet === 'Summary')).toBe(false);
  });
  it('compares a wide sparse sheet without scanning every empty cell', () => {
    const result = compare({ A1: 'first', XFD10000: 1 }, { A1: 'first', XFD10000: 2 });
    expect(result.changes.find((c) => c.cell === 'XFD10000')?.kind).toBe('input-changed');
    expect(result.compared.cells).toBe(2);
  });

  it('does not hide value type changes behind identical display text', () => {
    expect(
      compare({ A1: 1, B1: true }, { A1: '1', B1: 'true' }).changes.filter((c) => c.kind === 'text-changed'),
    ).toHaveLength(2);
  });

  it('does not let separator text impersonate a different row signature', () => {
    const result = compare({ A1: 'a|1=Vb' }, { A1: 'a', B1: 'b' });
    expect(result.changes.some((c) => c.kind === 'text-changed')).toBe(true);
    expect(result.changes.some((c) => c.kind === 'cell-added')).toBe(true);
  });

  it('flags errors added in new cells or sheets as critical', () => {
    const a = readWorkbook(writeWorkbook([{ name: 'S', cells: { A1: 1 } }]));
    const b = readWorkbook(
      writeWorkbook([
        { name: 'S', cells: { A1: 1, B1: { f: '1/0', v: '#DIV/0!' } } },
        { name: 'New', cells: { A1: { f: '#REF!', v: '#REF!' } } },
      ]),
    );
    const result = compareWorkbooks(a, b);
    expect(result.changes.filter((c) => c.kind === 'error-introduced')).toHaveLength(2);
    expect(result.verdict).toBe('risky');
  });

  it('does not call a literal #REF! string a formula error', () => {
    const result = compare({ A1: 'label' }, { A1: 'label', B1: { f: '"#REF!"', v: '#REF! text' } });
    expect(result.changes.some((c) => c.kind === 'error-introduced')).toBe(false);
  });

  it('requires reference review after deleted rows rather than accepting the next cell as equivalent', () => {
    const a = readWorkbook(
      writeWorkbook([
        { name: 'Data', cells: { A1: 'first', A2: 'deleted', A3: 'third' } },
        { name: 'Summary', cells: { A1: { f: 'Data!A2', v: 'deleted' } } },
      ]),
    );
    const b = readWorkbook(
      writeWorkbook([
        { name: 'Data', cells: { A1: 'first', A2: 'third' } },
        { name: 'Summary', cells: { A1: { f: 'Data!A2', v: 'third' } } },
      ]),
    );
    expect(compareWorkbooks(a, b).changes.find((c) => c.sheet === 'Summary')?.kind).toBe(
      'references-changed',
    );
  });

  it('preserves structured labels and emits #REF! for shifted references outside worksheet bounds', () => {
    expect(shiftFormula('SUM(Table1[A1])+A1', 1, 0)).toBe('SUM(Table1[A1])+A2');
    expect(shiftFormula('A1', -1, 0)).toBe('#REF!');
    expect(shiftFormula('XFD1048576', 1, 0)).toBe('#REF!');
    expect(shiftFormula('[Book.xlsx]Sheet1!A1+A1', 1, 0)).toBe('[Book.xlsx]Sheet1!A1+A2');
  });
});

describe('safe report serialization', () => {
  const result = compare({ A1: 1 }, { A1: 2 });
  it('quotes carriage returns and prefixes formula-like fields behind whitespace', () => {
    const csv = csvReport({
      ...result,
      changes: [{ ...result.changes[0], sheet: ' =SUM(A1)', detail: 'plain\rnext' }],
    });
    expect(csv).toContain("' =SUM(A1)");
    expect(csv).toContain('"plain\rnext"');
  });
  it('escapes workbook text and filenames in Markdown and preserves it in JSON', () => {
    const payload = '<img src=x onerror=alert(1)>\n| injected |';
    const changed = { ...result, changes: [{ ...result.changes[0], sheet: payload, notes: [payload] }] };
    const md = markdownReport(changed, { original: '**input**.xlsx', edited: payload });
    expect(md).not.toContain('<img');
    expect(md).toContain('&lt;img');
    expect(md).not.toContain('\n| injected |');
    expect(md).toContain('\\*\\*input\\*\\*');
    expect(JSON.parse(jsonReport(changed, { original: 'a', edited: 'b' })).changes[0].sheet).toBe(payload);
  });
});
