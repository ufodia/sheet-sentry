import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { address, parseAddress, shiftFormula } from './refs.ts';

/** A cell as stored in the file: the formula (if any) and the cached value Excel last calculated. */
export interface Cell {
  /** Formula without the leading "=", with shared formulas expanded for this cell. */
  formula?: string;
  /** Cached or constant value; `undefined` when a formula has no cached result. */
  value?: string | number | boolean;
  /** `e` for error values such as #REF!, `s` text, `n` number, `b` boolean. */
  type: 'n' | 's' | 'b' | 'e';
  /** Number format code, for example `0.0%`. */
  format?: string;
}

export interface Sheet {
  name: string;
  state: 'visible' | 'hidden' | 'veryHidden';
  /** Cells keyed by A1 address. */
  cells: Map<string, Cell>;
  rows: number;
  cols: number;
  hiddenRows: Set<number>;
}

export interface Workbook {
  sheets: Sheet[];
  /** Defined names: name → formula text. */
  names: Map<string, string>;
  hasMacros: boolean;
  externalLinks: number;
}

export class WorkbookError extends Error {}

export const LIMITS = { files: 50 * 1024 * 1024, xml: 300 * 1024 * 1024, cells: 3_000_000 };

const ENTITY: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
export const decodeXml = (s: string) =>
  s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e: string) =>
    e[0] === '#'
      ? String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : Number(e.slice(1)))
      : (ENTITY[e] ?? m),
  );
const encodeXml = (s: string) =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function attrs(tag: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const m of tag.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = decodeXml(m[2]);
  return out;
}

const BUILTIN_FORMATS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

function resolvePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const p of target.split('/')) {
    if (p === '..') parts.pop();
    else if (p !== '.') parts.push(p);
  }
  return parts.join('/');
}

/** Read an .xlsx or .xlsm workbook. Only the parts needed for comparison are decompressed. */
export function readWorkbook(bytes: Uint8Array): Workbook {
  if (bytes.length > LIMITS.files) throw new WorkbookError('The file is larger than 50 MB.');
  if (!(bytes[0] === 0x50 && bytes[1] === 0x4b)) {
    throw new WorkbookError(
      bytes[0] === 0xd0 && bytes[1] === 0xcf
        ? 'This is a legacy .xls file or a password-protected workbook. Save an unprotected copy as .xlsx first.'
        : 'This is not an .xlsx workbook.',
    );
  }
  let total = 0;
  let files: Record<string, Uint8Array>;
  let names: string[] = [];
  try {
    files = unzipSync(bytes, {
      filter(file) {
        names.push(file.name);
        const wanted =
          /^xl\/(?:workbook\.xml|_rels\/workbook\.xml\.rels|sharedStrings\.xml|styles\.xml|worksheets\/[^/]+\.xml)$/.test(
            file.name,
          );
        if (!wanted) return false;
        total += file.originalSize;
        if (total > LIMITS.xml) throw new WorkbookError('The workbook expands to more than 300 MB of XML.');
        return true;
      },
    });
  } catch (error) {
    if (error instanceof WorkbookError) throw error;
    throw new WorkbookError('The workbook could not be unzipped.');
  }
  names = names.map((n) => n.replace(/\\/g, '/'));
  const text = (path: string) => (files[path] ? strFromU8(files[path]) : '');
  const workbookXml = text('xl/workbook.xml');
  if (!workbookXml) throw new WorkbookError('The file has no workbook part; it may not be a spreadsheet.');

  // Shared strings: plain or rich text runs.
  const shared: string[] = [];
  for (const m of text('xl/sharedStrings.xml').matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    const runs = [...m[1].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((t) => decodeXml(t[1]));
    shared.push(runs.join(''));
  }

  // Number formats by style index.
  const stylesXml = text('xl/styles.xml');
  const custom = new Map<number, string>();
  for (const m of stylesXml.matchAll(/<numFmt\s[^>]*>/g)) {
    const a = attrs(m[0]);
    custom.set(Number(a.numFmtId), a.formatCode);
  }
  const cellXfs = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? '';
  const styleFormats = [...cellXfs.matchAll(/<xf\s[^>]*?\/?>/g)].map((m) => {
    const id = Number(attrs(m[0]).numFmtId ?? 0);
    return custom.get(id) ?? BUILTIN_FORMATS[id] ?? 'General';
  });

  // Sheet list and relationships.
  const rels = new Map<string, string>();
  for (const m of text('xl/_rels/workbook.xml.rels').matchAll(/<Relationship\s[^>]*>/g)) {
    const a = attrs(m[0]);
    rels.set(a.Id, resolvePath('xl/workbook.xml', a.Target));
  }
  const definedNames = new Map<string, string>();
  for (const m of workbookXml.matchAll(/<definedName\s([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const a = attrs(`<x ${m[1]}>`);
    if (a.name?.startsWith('_xlnm.')) continue; // print areas and built-in names
    const scope = a.localSheetId !== undefined ? `[sheet ${Number(a.localSheetId) + 1}] ` : '';
    definedNames.set(scope + a.name, decodeXml(m[2]));
  }

  let cellCount = 0;
  const sheets: Sheet[] = [];
  for (const m of workbookXml.matchAll(/<sheet\s[^>]*>/g)) {
    const a = attrs(m[0]);
    const path = rels.get(a['r:id']);
    const xml = path ? text(path) : '';
    const sheet: Sheet = {
      name: a.name,
      state: a.state === 'hidden' ? 'hidden' : a.state === 'veryHidden' ? 'veryHidden' : 'visible',
      cells: new Map(),
      rows: 0,
      cols: 0,
      hiddenRows: new Set(),
    };
    if (!xml) {
      sheets.push(sheet); // chart sheets and missing parts are listed without cells
      continue;
    }
    const sharedFormulas = new Map<string, { formula: string; row: number; col: number }>();
    for (const row of xml.matchAll(/<row\s([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
      const ra = attrs(`<r ${row[1]}>`);
      if (ra.hidden === '1' || ra.hidden === 'true') sheet.hiddenRows.add(Number(ra.r) - 1);
      if (!row[2]) continue;
      for (const c of row[2].matchAll(/<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const ca = attrs(`<c ${c[1]}>`);
        if (!ca.r) continue;
        const { col, row: r } = parseAddress(ca.r);
        const body = c[2] ?? '';
        const fTag = /<f(\s[^>]*)?(?:\/>|>([\s\S]*?)<\/f>)/.exec(body);
        let formula: string | undefined;
        if (fTag) {
          const fa = attrs(`<f ${fTag[1] ?? ''}>`);
          const ftext = fTag[2] !== undefined ? decodeXml(fTag[2]) : '';
          if (fa.t === 'shared' && fa.si !== undefined) {
            if (ftext) {
              sharedFormulas.set(fa.si, { formula: ftext, row: r, col });
              formula = ftext;
            } else {
              const master = sharedFormulas.get(fa.si);
              if (master) formula = shiftFormula(master.formula, r - master.row, col - master.col);
            }
          } else formula = ftext || undefined;
        }
        const raw = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1];
        const inline = /<is>([\s\S]*?)<\/is>/.exec(body)?.[1];
        let value: Cell['value'];
        let type: Cell['type'] = 'n';
        if (ca.t === 's' && raw !== undefined) {
          value = shared[Number(raw)] ?? '';
          type = 's';
        } else if (ca.t === 'inlineStr' && inline !== undefined) {
          value = [...inline.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)]
            .map((t) => decodeXml(t[1]))
            .join('');
          type = 's';
        } else if (ca.t === 'str' && raw !== undefined) {
          value = decodeXml(raw);
          type = 's';
        } else if (ca.t === 'b' && raw !== undefined) {
          value = raw === '1';
          type = 'b';
        } else if (ca.t === 'e' && raw !== undefined) {
          value = decodeXml(raw);
          type = 'e';
        } else if (raw !== undefined) {
          value = Number(raw);
        }
        if (formula === undefined && value === undefined) continue; // styled but empty
        if (++cellCount > LIMITS.cells)
          throw new WorkbookError('The workbook has more than 3 million cells.');
        const format = ca.s !== undefined ? styleFormats[Number(ca.s)] : undefined;
        sheet.cells.set(address(col, r), {
          formula,
          value,
          type,
          format: format && format !== 'General' ? format : undefined,
        });
        sheet.rows = Math.max(sheet.rows, r + 1);
        sheet.cols = Math.max(sheet.cols, col + 1);
      }
    }
    sheets.push(sheet);
  }
  if (!sheets.length) throw new WorkbookError('The workbook has no sheets.');
  return {
    sheets,
    names: definedNames,
    hasMacros: names.some((n) => /^xl\/vbaProject\.bin$/i.test(n)),
    externalLinks: names.filter((n) => /^xl\/externalLinks\/externalLink\d+\.xml$/i.test(n)).length,
  };
}

// ——— Writing (used for the demo and tests) ———

export type CellSpec =
  | number
  | string
  | boolean
  | { f: string; v?: number | string; fmt?: 'percent' }
  | { v: number; fmt: 'percent' };

export interface SheetSpec {
  name: string;
  state?: 'hidden';
  cells: Record<string, CellSpec>;
  hiddenRows?: number[];
}

/** Write a minimal, valid .xlsx that Excel, LibreOffice and Google Sheets open. */
export function writeWorkbook(sheets: SheetSpec[], names: Record<string, string> = {}): Uint8Array {
  const strings: string[] = [];
  const stringIndex = new Map<string, number>();
  const sst = (s: string) => {
    if (!stringIndex.has(s)) {
      stringIndex.set(s, strings.length);
      strings.push(s);
    }
    return stringIndex.get(s)!;
  };
  const files: Record<string, Uint8Array> = {};
  sheets.forEach((sheet, index) => {
    const byRow = new Map<number, string[]>();
    const entries = Object.entries(sheet.cells).sort(([a], [b]) => {
      const pa = parseAddress(a);
      const pb = parseAddress(b);
      return pa.row - pb.row || pa.col - pb.col;
    });
    for (const [ref, spec] of entries) {
      const { row } = parseAddress(ref);
      let xml: string;
      if (typeof spec === 'number') xml = `<c r="${ref}"><v>${spec}</v></c>`;
      else if (typeof spec === 'boolean') xml = `<c r="${ref}" t="b"><v>${spec ? 1 : 0}</v></c>`;
      else if (typeof spec === 'string') xml = `<c r="${ref}" t="s"><v>${sst(spec)}</v></c>`;
      else if ('f' in spec) {
        const style = spec.fmt === 'percent' ? ' s="1"' : '';
        const t = typeof spec.v === 'string' ? (spec.v.startsWith('#') ? ' t="e"' : ' t="str"') : '';
        xml = `<c r="${ref}"${style}${t}><f>${encodeXml(spec.f)}</f>${spec.v !== undefined ? `<v>${encodeXml(String(spec.v))}</v>` : ''}</c>`;
      } else xml = `<c r="${ref}" s="1"><v>${spec.v}</v></c>`;
      if (!byRow.has(row)) byRow.set(row, []);
      byRow.get(row)!.push(xml);
    }
    const hidden = new Set(sheet.hiddenRows ?? []);
    const rows = [...new Set([...byRow.keys(), ...hidden])].sort((a, b) => a - b);
    const sheetData = rows
      .map(
        (r) =>
          `<row r="${r + 1}"${hidden.has(r) ? ' hidden="1"' : ''}>${(byRow.get(r) ?? []).join('')}</row>`,
      )
      .join('');
    files[`xl/worksheets/sheet${index + 1}.xml`] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetData}</sheetData></worksheet>`,
    );
  });
  const ns =
    'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const definedNames = Object.entries(names)
    .map(([n, f]) => `<definedName name="${encodeXml(n)}">${encodeXml(f)}</definedName>`)
    .join('');
  files['xl/workbook.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook ${ns}><sheets>${sheets
      .map(
        (s, i) =>
          `<sheet name="${encodeXml(s.name)}" sheetId="${i + 1}"${s.state ? ` state="${s.state}"` : ''} r:id="rId${i + 1}"/>`,
      )
      .join('')}</sheets>${definedNames ? `<definedNames>${definedNames}</definedNames>` : ''}</workbook>`,
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets
      .map(
        (_, i) =>
          `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
      )
      .join(
        '',
      )}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId${sheets.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  );
  files['xl/sharedStrings.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">${strings
      .map((s) => `<si><t xml:space="preserve">${encodeXml(s)}</t></si>`)
      .join('')}</sst>`,
  );
  files['xl/styles.xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>`,
  );
  files['[Content_Types].xml'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets
      .map(
        (_, i) =>
          `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
      )
      .join(
        '',
      )}<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
  );
  files['_rels/.rels'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  );
  return zipSync(files);
}
