import { TITLE, type CellState, type Comparison } from './compare.ts';

export const TOOL = { name: 'Sheet Sentry', version: '1.0.0', url: 'https://github.com/ufodia/sheet-sentry' };

const show = (s?: CellState) =>
  !s ? '' : s.formula !== undefined ? `=${s.formula}` : s.value === undefined ? '' : String(s.value);

export function jsonReport(
  result: Comparison,
  files: { original: string; edited: string },
  now = new Date(),
): string {
  return (
    JSON.stringify(
      {
        tool: TOOL,
        generatedAt: now.toISOString(),
        files,
        verdict: result.verdict,
        counts: result.counts,
        compared: result.compared,
        sheets: result.sheets.map((s) => ({
          name: s.name,
          status: s.status,
          rowsInserted: s.inserted,
          rowsDeleted: s.deleted,
        })),
        changes: result.changes.map((c) => ({
          severity: c.severity,
          kind: c.kind,
          title: TITLE[c.kind],
          sheet: c.sheet,
          cell: c.cell ?? null,
          originalCell: c.originalCell ?? null,
          before: c.before ?? null,
          after: c.after ?? null,
          expectedFormula: c.expected ? `=${c.expected}` : null,
          detail: c.detail,
          notes: c.notes,
        })),
      },
      null,
      2,
    ) + '\n'
  );
}

/** CSV cells starting with = + - @ are prefixed with ' so spreadsheet apps do not execute them. */
const csvCell = (v: string) => {
  const safe = /^[=+\-@\t\r]/.test(v) ? `'${v}` : v;
  return /[",\n]/.test(safe) ? `"${safe.replace(/"/g, '""')}"` : safe;
};

export function csvReport(result: Comparison): string {
  const rows = [['severity', 'kind', 'sheet', 'cell', 'original_cell', 'before', 'after', 'detail']];
  for (const c of result.changes)
    rows.push([
      c.severity,
      c.kind,
      c.sheet,
      c.cell ?? '',
      c.originalCell ?? '',
      show(c.before),
      show(c.after),
      c.detail,
    ]);
  return rows.map((r) => r.map(csvCell).join(',')).join('\n') + '\n';
}

const VERDICT = {
  risky: 'Risky — critical changes need review before this workbook is used',
  review: 'Review — formulas or structure changed in ways that need a second look',
  minor: 'Minor — values or presentation changed; no formula risks found',
  clean: 'Clean — no differences beyond recalculated results',
};

export function markdownReport(
  result: Comparison,
  files: { original: string; edited: string },
  now = new Date(),
): string {
  const esc = (s: string) => s.replace(/\|/g, '\\|');
  const lines = [
    '# Sheet Sentry review',
    '',
    `**${files.original}** → **${files.edited}** · ${now.toISOString().slice(0, 10)}`,
    '',
    `**Verdict:** ${VERDICT[result.verdict]}`,
    '',
    `| Critical | High | Medium | Low | Info |`,
    `| ---: | ---: | ---: | ---: | ---: |`,
    `| ${result.counts.critical} | ${result.counts.high} | ${result.counts.medium} | ${result.counts.low} | ${result.counts.info} |`,
    '',
    `${result.compared.cells} cells and ${result.compared.formulas} formulas compared.`,
    '',
    '## Changes to review',
    '',
    '| Severity | Where | What | Detail |',
    '| --- | --- | --- | --- |',
  ];
  for (const c of result.changes.filter((x) => x.severity !== 'info'))
    lines.push(
      `| ${c.severity} | ${esc(c.sheet)}${c.cell ? '!' + c.cell : ''} | ${TITLE[c.kind]} | ${esc(c.notes.join('; ') || c.detail)} |`,
    );
  const info = result.changes.filter((x) => x.severity === 'info').length;
  if (info)
    lines.push(
      '',
      `${info} recalculated result(s) and inserted-row notes are listed in the JSON and CSV exports.`,
    );
  lines.push('');
  return lines.join('\n');
}
