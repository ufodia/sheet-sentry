import { writeWorkbook, type SheetSpec } from './core/xlsx.ts';

/**
 * A synthetic Q3 budget and the version an AI assistant returned after being
 * asked to "add a Marketing line to costs and tidy up the summary".
 *
 * The edit also, silently: left the cost total's SUM range short, hardcoded
 * the profit, wrapped the margin in IFERROR, changed a growth factor,
 * shortened a monthly total, mistyped an input, hid a row, hid the
 * assumptions sheet and pointed the TaxRate name at a different cell.
 * Cached values are what Excel would calculate for each version.
 */

const assumptions = (state?: 'hidden'): SheetSpec => ({
  name: 'Assumptions',
  state,
  cells: { A1: 'Tax rate', B1: { v: 0.2, fmt: 'percent' }, A2: 'Reserve', B2: { v: 0.1, fmt: 'percent' } },
});

function revenue(aug: number, d5: string, growth: number): SheetSpec {
  const rows = {
    Starter: [12000, 12600, 13230],
    Pro: [28000, 29400, 30870],
    Enterprise: [54000, aug, 61000],
  };
  const cells: SheetSpec['cells'] = { A1: 'Product', B1: 'Jul', C1: 'Aug', D1: 'Sep', E1: 'Q3' };
  const totals = [0, 0, 0];
  let q3 = 0;
  Object.entries(rows).forEach(([name, months], i) => {
    const r = i + 2;
    cells[`A${r}`] = name;
    months.forEach((v, m) => {
      cells[`${'BCD'[m]}${r}`] = v;
      totals[m] += v;
    });
    const sum = months.reduce((a, b) => a + b, 0);
    q3 += sum;
    cells[`E${r}`] = { f: `SUM(B${r}:D${r})`, v: sum };
  });
  const d5Value = d5 === 'SUM(D2:D4)' ? totals[2] : 13230 + 30870;
  cells.A5 = 'Total';
  cells.B5 = { f: 'SUM(B2:B4)', v: totals[0] };
  cells.C5 = { f: 'SUM(C2:C4)', v: totals[1] };
  cells.D5 = { f: d5, v: d5Value };
  cells.E5 = { f: 'SUM(E2:E4)', v: q3 };
  cells.A7 = 'Next quarter (growth)';
  cells.B7 = { f: `E5*${growth}`, v: Math.round(q3 * growth * 100) / 100 };
  return { name: 'Revenue', cells };
}

export function demoWorkbooks(): { original: Uint8Array; edited: Uint8Array } {
  // ——— Original ———
  const rev = 37830 + 88270 + 169000;
  const cost = 98000 + 14500 + 6200 + 9000;
  const profit = rev - cost;
  const tax = profit * 0.2;
  const original = writeWorkbook(
    [
      revenue(54000, 'SUM(D2:D4)', 1.05),
      {
        name: 'Costs',
        cells: {
          A1: 'Item',
          B1: 'Q3',
          A2: 'Salaries',
          B2: 98000,
          A3: 'Hosting',
          B3: 14500,
          A4: 'Software',
          B4: 6200,
          A5: 'Office',
          B5: 9000,
          A6: 'Total costs',
          B6: { f: 'SUM(B2:B5)', v: cost },
        },
      },
      {
        name: 'Summary',
        cells: {
          A1: 'Revenue',
          B1: { f: 'Revenue!E5', v: rev },
          A2: 'Costs',
          B2: { f: 'Costs!B6', v: cost },
          A3: 'Profit',
          B3: { f: 'B1-B2', v: profit },
          A4: 'Tax',
          B4: { f: 'B3*TaxRate', v: tax },
          A5: 'Net profit',
          B5: { f: 'B3-B4', v: profit - tax },
          A6: 'Net margin',
          B6: { f: 'B5/B1', v: (profit - tax) / rev, fmt: 'percent' },
          A7: 'Cost ratio',
          B7: { f: 'B2/B1', v: cost / rev, fmt: 'percent' },
        },
      },
      assumptions(),
    ],
    { TaxRate: 'Assumptions!$B$1' },
  );

  // ——— AI edit ———
  const rev2 = 37830 + 88270 + (54000 + 45000 + 61000);
  const cost2 = 98000 + 14500 + 12000 + 6200; // the SUM misses Office
  const hardcodedProfit = profit; // pasted value from the old file
  const tax2 = hardcodedProfit * 0.1; // TaxRate now points at the 10 % reserve
  const edited = writeWorkbook(
    [
      revenue(45000, 'SUM(D2:D3)', 1.08),
      {
        name: 'Costs',
        hiddenRows: [4],
        cells: {
          A1: 'Item',
          B1: 'Q3',
          A2: 'Salaries',
          B2: 98000,
          A3: 'Hosting',
          B3: 14500,
          A4: 'Marketing',
          B4: 12000,
          A5: 'Software',
          B5: 6200,
          A6: 'Office',
          B6: 9000,
          A7: 'Total costs',
          B7: { f: 'SUM(B2:B5)', v: cost2 },
        },
      },
      {
        name: 'Summary',
        cells: {
          A1: 'Revenue',
          B1: { f: 'Revenue!E5', v: rev2 },
          A2: 'Costs',
          B2: { f: 'Costs!B7', v: cost2 },
          A3: 'Profit',
          B3: hardcodedProfit,
          A4: 'Tax',
          B4: { f: 'B3*TaxRate', v: tax2 },
          A5: 'Net profit',
          B5: { f: 'B3-B4', v: hardcodedProfit - tax2 },
          A6: 'Net margin',
          B6: { f: 'IFERROR(B5/B1,0)', v: (hardcodedProfit - tax2) / rev2, fmt: 'percent' },
          A7: 'Cost ratio',
          B7: { f: 'B2/B1', v: cost2 / rev2, fmt: 'percent' },
        },
      },
      assumptions('hidden'),
    ],
    { TaxRate: 'Assumptions!$B$2' },
  );
  return { original, edited };
}
