import './style.css';
import {
  compareWorkbooks,
  TITLE,
  type Change,
  type Comparison,
  type Severity,
  type SheetResult,
} from './core/compare.ts';
import { address, indexToCol, parseAddress } from './core/refs.ts';
import { csvReport, jsonReport, markdownReport } from './core/report.ts';
import { LIMITS, readWorkbook, type Cell, type Workbook } from './core/xlsx.ts';
import { demoWorkbooks } from './demo.ts';
import { byId, download, h } from './dom.ts';

const WINDOW_ROWS = 120;
const MAX_COLS = 40;
const SEVERITIES: Severity[] = ['critical', 'high', 'medium', 'low', 'info'];
const LABEL: Record<Severity, string> = {
  critical: 'Critical',
  high: 'High',
  medium: 'Medium',
  low: 'Low',
  info: 'Info',
};

const VERDICT = {
  risky: {
    title: 'Risky edit',
    text: 'Critical changes: formulas were hardcoded, ranges shrank or errors appeared. Do not use this workbook before reviewing them.',
  },
  review: {
    title: 'Needs review',
    text: 'Formulas, names or structure changed in ways that deserve a second look.',
  },
  minor: { title: 'Minor changes', text: 'Values or presentation changed. No formula risks were found.' },
  clean: {
    title: 'No meaningful changes',
    text: 'No review findings within the supported scope. Formulas are not recalculated; this is not a correctness guarantee.',
  },
} as const;

interface Loaded {
  name: string;
  workbook: Workbook;
}

let original: Loaded | null = null;
let edited: Loaded | null = null;
let result: Comparison | null = null;
let sheetIndex = 0;
let selected: Change | null = null;
let top = 0;
let showInfo = false;
let filter: Severity | 'all' = 'all';
const requests = { original: 0, edited: 0 };

const intro = byId('intro');
const workspace = byId('workspace');
const introError = byId('intro-error');

function fail(message: string) {
  intro.hidden = false;
  workspace.hidden = true;
  introError.textContent = message;
  introError.hidden = false;
}

function clearSlot(which: 'original' | 'edited') {
  if (which === 'original') original = null;
  else edited = null;
  result = null;
  selected = null;
  byId(`slot-${which}`).classList.remove('filled');
  byId(`name-${which}`).textContent = '.xlsx or .xlsm — or drop it here';
  intro.hidden = false;
  workspace.hidden = true;
  workspace.replaceChildren();
}

function load(which: 'original' | 'edited', name: string, bytes: Uint8Array) {
  introError.hidden = true;
  try {
    const loaded = { name, workbook: readWorkbook(bytes) };
    if (which === 'original') original = loaded;
    else edited = loaded;
    const slot = byId(`slot-${which}`);
    slot.classList.add('filled');
    const sheets = loaded.workbook.sheets;
    const cells = sheets.reduce((n, s) => n + s.cells.size, 0);
    byId(`name-${which}`).textContent =
      `${name} · ${sheets.length} sheet${sheets.length === 1 ? '' : 's'} · ${cells.toLocaleString('en')} cells`;
    if (original && edited) run();
  } catch (error) {
    clearSlot(which);
    fail(`${name}: ${(error as Error).message}`);
  }
}

async function loadFile(which: 'original' | 'edited', file: File) {
  const request = ++requests[which];
  clearSlot(which);
  introError.hidden = true;
  try {
    if (file.size > LIMITS.files) throw new Error('The file is larger than 50 MB.');
    const bytes = new Uint8Array(await file.arrayBuffer());
    if (request === requests[which]) load(which, file.name, bytes);
  } catch (error) {
    if (request === requests[which]) fail(`${file.name}: ${(error as Error).message}`);
  }
}

function run() {
  if (!original || !edited) return;
  result = compareWorkbooks(original.workbook, edited.workbook);
  // Start on the sheet with the most severe change.
  const first = result.changes.find(
    (c) => c.cell && result!.sheets.some((s) => s.name === c.sheet && s.sheet),
  );
  sheetIndex = Math.max(
    0,
    result.sheets.findIndex((s) => s.name === first?.sheet),
  );
  selected = first ?? null;
  top = selected?.cell ? Math.max(0, parseAddress(selected.cell).row - 10) : 0;
  intro.hidden = true;
  workspace.hidden = false;
  render();
  window.scrollTo({ top: 0 });
}

for (const which of ['original', 'edited'] as const) {
  const input = byId<HTMLInputElement>(`file-${which}`);
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    input.value = '';
    if (file) await loadFile(which, file);
  });
  const slot = byId(`slot-${which}`);
  slot.addEventListener('dragover', (e) => {
    e.preventDefault();
    slot.classList.add('over');
  });
  slot.addEventListener('dragleave', () => slot.classList.remove('over'));
  slot.addEventListener('drop', async (e) => {
    e.preventDefault();
    e.stopPropagation();
    slot.classList.remove('over');
    const file = e.dataTransfer?.files[0];
    if (file) await loadFile(which, file);
  });
}
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', async (e) => {
  e.preventDefault();
  const files = [...(e.dataTransfer?.files ?? [])];
  if (files.length >= 2) {
    await Promise.all([loadFile('original', files[0]), loadFile('edited', files[1])]);
  } else if (files.length === 1) await loadFile(original ? 'edited' : 'original', files[0]);
});

byId('demo-button').addEventListener('click', () => {
  reset();
  const demo = demoWorkbooks();
  load('original', 'q3-budget.xlsx', demo.original);
  load('edited', 'q3-budget-ai-edit.xlsx', demo.edited);
});
byId('demo-download').addEventListener('click', () => {
  const demo = demoWorkbooks();
  saveBytes('q3-budget.xlsx', demo.original);
  setTimeout(() => saveBytes('q3-budget-ai-edit.xlsx', demo.edited), 300);
});

function saveBytes(name: string, bytes: Uint8Array) {
  const url = URL.createObjectURL(
    new Blob([bytes as Uint8Array<ArrayBuffer>], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    }),
  );
  const link = h('a', { href: url, download: name });
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ——— Rendering ———

function formatValue(cell?: Cell | { value?: Cell['value']; format?: string }): string {
  if (!cell || cell.value === undefined) return '';
  const v = cell.value;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  const format = cell.format ?? '';
  if (format.includes('%')) {
    const decimals = Math.min(10, /0\.(0+)%/.exec(format)?.[1].length ?? 0);
    return `${(v * 100).toFixed(decimals)}%`;
  }
  return Math.abs(v) >= 1000 || Number.isInteger(v)
    ? v.toLocaleString('en', { maximumFractionDigits: 2 })
    : String(+v.toPrecision(10));
}

function render() {
  if (!result || !original || !edited) return;
  const r = result;
  const verdict = VERDICT[r.verdict];
  const head = h(
    'section',
    { class: `verdict-bar v-${r.verdict}` },
    h(
      'div',
      {},
      h('p', { class: 'eyebrow' }, `${original.name} → ${edited.name}`),
      h('h1', {}, verdict.title),
      h('p', { class: 'muted' }, verdict.text),
    ),
    h(
      'div',
      { class: 'counts' },
      ...SEVERITIES.map((s) =>
        h(
          'button',
          {
            type: 'button',
            class: `count sev-${s}`,
            'aria-pressed': String(filter === s),
            onclick: () => {
              filter = filter === s ? 'all' : s;
              if (s === 'info') showInfo = true;
              render();
            },
          },
          h('strong', {}, String(r.counts[s])),
          h('span', {}, LABEL[s]),
        ),
      ),
    ),
    h(
      'div',
      { class: 'head-actions' },
      h('button', { class: 'button ghost small', type: 'button', onclick: () => reset() }, 'New comparison'),
    ),
  );

  const tabs = h('div', { class: 'sheet-tabs', role: 'tablist', 'aria-label': 'Sheets' });
  r.sheets.forEach((s, i) => {
    const worst = SEVERITIES.find((sev) =>
      r.changes.some((c) => c.sheet === s.name && c.severity === sev && sev !== 'info'),
    );
    tabs.append(
      h(
        'button',
        {
          type: 'button',
          role: 'tab',
          class: `sheet-tab ${worst ? `has-${worst}` : ''}`,
          'aria-selected': String(i === sheetIndex),
          onclick: () => {
            sheetIndex = i;
            top = 0;
            selected = null;
            render();
          },
        },
        s.name,
        s.status !== 'matched'
          ? h('span', { class: 'tag' }, s.status)
          : s.sheet && s.sheet.state !== 'visible'
            ? h('span', { class: 'tag' }, 'hidden')
            : null,
      ),
    );
  });

  workspace.replaceChildren(
    head,
    h(
      'div',
      { class: 'ws-grid' },
      h('div', { class: 'viewer' }, tabs, grid(r.sheets[sheetIndex]), inspector()),
      h('aside', { class: 'panel' }, findingsPanel(r), exportsPanel(r)),
    ),
  );
  document.querySelector('.cell.selected')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

function grid(sheetResult: SheetResult): HTMLElement {
  const sheet = sheetResult.sheet;
  if (!sheet)
    return h(
      'div',
      { class: 'grid-empty' },
      sheetResult.status === 'removed'
        ? 'This sheet was removed in the edited workbook.'
        : 'This sheet has no cells.',
    );
  const cols = Math.min(Math.max(sheet.cols, 1), MAX_COLS);
  const end = Math.min(sheet.rows, top + WINDOW_ROWS);
  const table = h('table', { class: 'grid', 'aria-label': `${sheet.name} cells` });
  const headRow = h('tr', {}, h('th', { class: 'corner' }));
  for (let c = 0; c < cols; c++) headRow.append(h('th', { scope: 'col' }, indexToCol(c)));
  table.append(h('thead', {}, headRow));
  const body = h('tbody');
  for (let r = top; r < end; r++) {
    const tr = h(
      'tr',
      { class: sheet.hiddenRows.has(r) ? 'hidden-row' : '' },
      h('th', { scope: 'row' }, String(r + 1)),
    );
    for (let c = 0; c < cols; c++) {
      const ref = address(c, r);
      const cell = sheet.cells.get(ref);
      const mark = sheetResult.marks.get(ref);
      const isSelected = selected?.cell === ref && selected.sheet === sheet.name;
      const td = h(
        'td',
        {
          class: `cell ${cell?.formula !== undefined ? 'formula' : ''} ${typeof cell?.value === 'number' ? 'num' : ''} ${mark && (mark.severity !== 'info' || showInfo) ? `mark sev-${mark.severity}` : ''} ${isSelected ? 'selected' : ''}`,
          title: cell?.formula !== undefined ? `=${cell.formula}` : undefined,
          tabindex: mark ? 0 : undefined,
          onclick: mark
            ? () => {
                selected = mark;
                render();
              }
            : undefined,
        },
        formatValue(cell),
      );
      tr.append(td);
    }
    body.append(tr);
  }
  table.append(body);
  const nav =
    sheet.rows > WINDOW_ROWS
      ? h(
          'div',
          { class: 'grid-nav' },
          h(
            'button',
            {
              type: 'button',
              class: 'link',
              disabled: top === 0,
              onclick: () => {
                top = Math.max(0, top - WINDOW_ROWS);
                render();
              },
            },
            '← Previous rows',
          ),
          h('span', { class: 'muted' }, `Rows ${top + 1}–${end} of ${sheet.rows}`),
          h(
            'button',
            {
              type: 'button',
              class: 'link',
              disabled: end >= sheet.rows,
              onclick: () => {
                top = Math.min(sheet.rows - 1, top + WINDOW_ROWS);
                render();
              },
            },
            'Next rows →',
          ),
        )
      : null;
  const note =
    sheet.cols > MAX_COLS
      ? h(
          'p',
          { class: 'fine' },
          `Showing columns A–${indexToCol(MAX_COLS - 1)} of ${sheet.cols}. Every column is compared; the list on the right covers all of them.`,
        )
      : null;
  return h('div', { class: 'grid-wrap' }, h('div', { class: 'grid-scroll' }, table), nav, note);
}

function stateBlock(label: string, s?: { formula?: string; value?: Cell['value']; format?: string }) {
  return h(
    'div',
    { class: 'state' },
    h('span', { class: 'state-label' }, label),
    s ? h('code', {}, s.formula !== undefined ? `=${s.formula}` : '(constant)') : h('code', {}, '(empty)'),
    s ? h('strong', {}, formatValue(s) || '—') : null,
  );
}

function inspector(): HTMLElement {
  if (!selected)
    return h(
      'section',
      { class: 'inspector empty' },
      h(
        'p',
        { class: 'muted' },
        'Select a highlighted cell or a change on the right to see the before and after.',
      ),
    );
  const c = selected;
  return h(
    'section',
    { class: `inspector sev-${c.severity}` },
    h(
      'header',
      {},
      h('span', { class: `pill sev-${c.severity}` }, LABEL[c.severity]),
      h('h2', {}, TITLE[c.kind]),
      h(
        'span',
        { class: 'where' },
        `${c.sheet}${c.cell ? '!' + c.cell : ''}${c.originalCell && c.originalCell !== c.cell ? ` (was ${c.originalCell})` : ''}`,
      ),
    ),
    c.before || c.after
      ? h(
          'div',
          { class: 'states' },
          stateBlock(
            c.originalCell && c.originalCell !== c.cell ? `Before · ${c.originalCell}` : 'Before',
            c.before,
          ),
          h('span', { class: 'arrow', 'aria-hidden': 'true' }, '→'),
          stateBlock('After', c.after),
        )
      : null,
    c.expected
      ? h(
          'p',
          { class: 'expected' },
          h('span', {}, 'Expected after the row changes'),
          h('code', {}, `=${c.expected}`),
          h(
            'span',
            { class: 'muted' },
            'This formula follows the inferred row insertions. Check the row alignment before accepting the finding.',
          ),
        )
      : null,
    h('ul', { class: 'notes' }, ...c.notes.map((n) => h('li', {}, n))),
  );
}

function findingsPanel(r: Comparison): HTMLElement {
  const list = r.changes.filter((c) =>
    filter === 'all' ? c.severity !== 'info' || showInfo : c.severity === filter,
  );
  const items = list.slice(0, 400).map((c) =>
    h(
      'li',
      {},
      h(
        'button',
        {
          type: 'button',
          class: `change sev-${c.severity} ${selected?.id === c.id ? 'active' : ''}`,
          onclick: () => {
            selected = c;
            const index = r.sheets.findIndex((s) => s.name === c.sheet);
            if (index >= 0) sheetIndex = index;
            if (c.cell) top = Math.max(0, parseAddress(c.cell).row - 10);
            render();
          },
        },
        h('span', { class: 'dot', 'aria-hidden': 'true' }),
        h(
          'span',
          { class: 'change-body' },
          h('strong', {}, TITLE[c.kind]),
          h('span', { class: 'where' }, `${c.sheet}${c.cell ? '!' + c.cell : ''}`),
          h('span', { class: 'detail' }, c.notes[0] ?? c.detail),
        ),
      ),
    ),
  );
  const toggle = h('input', { type: 'checkbox', checked: showInfo });
  toggle.addEventListener('change', () => {
    showInfo = toggle.checked;
    render();
  });
  return h(
    'section',
    { class: 'findings' },
    h(
      'div',
      { class: 'findings-head' },
      h('h3', {}, filter === 'all' ? 'Changes to review' : `${LABEL[filter]} changes`),
      h('label', { class: 'check' }, toggle, 'Show recalculated results'),
    ),
    list.length
      ? h('ol', { class: 'changes' }, ...items)
      : h('p', { class: 'muted' }, 'No changes in this view.'),
    list.length > 400
      ? h(
          'p',
          { class: 'fine' },
          `Showing the first 400 of ${list.length.toLocaleString('en')} matching findings. Export CSV or JSON for the complete list.`,
        )
      : null,
    h(
      'p',
      { class: 'fine' },
      `${r.compared.cells.toLocaleString('en')} cells and ${r.compared.formulas.toLocaleString('en')} formulas compared.`,
    ),
  );
}

function exportsPanel(r: Comparison): HTMLElement {
  const files = { original: original!.name, edited: edited!.name };
  return h(
    'section',
    { class: 'exports' },
    h('h3', {}, 'Export'),
    h(
      'button',
      {
        class: 'button small primary',
        type: 'button',
        onclick: () => download('sheet-sentry-review.md', markdownReport(r, files), 'text/markdown'),
      },
      'Review summary (Markdown)',
    ),
    h(
      'button',
      {
        class: 'button small',
        type: 'button',
        onclick: () => download('sheet-sentry-changes.csv', csvReport(r), 'text/csv'),
      },
      'All changes (CSV)',
    ),
    h(
      'button',
      {
        class: 'button small',
        type: 'button',
        onclick: () => download('sheet-sentry-report.json', jsonReport(r, files)),
      },
      'Full report (JSON)',
    ),
  );
}

function reset() {
  requests.original++;
  requests.edited++;
  original = null;
  edited = null;
  result = null;
  selected = null;
  introError.hidden = true;
  filter = 'all';
  showInfo = false;
  sheetIndex = 0;
  top = 0;
  for (const which of ['original', 'edited']) {
    byId(`slot-${which}`).classList.remove('filled');
    byId(`name-${which}`).textContent = '.xlsx or .xlsm — or drop it here';
  }
  workspace.hidden = true;
  workspace.replaceChildren();
  intro.hidden = false;
}
