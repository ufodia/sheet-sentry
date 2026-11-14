import { expect, test, type Page } from '@playwright/test';
import { readFile } from 'node:fs/promises';

async function loadDemo(page: Page) {
  await page.goto('/');
  await page.getByRole('button', { name: 'Review the demo budget' }).click();
  await expect(page.locator('.verdict-bar')).toBeVisible();
}

const count = (page: Page, label: string) => page.locator('.count', { hasText: label }).locator('strong');

test('demo: verdict, counts and the first critical change', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await loadDemo(page);
  await expect(page.locator('html')).toHaveAttribute('lang', 'en');
  await expect(page.locator('.verdict-bar h1')).toHaveText('Risky edit');
  await expect(count(page, 'Critical')).toHaveText('3');
  await expect(count(page, 'High')).toHaveText('3');
  await expect(count(page, 'Medium')).toHaveText('3');
  await expect(page.getByRole('tab', { name: /Revenue/ })).toHaveAttribute('aria-selected', 'true');
  await expect(page.locator('.inspector h2')).toHaveText('Range no longer covers all rows');
  await expect(page.locator('.inspector')).toContainText('=SUM(D2:D4)');
  await expect(page.locator('.inspector')).toContainText('=SUM(D2:D3)');
  await expect(page.locator('.grid td.selected')).toHaveText('44,100');
  await expect(page.getByRole('tab', { name: /Assumptions/ })).toContainText('hidden');
  expect(errors).toEqual([]);
});

test('jumping from the list to a cell on another sheet', async ({ page }) => {
  await loadDemo(page);
  await page.locator('.change', { hasText: 'Formula replaced by its own result' }).click();
  await expect(page.getByRole('tab', { name: /Summary/ })).toHaveAttribute('aria-selected', 'true');
  await expect(page.locator('.inspector .where')).toHaveText('Summary!B3');
  await expect(page.locator('.inspector')).toContainText('=B1-B2');
  await expect(page.locator('.inspector')).toContainText('(constant)');
  await expect(page.locator('.grid td.selected')).toHaveText('167,400');

  await page
    .locator('.change', { hasText: 'Range no longer covers all rows' })
    .filter({ hasText: 'Costs' })
    .click();
  await expect(page.locator('.inspector .where')).toHaveText('Costs!B7 (was B6)');
  await expect(page.locator('.inspector .notes')).toContainText('B2:B6 became B2:B5');
  await expect(page.locator('.inspector .expected code')).toHaveText('=SUM(B2:B6)');
  await expect(page.locator('.inspector .state-label').first()).toHaveText('Before · B6');
  await expect(page.locator('.grid tr.hidden-row')).toHaveCount(1);
});

test('severity filters and recalculated results', async ({ page }) => {
  await loadDemo(page);
  const items = page.locator('.changes li');
  await expect(items).toHaveCount(9);
  await page.locator('.count', { hasText: 'Critical' }).click();
  await expect(items).toHaveCount(3);
  await expect(page.locator('.findings h3')).toHaveText('Critical changes');
  await page.locator('.count', { hasText: 'Critical' }).click();
  await expect(items).toHaveCount(9);
  await page.getByLabel('Show recalculated results').check();
  await expect(items).toHaveCount(18);
  await expect(page.locator('.change.sev-info').first()).toBeVisible();
});

test('real files through the file inputs, and unsupported files', async ({ page }) => {
  await page.goto('/');
  await page.locator('#file-original').setInputFiles({
    name: 'budget.xls',
    mimeType: 'application/vnd.ms-excel',
    buffer: Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
  });
  await expect(page.locator('#intro-error')).toContainText('legacy .xls');
  await page.locator('#file-original').setInputFiles({
    name: 'q3-budget.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await readFile('examples/q3-budget.xlsx'),
  });
  await expect(page.locator('#intro-error')).toBeHidden();
  await expect(page.locator('#name-original')).toContainText('4 sheets');
  await page.locator('#file-edited').setInputFiles({
    name: 'q3-budget-ai-edit.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await readFile('examples/q3-budget-ai-edit.xlsx'),
  });
  await expect(page.locator('.verdict-bar h1')).toHaveText('Risky edit');
  await expect(page.locator('.verdict-bar .eyebrow')).toHaveText('q3-budget.xlsx → q3-budget-ai-edit.xlsx');
  await page.getByRole('button', { name: 'New comparison' }).click();
  await expect(page.locator('#intro')).toBeVisible();
});

test('identical workbooks are clean', async ({ page }) => {
  await page.goto('/');
  const file = {
    name: 'q3-budget.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await readFile('examples/q3-budget.xlsx'),
  };
  await page.locator('#file-original').setInputFiles(file);
  await page.locator('#file-edited').setInputFiles(file);
  await expect(page.locator('.verdict-bar h1')).toHaveText('No meaningful changes');
  await expect(page.locator('.findings')).toContainText('No changes in this view.');
});

test('exports and the demo files are generated in the tab', async ({ page }) => {
  await loadDemo(page);
  const [md] = await Promise.all([
    page.waitForEvent('download'),
    page.getByRole('button', { name: 'Review summary (Markdown)' }).click(),
  ]);
  expect(md.suggestedFilename()).toBe('sheet-sentry-review.md');
  const markdown = await readFile((await md.path())!, 'utf8');
  expect(markdown).toContain('# Sheet Sentry review');
  expect(markdown).toContain('**Verdict:** Risky');

  const [csv] = await Promise.all([
    page.waitForEvent('download'),
    page.getByRole('button', { name: 'All changes (CSV)' }).click(),
  ]);
  const text = await readFile((await csv.path())!, 'utf8');
  expect(text.split('\n')[0]).toBe('severity,kind,sheet,cell,original_cell,before,after,detail');
  // Formulas in the CSV are neutralized so a spreadsheet app shows them as text.
  expect(text).toContain("'=SUM(");

  const [json] = await Promise.all([
    page.waitForEvent('download'),
    page.getByRole('button', { name: 'Full report (JSON)' }).click(),
  ]);
  const report = JSON.parse(await readFile((await json.path())!, 'utf8'));
  expect(report.verdict).toBe('risky');
  expect(report.tool.name).toBe('Sheet Sentry');

  await page.getByRole('button', { name: 'New comparison' }).click();
  const [first] = await Promise.all([
    page.waitForEvent('download'),
    page.getByRole('button', { name: 'Download the demo files' }).click(),
  ]);
  expect(first.suggestedFilename()).toBe('q3-budget.xlsx');
  const bytes = await readFile((await first.path())!);
  expect(bytes.subarray(0, 2).toString()).toBe('PK');
});

test('no network requests leave the page origin', async ({ page, baseURL }) => {
  const origin = new URL(baseURL!).origin;
  const outside: string[] = [];
  // Local security software (for example Kaspersky web protection) injects its own script into pages
  // served over plain HTTP. Those requests are not made by Sheet Sentry, so they are ignored here.
  const injected = /\.kaspersky-labs\.com\//;
  page.on('request', (request) => {
    const url = request.url();
    if (injected.test(url)) return;
    if (!url.startsWith(origin) && !url.startsWith('blob:') && !url.startsWith('data:')) outside.push(url);
  });
  await loadDemo(page);
  await page.locator('.change').nth(2).click();
  expect(outside).toEqual([]);
});

test('phone layout has no horizontal scroll', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await loadDemo(page);
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
  expect(overflow).toBeLessThanOrEqual(0);
  await expect(page.locator('.inspector')).toBeVisible();
});
