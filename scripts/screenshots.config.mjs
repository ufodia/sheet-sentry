/** Screens captured for the README. Every image shows the synthetic Q3 budget demo. */
const demo = async (page) => {
  await page.getByRole('button', { name: 'Review the demo budget' }).click();
  await page.locator('.verdict-bar').waitFor();
};
const pick = async (page, text, sheet) => {
  let change = page.locator('.change', { hasText: text });
  if (sheet) change = change.filter({ hasText: sheet });
  await change.first().click();
  await page.waitForTimeout(150);
};

export default {
  shots: [
    {
      file: 'docs/assets/screenshot-home.png',
      viewport: { width: 1440, height: 1000 },
      prepare: async () => {},
    },
    {
      file: 'docs/assets/screenshot-review.png',
      viewport: { width: 1440, height: 1040 },
      prepare: async (page) => {
        await demo(page);
        await pick(page, 'Range no longer covers all rows', 'Costs');
      },
    },
    {
      file: 'docs/assets/screenshot-hardcoded.png',
      viewport: { width: 1440, height: 1200 },
      prepare: async (page) => {
        await demo(page);
        await pick(page, 'Formula replaced by its own result');
      },
      element: '.viewer',
    },
    {
      file: 'docs/assets/screenshot-findings.png',
      viewport: { width: 1440, height: 1400 },
      prepare: async (page) => {
        await demo(page);
        await page.evaluate(() => {
          const list = document.querySelector('.changes');
          if (list) list.style.maxHeight = 'none';
        });
      },
      element: '.panel',
    },
    {
      file: 'docs/assets/screenshot-mobile.png',
      viewport: { width: 390, height: 844 },
      scale: 2,
      prepare: async (page) => {
        await demo(page);
        await pick(page, 'Formula replaced by its own result');
        await page.evaluate(() => window.scrollTo(0, 0));
      },
    },
  ],
  cards: [{ source: 'docs/assets/social-card.svg', file: 'public/og-image.png', width: 1200, height: 630 }],
};
