const { chromium } = require('playwright');
const { pathToFileURL } = require('node:url');
const path = require('node:path');
const fs = require('node:fs');
const assert = require('node:assert/strict');

(async () => {
  const browser = await chromium.launch({ headless: true, channel: 'msedge' });
  try {
    const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(pathToFileURL(path.resolve('overview.html')).href);
    assert.deepEqual(await page.locator('.kpi-value').allTextContents(), JSON.parse(fs.readFileSync('outputs/overview-kpi-before.json', 'utf8')));
    for (const width of [1920, 1440, 1024, 390]) {
      await page.setViewportSize({ width, height: 1000 });
      const bounds = await page.evaluate(() => {
        const title = document.querySelector('#pageTitle').getBoundingClientRect();
        const controls = [...document.querySelectorAll('.filters button, .filters select, .kpi-title, .kpi-value')];
        const overlaps = [...document.querySelectorAll('.kpi')].some(node => {
          const header = node.querySelector('.kpi-heading').getBoundingClientRect();
          const value = node.querySelector('.kpi-value').getBoundingClientRect();
          const foot = node.querySelector('.kpi-foot').getBoundingClientRect();
          return header.bottom > value.top + 1 || value.bottom > foot.top + 1;
        });
        return { titleHeight: title.height, overflow: controls.filter(n => n.scrollWidth > n.clientWidth + 1).map(n => n.textContent), overlaps,
          bottom: document.querySelector('#kpis').getBoundingClientRect().bottom,
          left: document.querySelector('main').getBoundingClientRect().left,
          viewportOverflow: document.documentElement.scrollWidth > innerWidth };
      });
      assert(!bounds.overlaps);
      assert.deepEqual(bounds.overflow, []);
      if (width >= 1440) { assert(bounds.titleHeight < 40); assert(!bounds.viewportOverflow); }
      await page.screenshot({ path: `outputs/overview-header-${width}.png`, clip: { x: bounds.left, y: 0, width: width - bounds.left, height: Math.ceil(bounds.bottom + 10) } });
      console.log('Layout OK', width, bounds.bottom);
    }
    await page.setViewportSize({ width: 1440, height: 1000 });
    await page.locator('#dealerYard').selectOption('Perth');
    assert.equal(await page.evaluate(() => state.yard), 'Perth');
    await page.locator('#dateRange').selectOption('Aug 2026');
    assert.equal(await page.evaluate(() => state.month), 'Aug 2026');
    await page.locator('#serviceToButtons').getByRole('button', { name: 'Internal Sales', exact: true }).click();
    assert.equal(await page.evaluate(() => state.serviceTo), 'Internal Sales');
    await page.locator('#serviceToButtons').getByRole('button', { name: 'Retail', exact: true }).click();
    await page.locator('#workTypeButtons').getByRole('button', { name: 'Repair', exact: true }).click();
    assert.equal(await page.evaluate(() => state.workType), 'Repair');
    const pending = page.waitForEvent('download');
    await page.locator('#exportCsv').click();
    const download = await pending;
    assert(fs.statSync(await download.path()).size > 1000);
    for (const tab of ['workflow', 'service-orders', 'overview']) {
      await page.locator(`[data-page="${tab}"]`).click();
      assert.equal(await page.locator('.filters').isVisible(), tab === 'overview');
    }
    assert.deepEqual(errors, []);
    console.log('PASS: unchanged KPI values, header layout, filters, export, module navigation');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
