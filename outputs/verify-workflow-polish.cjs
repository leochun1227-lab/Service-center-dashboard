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
    await page.locator('[data-page="workflow"]').click();
    for (const width of [1920, 1440, 1024, 390]) {
      await page.setViewportSize({ width, height: 1000 });
      const layout = await page.evaluate(() => ({
        width: document.documentElement.scrollWidth,
        cards: [...document.querySelectorAll('.workflow-node')].map(node => {
          const bounds = node.getBoundingClientRect();
          return { top: bounds.top, bottom: bounds.bottom, section: node.closest('.workflow-section').className,
            rows: ['.workflow-node-metrics', '.workflow-aging', '.workflow-share', '.workflow-node-foot'].map(s => node.querySelector(s).getBoundingClientRect().top),
            overflow: [...node.querySelectorAll('.workflow-aging-value, .workflow-node-title, .workflow-export')].some(n => n.scrollWidth > n.clientWidth + 1)
          };
        }),
        tracks: [...document.querySelectorAll('.workflow-track')].map(n => ({ width: n.clientWidth, scroll: n.scrollWidth }))
      }));
      assert.equal(layout.cards.length, 12);
      assert(layout.width <= width);
      assert(layout.tracks.every(n => n.scroll <= n.width));
      assert(layout.cards.every(n => !n.overflow));
      for (const a of layout.cards) for (const b of layout.cards) {
        if (a.section === b.section && Math.abs(a.top - b.top) < 1) {
          assert(Math.abs(a.bottom - b.bottom) < 1);
          assert(a.rows.every((v, i) => Math.abs(v - b.rows[i]) < 1));
        }
      }
      await page.screenshot({ path: `outputs/workflow-polished-${width}.png`, fullPage: true });
      console.log('Layout OK', width, 'bottom:', layout.cards.at(-1).bottom);
    }
    await page.setViewportSize({ width: 1440, height: 1000 });
    for (const dealer of ['Perth', 'Traralgon', 'Launceston', 'Geelong', 'Frankston']) {
      await page.locator(`.workflow-dealer-button[data-dealer="${dealer}"]`).click();
      assert.equal(await page.locator('#pageSubtitle').innerText(), dealer);
      assert.equal(await page.locator('.workflow-export').count(), 12);
      assert.equal(await page.locator('.workflow-dealer-button[aria-pressed="true"]').innerText(), dealer);
      const expected = await page.evaluate(() => ['Repair ticket', 'PDI ticket'].map(type => workflowSection(type, workflowPeriod()).metrics.total));
      const displayed = await page.locator('.workflow-section').evaluateAll(ns => ns.map(n => Number(n.querySelector('.workflow-pill strong').textContent.replaceAll(',', ''))));
      assert.deepEqual(displayed, expected);
    }
    await page.locator('.workflow-dealer-button[data-dealer="Perth"]').click();
    for (const section of ['repair', 'pdi']) {
      const button = page.locator(`.workflow-section.${section} .workflow-export`).first();
      const expected = await button.evaluate(n => {
        const section = workflowSection(n.dataset.workflowType, workflowPeriod());
        const stage = section.pipeline.find(item => item.status === n.dataset.workflowStage);
        return workflowStageTicketRows(stage).length;
      });
      const downloadPromise = page.waitForEvent('download');
      await button.click();
      const xml = fs.readFileSync(await (await downloadPromise).path(), 'utf8');
      assert(xml.includes('lastchangedtime'));
      assert(xml.includes('ClaimHours (TotalLabourHours)'));
      assert.equal((xml.match(/<Row>/g) || []).length, expected + 3);
      console.log('Export OK', section, expected);
    }
    await page.locator('[data-page="service-orders"]').click();
    assert(await page.locator('#serviceOrderSearchResults').isVisible());
    assert.deepEqual(errors, []);
    console.log('PASS: all 12 stages, aligned rows, responsive layout, five dealers, Repair/PDI exports, navigation');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
