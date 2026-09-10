// Usage: node tools/verify_quote_rejected.cjs <Python executable with pandas>
const { chromium } = require('playwright');
const { pathToFileURL } = require('node:url');
const { execFileSync } = require('node:child_process');
const path = require('node:path');
const fs = require('node:fs');
const assert = require('node:assert/strict');

(async () => {
  const fixture = JSON.parse(execFileSync(process.argv[2] || 'python',
    ['tools/test_quote_rejected.py', '--fixture'], { encoding: 'utf8' }));
  const browser = await chromium.launch({ headless: true, channel: 'msedge' });
  try {
    const tab = await browser.newPage({ viewport: { width: 1440, height: 1100 }, acceptDownloads: true });
    const errors = [];
    tab.on('pageerror', error => errors.push(error.message));
    await tab.goto(pathToFileURL(path.resolve('overview.html')).href);
    assert.equal(await tab.locator('#statusLegend [data-status="Quote Rejected"] strong').innerText(), '0');
    // All populated statuses are included in the donut, including those after the old top-five cutoff.
    assert.equal(await tab.locator('#statusLegend .donut-line').count(), await tab.evaluate(() =>
      getStatusMix().segments.filter(s => s.qty || s.name === 'Quote Rejected').length));
    fs.mkdirSync('outputs/quote-rejected', { recursive: true });
    await tab.screenshot({ path: 'outputs/quote-rejected/distribution.png', fullPage: true });
    assert.equal(await tab.evaluate(() => withQuoteRejectedStatus({ total: '0', segments: [] }, 'PDI ticket').segments.length), 0);
    await tab.locator('[data-page="workflow"]').click();
    for (const width of [1920, 1440, 1024, 390]) {
      await tab.setViewportSize({ width, height: 1100 });
      const layout = await tab.evaluate(() => ({
        width: document.documentElement.scrollWidth,
        count: document.querySelectorAll('.workflow-node').length,
        overflow: [...document.querySelectorAll('.workflow-node-title, .workflow-export, .workflow-aging-value')]
          .some(n => n.scrollWidth > n.clientWidth + 1)
      }));
      assert.equal(layout.count, 12);
      assert(layout.width <= width, `Page overflow at ${width}`);
      assert(!layout.overflow, `Card text overflow at ${width}`);
      if (width === 390) await tab.screenshot({ path: 'outputs/quote-rejected/workflow-mobile.png', fullPage: true });
    }
    await tab.setViewportSize({ width: 1440, height: 1100 });
    for (const dealer of ['Perth', 'Traralgon', 'Launceston', 'Geelong', 'Frankston']) {
      await tab.locator(`.workflow-dealer-button[data-dealer="${dealer}"]`).click();
      assert.equal(await tab.locator('.workflow-section.repair .workflow-export[data-workflow-stage="Quote Rejected"]').count(), 0);
      assert.equal(await tab.locator('.workflow-section.pdi .workflow-node').count(), 4);
    }
    await tab.locator('.workflow-dealer-button[data-dealer="Perth"]').click();
    await tab.screenshot({ path: 'outputs/quote-rejected/workflow.png', fullPage: true });
    // Integrate the Python-generated fixture without modifying production data.
    await tab.evaluate(fixture => {
      state.month = 'Sep 2026';
      state.yard = 'All';
      state.type = 'All Ticket Types';
      state.workflowDealer = 'Perth';
      page.workflowDaily[state.month]['Repair ticket'].Perth.pipeline = fixture.pipeline;
      page.monthlyOpenStatusMix[state.month] = fixture.mix;
      const sample = page.ticketDetails.find(row => row.serviceType === 'Repair ticket');
      page.ticketDetails = fixture.details.map(row => ({
        ...sample, ...row, serviceOrderId: row.ticketId,
        dealerYard: 'Perth', createdDate: '2026-09-01',
        lastchangedtime: '2026-09-02', invoiceNo: '', completedDate: '', period: state.month
      }));
      // Even previously generated data containing the stage must not reintroduce it in Workflow.
      page.workflowDaily[state.month]['Repair ticket'].Perth.pipeline.push({
        status: 'Quote Rejected', qty: 1, ticketIds: ['REJECTED'], aging: []
      });
      renderWorkflowPage(modulePages.workflow);
    }, fixture);
    assert.equal(await tab.locator('.workflow-export[data-workflow-stage="Quote Rejected"]').count(), 0);
    assert(!(await tab.locator('.workflow-section.repair').innerText()).includes('Quote Rejected'));
    assert.deepEqual(await tab.evaluate(() => workflowStageTicketRows(
      workflowSection('Repair ticket', workflowPeriod()).pipeline.find(s => s.status === 'Approved, Awaiting Repair')
    ).map(r => r.ticketId)), ['APPROVED']);
    await tab.locator('[data-page="overview"]').click();
    assert.equal(await tab.locator('#statusLegend [data-status="Quote Rejected"] strong').innerText(), '1');
    await tab.locator('#statusLegend [data-status="Quote Rejected"]').click();
    assert.deepEqual(await tab.evaluate(() => drilldownState.rows.map(r => r.ticketId)), ['REJECTED']);
    await tab.locator('[data-page="service-orders"]').click();
    const rateCard = tab.locator('.orders-metric').filter({ hasText: 'Rejected Rate' });
    assert.equal(await rateCard.locator('.orders-metric-value').innerText(), '25.0%');
    assert.equal(await rateCard.locator('.orders-metric-note').innerText(), '1 rejected / 4 repair tickets');
    assert(!(await tab.locator('#serviceOrderMetrics').innerText()).includes('Service Orders Trend'));
    // PDI is never part of the denominator; month and dealer are part of the scope.
    await tab.evaluate(() => {
      const sample = page.ticketDetails[0];
      page.ticketDetails.push(
        { ...sample, ticketId: 'PDI', serviceOrderId: 'PDI', serviceType: 'PDI ticket' },
        { ...sample, ticketId: 'OTHER-DEALER', serviceOrderId: 'OTHER-DEALER', dealerYard: 'Geelong' },
        { ...sample, ticketId: 'OLD', serviceOrderId: 'OLD', createdDate: '2026-08-01', period: 'Aug 2026' }
      );
      renderServiceOrdersPage(modulePages['service-orders']);
    });
    assert.equal(await rateCard.locator('.orders-metric-value').innerText(), '16.7%');
    await tab.locator('#serviceOrderMonth').selectOption('Sep 2026');
    assert.equal(await rateCard.locator('.orders-metric-value').innerText(), '20.0%');
    await tab.locator('#serviceOrderDealer').selectOption('Perth');
    assert.equal(await rateCard.locator('.orders-metric-value').innerText(), '25.0%');
    await tab.locator('#serviceOrderStatus').selectOption('Quote Rejected');
    assert.equal(await rateCard.locator('.orders-metric-value').innerText(), '25.0%');
    assert((await tab.locator('#serviceOrderSearchResults').innerText()).includes('REJECTED'));
    assert(!(await tab.locator('#serviceOrderSearchResults').innerText()).includes('APPROVED'));
    const downloadPromise = tab.waitForEvent('download');
    await tab.locator('#exportServiceOrderSearch').click();
    const xml = fs.readFileSync(await (await downloadPromise).path(), 'utf8');
    assert(xml.includes('REJECTED'));
    assert(!xml.includes('APPROVED'));
    await tab.locator('#serviceOrderType').selectOption('PDI ticket');
    assert.equal(await rateCard.locator('.orders-metric-value').innerText(), '—');
    await tab.locator('#resetServiceOrderFilters').click();
    await tab.goto(pathToFileURL(path.resolve('overview.html')).href);
    await tab.locator('[data-page="service-orders"]').click();
    await tab.screenshot({ path: 'outputs/quote-rejected/service-orders-rejected-rate.png', fullPage: true });
    assert.equal(await rateCard.locator('.orders-metric-value').innerText(), '0.0%');
    assert.deepEqual(errors, []);
    console.log('PASS: Workflow restored to 12 stages; rejection distribution, export and KPI; repair-only denominator, month/dealer scope, status filter independence and empty scope.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
