// Exercise the actual webpage functions and SpreadsheetML download serializer.
// Usage: node tools/test_invoice_exports.cjs [--write-exports]
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'overview.html'), 'utf8');
const context = vm.createContext({ window: {}, console, Blob, URL });
vm.runInContext(fs.readFileSync(path.join(root, 'dashboard-data.js'), 'utf8'), context);
vm.runInContext(`const data = window.serviceCentreData; const page = data.pages.overview;
  const state = {month: '2026', yard: 'All', type: 'All Ticket Types'};
  ${html.match(/    const invoiceMoney = .*;/)[0]}`, context);
const names = [
  'parseDashboardDate', 'periodEndDate', 'rowInPeriod', 'rowMatchesScope',
  'ticketDetailsFor', 'getTicketDetails', 'getInvoiceMix', 'isMissingInvoiceDetail',
  'invoicePeriodMatches', 'invoiceDetailRows', 'invoiceAmount', 'invoiceTotal',
  'invoiceDocumentStatus', 'buildInvoiceDetailSheet', 'buildDashboardExportSheets',
  'buildTicketDetailSheet', 'ticketDetailRows', 'ticketDetailValues', 'serviceTypeText',
  'numberIfAvailable', 'exportText', 'downloadWorkbook', 'excelCell', 'sheetName',
  'xmlText', 'slug', 'exportCurrentView', 'exportInvoiceDetails'
];
for (const name of names) {
  const match = html.match(new RegExp(`    function ${name}\\([^]*?(?=\\n    (?:function|const|let) |\\n  </script>)`));
  assert.ok(match, `missing function ${name}`);
  vm.runInContext(match[0], context);
}
const run = code => vm.runInContext(code, context);
const cents = x => Math.round(x * 100);
const summaryValue = (sheet, field) => sheet.rows.find(row => row[0] === field)?.[1];
const netSum = sheet => {
  const index = sheet.rows[0].indexOf('InvoiceAmount (Net, AUD)');
  assert.ok(index >= 0);
  return sheet.rows.slice(1).reduce((sum, row) => sum + cents(row[index]), 0);
};
let checks = 0;
for (const month of run('data.meta.months')) {
  for (const yard of ['All', 'Perth', 'Traralgon', 'Launceston', 'Geelong', 'Frankston', 'Other']) {
    for (const type of run('data.meta.ticketTypes')) {
      context.scope = { month, yard, type };
      run('Object.assign(state, scope)');
      const result = run(`({mix: getInvoiceMix(), top: buildDashboardExportSheets(),
        invoice: buildInvoiceDetailSheet(invoiceDetailRows()),
        expected: page.monthlyDealerActivityByType[state.month][state.type]
          .filter(row => state.yard === 'All' || row.yard === state.yard)})`);
      const expected = result.expected.reduce((sum, row) => sum + cents(row.invoicedAmount), 0);
      assert.equal(cents(result.mix.totalAmount), expected, JSON.stringify(context.scope));
      assert.equal(netSum(result.invoice[1]), expected);
      assert.equal(cents(summaryValue(result.top[0], 'Invoice Amount')), expected);
      assert.deepEqual(result.top.slice(2), result.invoice, 'both buttons must export identical invoice sheets');
      assert.equal(summaryValue(result.top[0], 'New Service Orders'), result.expected.reduce((sum, row) => sum + row.newTickets, 0));
      const negatives = run('invoiceDetailRows().filter(row => invoiceAmount(row) < 0)');
      for (const row of negatives) {
        assert.equal(row.billingType, 'S1');
        assert.equal(cents(row.invoiceRawAmount), -cents(row.invoiceAmount));
      }
      checks++;
    }
  }
}
// Empty selections must not silently export a different month.
run(`Object.assign(state, {month: '1900', yard: 'All', type: 'All Ticket Types'})`);
assert.equal(run('getInvoiceMix().totalAmount'), 0);
assert.equal(run('buildDashboardExportSheets()[1].rows.length'), 1);
assert.equal(run('buildDashboardExportSheets()[3].rows.length'), 1);
// Missing CreatedDate must not drop an invoice dated in the selected period.
run(`Object.assign(state, {month: 'Sep 2026'});
  page.invoiceDetails.push({invoiceNo: 'TEST', invoiceAmount: 12.34, invoiceAmountLabel: '$12',
    createdDate: 'TBC', billingDate: '01/09/2026', dealerYard: 'Other', invoiceScope: 'Other'});`);
assert.equal(run("invoiceDetailRows().some(row => row.invoiceNo === 'TEST')"), true);
run('page.invoiceDetails.pop()');
// Capture the real button output through the real XML serializer.
let downloaded;
context.URL = { createObjectURL(blob) { downloaded = blob; return 'blob:test'; }, revokeObjectURL() {} };
context.document = { body: { appendChild() {} }, createElement() { return { click() {}, remove() {} }; } };
(async () => {
  for (const month of ['2026', 'Sep 2026']) {
    context.selectedPeriod = month;
    run('state.month = selectedPeriod');
    for (const [button, fn] of [['dashboard', 'exportCurrentView'], ['invoices', 'exportInvoiceDetails']]) {
      run(`${fn}()`);
      const xml = await downloaded.text();
      assert.ok(xml.includes('SAP cancellation (deducted from net total)'));
      assert.ok(xml.includes('SAP NETWR (Raw, AUD)'));
      assert.ok(xml.includes(`<Data ss:Type="Number">${run('getInvoiceMix().totalAmount')}</Data>`));
      if (process.argv.includes('--write-exports')) {
        const dir = path.join(root, 'outputs', 'invoice-export-validation');
        fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(path.join(dir, `${button}-${month.replaceAll(' ', '-')}.xls`), xml);
      }
    }
  }
  console.log(`PASS: ${checks} period/dealer/type scopes; both buttons, empty selection, missing created date, cancellation signs and real Excel XML.`);
})().catch(error => { console.error(error); process.exitCode = 1; });
