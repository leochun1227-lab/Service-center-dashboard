// Verify the actual chart/export functions against all monthly and annual totals.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'overview.html'), 'utf8');
const context = vm.createContext({ window: {}, console, Blob, URL });
vm.runInContext(fs.readFileSync(path.join(root, 'dashboard-data.js'), 'utf8'), context);
vm.runInContext(`const data = window.serviceCentreData; const page = data.pages.overview;
  const state = {month: 'Nov 2026', yard: 'Perth', type: 'Repair ticket'};`, context);
for (const name of ['labourExportYear', 'buildLabourExportSheets', 'exportLabourHours',
  'labourTotals', 'downloadWorkbook', 'excelCell', 'sheetName', 'xmlText']) {
  const match = html.match(new RegExp(`    function ${name}\\([^]*?(?=\\n    (?:function|const|let) |\\n  </script>)`));
  assert.ok(match, `Missing function ${name}`);
  vm.runInContext(match[0], context);
}
const run = code => vm.runInContext(code, context);
const cents = value => Math.round(value * 100);
const sheets = run('buildLabourExportSheets(2026)');
const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const dealers = run('data.meta.yards.filter(yard => yard !== "All Dealers")');
let checks = 0;
for (const [index, year] of [2025, 2026].entries()) {
  const sheet = sheets[index];
  assert.equal(sheet.name, `Monthly_Labour_${year}`);
  assert.equal(sheet.rows.length, 1 + 12 * dealers.length);
  for (const row of sheet.rows.slice(1)) {
    const [rowYear, period, dealer, repair, pdi, total] = row;
    assert.equal(rowYear, year);
    context.scope = {month: `${months[Number(period.slice(-2)) - 1]} ${year}`, yard: dealer, type: 'All Ticket Types'};
    run('Object.assign(state, scope)');
    const chart = run('labourTotals()');
    assert.equal(cents(repair), cents(chart.repair));
    assert.equal(cents(pdi), cents(chart.pdi));
    assert.equal(cents(total), cents(chart.total));
    checks++;
  }
  for (const dealer of dealers) {
    context.scope = {month: String(year), yard: dealer, type: 'All Ticket Types'};
    run('Object.assign(state, scope)');
    const annual = run('labourTotals()');
    for (const [col, key] of [[3, 'repair'], [4, 'pdi'], [5, 'total']]) {
      const sum = sheet.rows.slice(1).filter(row => row[2] === dealer).reduce((sum, row) => sum + cents(row[col]), 0);
      assert.ok(Math.abs(sum - cents(annual[key])) <= 1, `${year}/${dealer}/${key}: ${sum} vs ${cents(annual[key])}`);
    }
  }
}
// Changing page filters must never truncate this all-dealer, two-year export.
run(`Object.assign(state, {month: '2023', yard: 'Other', type: 'PDI ticket'})`);
assert.deepEqual(run('buildLabourExportSheets(2026)'), sheets);
assert.ok(sheets[0].rows.filter(row => row[1] === '2025-01').every(row => row[5] === 0));
assert.ok(sheets[1].rows.filter(row => row[1] === '2026-12').every(row => row[5] === 0));
// A new calendar year rolls both export sheets forward, with missing months zero-filled.
const nextYear = run('buildLabourExportSheets(2027)');
assert.equal(nextYear[0].name, 'Monthly_Labour_2026');
assert.equal(nextYear[1].name, 'Monthly_Labour_2027');
assert.ok(nextYear[1].rows.slice(1).every(row => row[5] === 0));
assert.equal(run('labourExportYear()'), Number(new Intl.DateTimeFormat('en-AU', {timeZone: 'Australia/Sydney', year: 'numeric'}).format(new Date())));
// Capture the download through the existing Excel serializer.
let downloaded, filename;
context.URL = {createObjectURL(blob) { downloaded = blob; return 'blob:test'; }, revokeObjectURL() {}};
context.document = {body: {appendChild() {}}, createElement() {return {click() {filename = this.download;}, remove() {}};}};
run('exportLabourHours()');
(async () => {
  const xml = await downloaded.text();
  const year = run('labourExportYear()');
  assert.equal(filename, `claim-labour-hours-by-dealer-${year - 1}-${year}.xls`);
  assert.ok(xml.includes(`<Worksheet ss:Name="Monthly_Labour_${year}">`));
  assert.ok(xml.includes('Repair Claim Hours'));
  assert.ok(xml.includes('PDI Claim Hours'));
  assert.ok(xml.includes('<Data ss:Type="Number">'));
  assert.ok(xml.includes('Read_Me'));
  console.log(`PASS: ${checks} month/dealer rows match the chart; annual totals, zero months, filter independence, year rollover and Excel download.`);
})().catch(error => {console.error(error); process.exitCode = 1;});
