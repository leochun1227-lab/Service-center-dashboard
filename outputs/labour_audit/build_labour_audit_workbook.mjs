import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const here = path.dirname(new URL(import.meta.url).pathname).replace(/^\/([A-Za-z]:)/, "$1");
const dataPath = path.join(here, "labour_audit_data.json");
const outputPath = path.join(here, "labour_hours_audit_20260820.xlsx");
const previewPath = path.join(here, "labour_hours_audit_preview.png");

const payload = JSON.parse(await fs.readFile(dataPath, "utf8"));
const workbook = Workbook.create();

const widths = {
  Dashboard_Check: [14, 14, 14, 12, 10, 17, 14, 12],
  Perth_Workers_Aug_2026: [13, 28, 12, 18, 15, 17, 18, 23, 13, 10, 13, 10],
  All_Workers_Aug_2026: [13, 28, 12, 18, 15, 17, 18, 23, 13, 10, 13, 10],
  Dealer_Type_Aug_2026: [14, 16, 12, 18],
  Status_Aug_2026: [14, 16, 34, 12, 18],
  Ticket_Detail_Aug_2026: [15, 11, 16, 28, 13, 18, 34, 18, 13, 28, 28, 15, 13, 18, 18, 18, 16],
};

const numberFormats = new Map([
  ["TicketCount", "#,##0"],
  ["Workers", "#,##0"],
  ["NonZeroHourTickets", "#,##0"],
  ["ZeroHourTickets", "#,##0"],
  ["TotalLabourHours", "#,##0.0000"],
  ["LabourHours", "#,##0.0000"],
  ["AvgHoursAllTickets", "#,##0.0000"],
  ["AvgHoursNonZeroTickets", "#,##0.0000"],
  ["RepairHours", "#,##0.0000"],
  ["PDIHours", "#,##0.0000"],
  ["RepairTickets", "#,##0"],
  ["PDITickets", "#,##0"],
  ["ERPInvoiceNumberPrice", "$#,##0.00"],
  ["AmountIncludingTax", "$#,##0.00"],
]);

function colLetter(index) {
  let n = index + 1;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}

function addSheet(name, rows) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const rowCount = rows.length;
  const colCount = rows[0]?.length ?? 1;
  const end = `${colLetter(colCount - 1)}${rowCount}`;
  sheet.getRange(`A1:${end}`).values = rows;
  sheet.getRange(`A1:${colLetter(colCount - 1)}1`).format = {
    fill: "#1F4E78",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange(`A1:${end}`).format.borders = {
    insideHorizontal: { style: "thin", color: "#D9E2EF" },
    bottom: { style: "thin", color: "#AFC5DD" },
  };
  sheet.freezePanes.freezeRows(1);
  const sheetWidths = widths[name] ?? [];
  for (let c = 0; c < colCount; c += 1) {
    const col = sheet.getRange(`${colLetter(c)}:${colLetter(c)}`);
    col.format.columnWidth = sheetWidths[c] ?? 16;
    if (numberFormats.has(rows[0][c])) {
      col.format.numberFormat = numberFormats.get(rows[0][c]);
    }
  }
  const tableName = `${name.replace(/[^A-Za-z0-9]/g, "")}Table`.slice(0, 240);
  if (rowCount > 1) {
    const table = sheet.tables.add(`A1:${end}`, true, tableName);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
  }
  return sheet;
}

for (const [name, rows] of Object.entries(payload.sheets)) {
  addSheet(name, rows);
}

const readme = workbook.worksheets.add("Read_Me");
readme.showGridLines = false;
readme.getRange("A1:B8").values = [
  ["Labour hours audit", ""],
  ["Source workbook", payload.source],
  ["Audit month", payload.auditMonth],
  ["Dashboard口径", "CreatedMonth + DealerBucket + TicketTypeBucket + sum(TotalLabourHours)"],
  ["AvgHoursAllTickets", "TotalLabourHours / all tickets, matches current webpage average"],
  ["AvgHoursNonZeroTickets", "TotalLabourHours / tickets where LabourHours is not zero"],
  ["网页截图 Perth total", "Dashboard_Check sheet should show WebMonthLabel Aug 2026 / Perth / 34.25h, displayed on webpage as 34.2h"],
  ["Notes", "This file is for checking only; it does not update dashboard source files."],
];
readme.getRange("A1:B1").merge();
readme.getRange("A1").format = {
  fill: "#1F4E78",
  font: { bold: true, color: "#FFFFFF", size: 14 },
};
readme.getRange("A2:A8").format = { fill: "#EAF2F8", font: { bold: true } };
readme.getRange("A1:B8").format.borders = { preset: "all", style: "thin", color: "#D9E2EF" };
readme.getRange("A:A").format.columnWidth = 28;
readme.getRange("B:B").format.columnWidth = 110;

const preview = await workbook.render({
  sheetName: "Perth_Workers_Aug_2026",
  range: "A1:L8",
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const check = await workbook.inspect({
  kind: "table",
  sheetId: "Dashboard_Check",
  range: "A1:H7",
  include: "values",
  tableMaxRows: 10,
  tableMaxCols: 10,
});
console.log(check.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath }));
