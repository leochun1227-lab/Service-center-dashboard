import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/latest_ticket_export_like_reference.json";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/01a01c4f-25b7-7e91-8ffd-41a61b3fc906";
const outputPath = path.join(outputDir, "c4c_ticket_table_z007_z010_with_invoice_layout_checked_latest_like_reference.xlsx");
const previewPath = path.join(outputDir, "latest_ticket_export_preview.png");

const data = JSON.parse(await fs.readFile(inputPath, "utf8"));

function colLetter(index) {
  let n = index;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function sanitizeTableName(name) {
  return `${name.replace(/[^A-Za-z0-9_]/g, "")}Table`;
}

function writeSheet(workbook, sheetData) {
  const sheet = workbook.worksheets.add(sheetData.name);
  sheet.showGridLines = false;
  const columns = sheetData.columns;
  const matrix = [columns, ...sheetData.rows];
  const rowCount = matrix.length;
  const colCount = columns.length;
  const lastCol = colLetter(colCount);
  const range = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
  range.values = matrix;

  sheet.freezePanes.freezeRows(1);
  sheet.getRange(`A1:${lastCol}1`).format = {
    fill: "#263746",
    font: { bold: true, color: "#FFFFFF" },
  };
  range.format.borders = {
    insideHorizontal: { style: "thin", color: "#E6EEF7" },
    top: { style: "thin", color: "#C7D3E3" },
    bottom: { style: "thin", color: "#C7D3E3" },
  };
  if (sheetData.name === "Tickets" || sheetData.name === "NotAssigned") {
    const amountCols = ["ERPInvoiceNumberPrice", "AmountIncludingTax", "TotalLabourHours"];
    for (const name of amountCols) {
      const idx = columns.indexOf(name);
      if (idx >= 0) {
        const letter = colLetter(idx + 1);
        sheet.getRange(`${letter}:${letter}`).format.numberFormat = name === "TotalLabourHours" ? "#,##0.00" : "$#,##0.00";
      }
    }
  }
  sheet.tables.add(`A1:${lastCol}${rowCount}`, true, sanitizeTableName(sheetData.name));
  range.format.autofitColumns();
  range.format.autofitRows();
  return sheet;
}

const workbook = Workbook.create();
for (const sheetData of data.sheets) {
  writeSheet(workbook, sheetData);
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Tickets",
  range: "A1:L24",
  scale: 1,
  format: "png",
});
await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath }, null, 2));
