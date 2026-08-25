import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_data_center_.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/field_matrix_edit_work";
const outputPath = path.join(outputDir, "service_center_data_center_dashboard_focused.xlsx");
const previewPath = path.join(outputDir, "fact_labour_after.png");

await fs.mkdir(outputDir, { recursive: true });

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("4_FACT_LABOUR");

const before = await workbook.inspect({
  kind: "region",
  sheetId: "4_FACT_LABOUR",
  range: "A1:F16",
  maxChars: 7000,
});
console.log("BEFORE_INSPECT");
console.log(before.ndjson);

const focusedRows = [
  [
    "ServiceOrderID",
    "Yes",
    "C4C / SAP",
    "TicketID",
    "Tickets sheet / SAP labour patch grouped by TicketID",
    "Join labour hours back to the service order for dashboard reconciliation",
  ],
  [
    "LabourDate",
    "Missing",
    "Missing",
    null,
    null,
    "Daily labour view and date-level claim vs actual vs paid reconciliation",
  ],
  [
    "WorkerName",
    "Partial",
    "SAP / C4C",
    "Worker / Role_40_InvolvedPartyName",
    "SAP labour loader has Worker internally; C4C involved party can be a weak proxy",
    "Technician / worker display for labour accountability",
  ],
  [
    "ClaimHours",
    "Yes",
    "C4C",
    "LabourHours",
    "C4C / Tickets sheet",
    "Claimed labour hours",
  ],
  [
    "ActualWorkHours",
    "Partial",
    "SAP",
    "TotalLabourHours",
    "SAP labour patch / Tickets sheet at ticket-level total",
    "Actual worked-hour comparison against claim hours",
  ],
  [
    "InvoicePaidHours",
    "Missing",
    "Missing",
    null,
    null,
    "Invoice-paid labour hours for finance reconciliation against claim and actual hours",
  ],
];

sheet.getRange("A5:F10").values = focusedRows;
sheet.getRange("A11:F14").clear({ applyTo: "contents" });
sheet.getRange("A5:F10").format.wrapText = true;
sheet.getRange("D:D").format.columnWidth = 26;
sheet.getRange("E:E").format.columnWidth = 42;
sheet.getRange("F:F").format.columnWidth = 48;
sheet.getRange("A5:F10").format.autofitRows();

const after = await workbook.inspect({
  kind: "region",
  sheetId: "4_FACT_LABOUR",
  range: "A1:F16",
  maxChars: 7000,
});
console.log("AFTER_INSPECT");
console.log(after.ndjson);

const preview = await workbook.render({
  sheetName: "4_FACT_LABOUR",
  range: "A1:F14",
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log("ERROR_SCAN");
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath }, null, 2));
