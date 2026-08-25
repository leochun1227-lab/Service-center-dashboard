import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_dashboard_fields_grouped_with_abnormal_categories.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/sheet_matrix_update";
const outputPath = `${outputDir}/service_center_dashboard_fields_grouped_abnormal_simple.xlsx`;

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItemAt(1);

sheet.showGridLines = true;
for (const table of sheet.tables.items) {
  table.delete();
}
sheet.getRange("A1:I40").clear({ applyTo: "all" });

const rows = [
  ["Scenario", "C4C Value", "SAP Value", "Ticket Qty", "Invoice Qty", "Invoice Amount", "Workflow Count", "Internal / External", "Export"],
  ["Normal invoice", "Create invoice", "Invoice exists", "+1", "+1", "+SAP amount", "Completed / Invoiced +1", "Internal/external +1 by Role_1001", "No"],
  ["Create invoice but invoice blank", "Create invoice", "Missing invoice", "+1", "+1", "$0", "Completed / Invoiced +1", "No amount split until SAP value", "Yes"],
  ["Other status but invoice exists", "Not Create invoice", "Invoice exists", "+1", "+1", "+SAP amount", "Current status unchanged", "Internal/external +1 by Role_1001", "Yes"],
  ["Cancel", "Cancel", "Any", "-1 from open", "0", "$0", "Cancelled +1 / open -1", "No invoice KPI", "Maybe"],
  ["Cancel invoice", "Cancel invoice", "Invoice cancelled/open SO", "+1 waiting if SO open", "0", "$0", "Invoice Cancelled +1", "No invoice KPI", "Maybe"],
  ["Role_1001 is dealer name", "Dealer name in customer field", "Any", "0", "-1 normal split", "-SAP amount from normal split", "No change", "Move to abnormal", "Yes"],
  ["SAP invoice, no C4C ticket", "Missing C4C", "Invoice exists", "Missing ticket +1", "+1 SAP-only review", "+SAP amount review", "No workflow ticket", "Unknown", "Yes"],
  ["C4C ticket, no SAP invoice", "Ticket exists", "Missing invoice", "+1", "0", "$0", "Waiting/open as status says", "No invoice KPI", "No unless status says Create invoice"],
  ["PDI C4C, SAP GR missing", "PDI exists", "No GR evidence", "+1 PDI", "0", "$0", "PDI workflow +1", "No invoice KPI", "Need SAP join"],
  ["SAP GR manual, no C4C inbound PDI", "Missing inbound PDI", "GR manual user", "Missing PDI +1", "0", "$0", "Inbound missing +1", "No invoice KPI", "Yes"],
  ["SAP outbound manual, no C4C outbound PDI", "Missing outbound PDI", "Outbound manual user", "Missing PDI +1", "0", "$0", "Outbound missing +1", "No invoice KPI", "Yes"],
  ["Actual labour missing", "No actual time", "No payroll/time", "0", "0", "$0", "No workflow change", "No invoice change", "Data gap"],
  ["Worker missing", "No assigned worker", "No worker source", "0", "0", "$0", "No workflow change", "No invoice change", "Data gap"],
  ["Payroll hours missing", "No payroll hours", "No payroll source", "0", "0", "$0", "No workflow change", "No invoice change", "Data gap"],
];

sheet.getRange("A1:I15").values = rows;

const used = sheet.getRange("A1:I15");
used.format.font = { name: "Aptos", size: 10, color: "#000000" };
used.format.wrapText = true;
used.format.horizontalAlignment = "center";
used.format.verticalAlignment = "center";
used.format.borders = { preset: "all", style: "thin", color: "#D9D9D9" };
used.format.fill = "#FFFFFF";

sheet.getRange("A1:I1").format.font = { bold: true, color: "#000000" };
sheet.getRange("A2:A15").format.font = { bold: true, color: "#000000" };
sheet.getRange("A1:I15").format.borders = {
  top: { style: "thin", color: "#BFBFBF" },
  bottom: { style: "thin", color: "#BFBFBF" },
  left: { style: "thin", color: "#D9D9D9" },
  right: { style: "thin", color: "#D9D9D9" },
  insideHorizontal: { style: "thin", color: "#D9D9D9" },
  insideVertical: { style: "thin", color: "#D9D9D9" },
};

sheet.getRange("A:A").format.columnWidthPx = 300;
sheet.getRange("B:B").format.columnWidthPx = 225;
sheet.getRange("C:C").format.columnWidthPx = 205;
sheet.getRange("D:D").format.columnWidthPx = 130;
sheet.getRange("E:E").format.columnWidthPx = 150;
sheet.getRange("F:F").format.columnWidthPx = 195;
sheet.getRange("G:G").format.columnWidthPx = 215;
sheet.getRange("H:H").format.columnWidthPx = 255;
sheet.getRange("I:I").format.columnWidthPx = 170;
sheet.getRange("1:1").format.rowHeightPx = 26;
sheet.getRange("2:15").format.rowHeightPx = 22;
sheet.getRange("6:6").format.rowHeightPx = 38;
sheet.getRange("8:8").format.rowHeightPx = 34;
sheet.getRange("9:9").format.rowHeightPx = 40;
sheet.getRange("11:12").format.rowHeightPx = 38;

sheet.freezePanes.freezeRows(1);

const check = await workbook.inspect({
  kind: "region",
  sheetId: sheet.name,
  range: "A1:I15",
  maxChars: 9000,
});
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const firstPreview = await workbook.render({
  sheetName: workbook.worksheets.getItemAt(0).name,
  range: "A1:K22",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/abnormal_simple_first_sheet.png`, new Uint8Array(await firstPreview.arrayBuffer()));

const secondPreview = await workbook.render({
  sheetName: sheet.name,
  range: "A1:I22",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/abnormal_simple_second_sheet.png`, new Uint8Array(await secondPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
