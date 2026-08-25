import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_dashboard_fields_grouped_with_abnormal_categories.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/sheet_matrix_update";
const outputPath = `${outputDir}/service_center_dashboard_fields_grouped_abnormal_simple_categories.xlsx`;

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItemAt(1);

sheet.showGridLines = true;
for (const table of sheet.tables.items) {
  table.delete();
}
sheet.getRange("A1:J45").unmerge();
sheet.getRange("A1:J45").clear({ applyTo: "all" });

const header = ["Category", "Scenario", "C4C Value", "SAP Value", "Ticket Qty", "Invoice Qty", "Invoice Amount", "Workflow Count", "Internal / External", "Export"];
const rows = [
  header,
  ["1. Ticket status driven", "Normal invoice", "Create invoice", "Invoice exists", "+1", "+1", "+SAP amount", "Completed / Invoiced +1", "Internal/external +1 by Role_1001", "No"],
  ["1. Ticket status driven", "Create invoice but invoice blank", "Create invoice", "Missing invoice", "+1", "+1", "$0", "Completed / Invoiced +1", "No amount split until SAP value", "Yes"],
  ["1. Ticket status driven", "Other status but invoice exists", "Not Create invoice", "Invoice exists", "+1", "+1", "+SAP amount", "Current status unchanged", "Internal/external +1 by Role_1001", "Yes"],
  ["1. Ticket status driven", "Cancel", "Cancel", "Any", "-1 from open", "0", "$0", "Cancelled +1 / open -1", "No invoice KPI", "Maybe"],
  ["1. Ticket status driven", "Cancel invoice", "Cancel invoice", "Invoice cancelled/open SO", "+1 waiting if SO open", "0", "$0", "Invoice Cancelled +1", "No invoice KPI", "Maybe"],
  ["2. Record exists in one system only", "SAP invoice, no C4C ticket", "Missing C4C", "Invoice exists", "Missing ticket +1", "+1 SAP-only review", "+SAP amount review", "No workflow ticket", "Unknown", "Yes"],
  ["2. Record exists in one system only", "C4C ticket, no SAP invoice", "Ticket exists", "Missing invoice", "+1", "0", "$0", "Waiting/open as status says", "No invoice KPI", "No unless status says Create invoice"],
  ["2. Record exists in one system only", "PDI C4C, SAP GR missing", "PDI exists", "No GR evidence", "+1 PDI", "0", "$0", "PDI workflow +1", "No invoice KPI", "Need SAP join"],
  ["2. Record exists in one system only", "SAP GR manual, no C4C inbound PDI", "Missing inbound PDI", "GR manual user", "Missing PDI +1", "0", "$0", "Inbound missing +1", "No invoice KPI", "Yes"],
  ["2. Record exists in one system only", "SAP outbound manual, no C4C outbound PDI", "Missing outbound PDI", "Outbound manual user", "Missing PDI +1", "0", "$0", "Outbound missing +1", "No invoice KPI", "Yes"],
  ["3. C4C / SAP mismatch", "Role_1001 is dealer name", "Dealer name in customer field", "Any", "0", "-1 normal split", "-SAP amount from normal split", "No change", "Move to abnormal", "Yes"],
  ["3. C4C / SAP mismatch", "Invoice amount mismatch", "C4C amount differs", "SAP invoice price differs", "0", "0", "Variance review", "No change", "No scope change", "Yes"],
  ["3. C4C / SAP mismatch", "Dealer / bill-to mismatch", "C4C dealer or Role_1001 differs", "SAP bill-to differs", "0", "0", "Possible reclass", "No change", "May change internal / external split", "Yes"],
  ["4. Missing data fields", "Actual labour missing", "No actual time", "No payroll/time", "0", "0", "$0", "No workflow change", "No invoice change", "Data gap"],
  ["4. Missing data fields", "Worker missing", "No assigned worker", "No worker source", "0", "0", "$0", "No workflow change", "No invoice change", "Data gap"],
  ["4. Missing data fields", "Payroll hours missing", "No payroll hours", "No payroll source", "0", "0", "$0", "No workflow change", "No invoice change", "Data gap"],
  ["4. Missing data fields", "Billing date missing", "Create invoice / invoice no. present", "Billing date blank", "0", "+1 if invoice exists", "+SAP amount if price exists", "Completed / Invoiced retained", "Scope unchanged", "Yes"],
];

sheet.getRange(`A1:J${rows.length}`).values = rows;

const used = sheet.getRange(`A1:J${rows.length}`);
used.format.font = { name: "Aptos", size: 10, color: "#000000" };
used.format.wrapText = true;
used.format.horizontalAlignment = "center";
used.format.verticalAlignment = "center";
used.format.fill = "#FFFFFF";
used.format.borders = { preset: "all", style: "thin", color: "#D9D9D9" };

sheet.getRange("A1:J1").format.font = { bold: true, color: "#000000" };
sheet.getRange(`A2:B${rows.length}`).format.font = { bold: true, color: "#000000" };

const categoryStarts = [2, 7, 12, 15];
for (const row of categoryStarts) {
  sheet.getRange(`A${row}:J${row}`).format.borders = {
    top: { style: "medium", color: "#808080" },
    bottom: { style: "thin", color: "#D9D9D9" },
    left: { style: "thin", color: "#D9D9D9" },
    right: { style: "thin", color: "#D9D9D9" },
    insideVertical: { style: "thin", color: "#D9D9D9" },
  };
}

sheet.getRange("A2:A6").merge();
sheet.getRange("A7:A11").merge();
sheet.getRange("A12:A14").merge();
sheet.getRange("A15:A18").merge();
sheet.getRange(`A2:A${rows.length}`).format.verticalAlignment = "center";
sheet.getRange(`A2:A${rows.length}`).format.horizontalAlignment = "center";

sheet.getRange("A:A").format.columnWidthPx = 230;
sheet.getRange("B:B").format.columnWidthPx = 300;
sheet.getRange("C:C").format.columnWidthPx = 225;
sheet.getRange("D:D").format.columnWidthPx = 205;
sheet.getRange("E:E").format.columnWidthPx = 130;
sheet.getRange("F:F").format.columnWidthPx = 150;
sheet.getRange("G:G").format.columnWidthPx = 195;
sheet.getRange("H:H").format.columnWidthPx = 215;
sheet.getRange("I:I").format.columnWidthPx = 255;
sheet.getRange("J:J").format.columnWidthPx = 170;
sheet.getRange("1:1").format.rowHeightPx = 26;
sheet.getRange(`2:${rows.length}`).format.rowHeightPx = 26;
sheet.getRange("6:6").format.rowHeightPx = 38;
sheet.getRange("8:8").format.rowHeightPx = 40;
sheet.getRange("10:11").format.rowHeightPx = 38;
sheet.getRange("13:14").format.rowHeightPx = 36;

sheet.freezePanes.freezeRows(1);

const check = await workbook.inspect({
  kind: "region",
  sheetId: sheet.name,
  range: `A1:J${rows.length}`,
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
const preview = await workbook.render({
  sheetName: sheet.name,
  range: "A1:J24",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/abnormal_simple_categories_second_sheet.png`, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
