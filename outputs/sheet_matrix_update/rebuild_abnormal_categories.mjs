import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/sheet_matrix_update/service_center_dashboard_fields_grouped_by_source.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/sheet_matrix_update";
const outputPath = `${outputDir}/service_center_dashboard_fields_grouped_with_abnormal_categories.xlsx`;

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItemAt(1);

sheet.showGridLines = false;
for (const table of sheet.tables.items) {
  table.delete();
}
sheet.getRange("A1:K34").clear({ applyTo: "all" });

sheet.getRange("A1:K1").merge();
sheet.getRange("A1").values = [["Abnormal Examples - Categorised Count Impact"]];
sheet.getRange("A2:K2").merge();
sheet.getRange("A2").values = [["Classify abnormal cases by why the dashboard count or amount changes. Keep detailed ticket rows in the abnormal export."]];

const header = [
  "Category",
  "Scenario",
  "C4C condition",
  "SAP condition",
  "Ticket Qty impact",
  "Invoice Qty impact",
  "Invoice Amount impact",
  "Workflow / status impact",
  "Internal / External impact",
  "Export?",
  "Notes",
];

const rows = [
  [
    "1. Ticket status driven",
    "Create invoice but invoice blank",
    "TicketStatusText = Create invoice",
    "Invoice number / price missing",
    "+1 ticket",
    "+1 invoice count but $0",
    "$0",
    "Completed / Invoiced +1",
    "No reliable amount split until SAP value exists",
    "Yes",
    "Status says invoice stage, but billing evidence is incomplete.",
  ],
  [
    "1. Ticket status driven",
    "Other status but invoice exists",
    "TicketStatusText is not Create invoice",
    "Invoice exists",
    "+1 ticket",
    "+1 invoice",
    "+SAP amount",
    "Current C4C status unchanged",
    "Internal / external +1 by Role_1001",
    "Yes",
    "Status and invoice evidence disagree.",
  ],
  [
    "1. Ticket status driven",
    "Cancel",
    "TicketStatusText = Cancel",
    "Any",
    "-1 from open",
    "0",
    "$0",
    "Cancelled +1 / open -1",
    "No invoice KPI",
    "Maybe",
    "Normally not a billing abnormal unless paired with SAP invoice evidence.",
  ],
  [
    "1. Ticket status driven",
    "Cancel invoice",
    "TicketStatusText = Cancel invoice",
    "Invoice cancelled / SO still open",
    "+1 waiting if SO open",
    "0",
    "$0",
    "Invoice Cancelled +1",
    "No invoice KPI",
    "Maybe",
    "Useful as workflow exception rather than invoice revenue.",
  ],
  [
    "2. Record exists in one system only",
    "SAP invoice, no C4C ticket",
    "Missing C4C ticket",
    "Invoice exists",
    "Missing ticket +1",
    "+1 SAP-only review",
    "+SAP amount review",
    "No workflow ticket",
    "Unknown",
    "Yes",
    "SAP billing exists but no C4C service ticket to count in workflow.",
  ],
  [
    "2. Record exists in one system only",
    "C4C ticket, no SAP invoice",
    "Ticket exists",
    "Missing invoice",
    "+1 ticket",
    "0",
    "$0",
    "Waiting / open as status says",
    "No invoice KPI",
    "No unless status says Create invoice",
    "Normal open ticket unless C4C status implies billing should exist.",
  ],
  [
    "2. Record exists in one system only",
    "PDI C4C, SAP GR missing",
    "PDI exists",
    "No GR evidence",
    "+1 PDI",
    "0",
    "$0",
    "PDI workflow +1",
    "No invoice KPI",
    "Need SAP join",
    "C4C says PDI exists, SAP goods receipt / delivery support is missing.",
  ],
  [
    "2. Record exists in one system only",
    "SAP GR manual, no C4C inbound PDI",
    "Missing inbound PDI",
    "GR manual user",
    "Missing PDI +1",
    "0",
    "$0",
    "Inbound missing +1",
    "No invoice KPI",
    "Yes",
    "SAP has operational evidence, but C4C PDI record is missing.",
  ],
  [
    "2. Record exists in one system only",
    "SAP outbound manual, no C4C outbound PDI",
    "Missing outbound PDI",
    "Outbound manual user",
    "Missing PDI +1",
    "0",
    "$0",
    "Outbound missing +1",
    "No invoice KPI",
    "Yes",
    "Outbound side exists in SAP but no matching C4C PDI.",
  ],
  [
    "3. C4C / SAP mismatch",
    "Normal invoice",
    "Create invoice",
    "Invoice exists",
    "+1",
    "+1",
    "+SAP amount",
    "Completed / Invoiced +1",
    "Internal / external +1 by Role_1001",
    "No",
    "This is the expected aligned case, kept as baseline.",
  ],
  [
    "3. C4C / SAP mismatch",
    "Role_1001 is dealer name",
    "Dealer name in customer field",
    "Any",
    "0 normal ticket split",
    "-1 from normal invoice split",
    "-SAP amount from normal split",
    "No workflow change",
    "Move to abnormal / unknown scope",
    "Yes",
    "C4C party field is not a valid customer/internal party for dashboard split.",
  ],
  [
    "3. C4C / SAP mismatch",
    "Invoice amount mismatch",
    "C4C amount / expected amount differs",
    "SAP invoice price differs",
    "0",
    "0",
    "Variance review",
    "No workflow change",
    "No scope change",
    "Yes",
    "Use when both systems have records but amount does not reconcile.",
  ],
  [
    "3. C4C / SAP mismatch",
    "Dealer / bill-to mismatch",
    "C4C dealer or Role_1001 differs",
    "SAP bill-to differs",
    "0",
    "0",
    "Possible reclass",
    "No workflow change",
    "May change internal / external split",
    "Yes",
    "Use when record joins but customer/dealer identity conflicts.",
  ],
  [
    "4. Missing data fields",
    "Actual labour missing",
    "No actual time",
    "No payroll / time",
    "0",
    "0",
    "$0",
    "No workflow change",
    "No invoice change",
    "Data gap",
    "Dashboard can count tickets but cannot finalise productivity.",
  ],
  [
    "4. Missing data fields",
    "Worker missing",
    "No assigned worker",
    "No worker source",
    "0",
    "0",
    "$0",
    "No workflow change",
    "No invoice change",
    "Data gap",
    "Dashboard cannot split workload by technician.",
  ],
  [
    "4. Missing data fields",
    "Payroll hours missing",
    "No payroll hours",
    "No payroll source",
    "0",
    "0",
    "$0",
    "No workflow change",
    "No invoice change",
    "Data gap",
    "Assigned vs payable hours cannot be calculated.",
  ],
  [
    "4. Missing data fields",
    "Billing date missing",
    "Create invoice / invoice no. present",
    "Billing date blank",
    "0",
    "+1 if invoice exists",
    "+SAP amount if price exists",
    "Completed / Invoiced retained",
    "Scope unchanged",
    "Yes",
    "Invoice month trend cannot be trusted without billing date.",
  ],
];

sheet.getRange("A4:K4").values = [header];
sheet.getRange(`A5:K${4 + rows.length}`).values = rows;

sheet.getRange("A1").format.font = { name: "Aptos Display", size: 16, bold: true, color: "#111827" };
sheet.getRange("A1").format.horizontalAlignment = "center";
sheet.getRange("A2").format.font = { name: "Aptos", size: 9, italic: true, color: "#475569" };
sheet.getRange("A2").format.horizontalAlignment = "center";

const tableRange = sheet.getRange(`A4:K${4 + rows.length}`);
tableRange.format.font = { name: "Aptos", size: 9, color: "#111827" };
tableRange.format.wrapText = true;
tableRange.format.verticalAlignment = "center";
tableRange.format.borders = { preset: "all", style: "thin", color: "#D1D5DB" };

sheet.getRange("A4:K4").format = {
  fill: "#111827",
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "all", style: "thin", color: "#111827" },
};

const categoryColors = {
  "1. Ticket status driven": "#DBEAFE",
  "2. Record exists in one system only": "#DCFCE7",
  "3. C4C / SAP mismatch": "#FEF3C7",
  "4. Missing data fields": "#FCE7F3",
};

for (let index = 0; index < rows.length; index += 1) {
  const rowNumber = 5 + index;
  const category = rows[index][0];
  const fill = categoryColors[category] ?? "#F8FAFC";
  sheet.getRange(`A${rowNumber}:K${rowNumber}`).format.fill = fill;
  sheet.getRange(`A${rowNumber}`).format.font = { bold: true, color: "#111827" };
}

sheet.getRange("A5:A8").merge(true);
sheet.getRange("A9:A13").merge(true);
sheet.getRange("A14:A17").merge(true);
sheet.getRange("A18:A21").merge(true);
sheet.getRange("A5:A21").format.verticalAlignment = "center";
sheet.getRange("A5:A21").format.horizontalAlignment = "center";
sheet.getRange("A5:A21").format.font = { bold: true, color: "#111827" };

sheet.getRange("A:A").format.columnWidthPx = 190;
sheet.getRange("B:B").format.columnWidthPx = 220;
sheet.getRange("C:D").format.columnWidthPx = 210;
sheet.getRange("E:G").format.columnWidthPx = 135;
sheet.getRange("H:H").format.columnWidthPx = 210;
sheet.getRange("I:I").format.columnWidthPx = 215;
sheet.getRange("J:J").format.columnWidthPx = 92;
sheet.getRange("K:K").format.columnWidthPx = 260;

sheet.getRange("1:1").format.rowHeightPx = 34;
sheet.getRange("2:2").format.rowHeightPx = 24;
sheet.getRange("3:3").format.rowHeightPx = 10;
sheet.getRange("4:4").format.rowHeightPx = 44;
sheet.getRange("5:21").format.rowHeightPx = 58;

sheet.freezePanes.freezeRows(4);

const rangeAddress = `A4:K${4 + rows.length}`;
const table = sheet.tables.add(rangeAddress, true, "AbnormalCategoryTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;

const check = await workbook.inspect({
  kind: "region",
  sheetId: sheet.name,
  range: "A1:K22",
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
await fs.writeFile(`${outputDir}/abnormal_categories_first_sheet.png`, new Uint8Array(await firstPreview.arrayBuffer()));

const secondPreview = await workbook.render({
  sheetName: sheet.name,
  range: "A1:K22",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/abnormal_categories_second_sheet.png`, new Uint8Array(await secondPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
