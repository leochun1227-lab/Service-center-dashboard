import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_dashboard_horizontal_value_matrix_updated.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/sheet_matrix_update";
const outputPath = `${outputDir}/service_center_dashboard_horizontal_value_matrix_updated_formatted.xlsx`;

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItemAt(0);

sheet.showGridLines = false;

for (const table of sheet.tables.items) {
  table.delete();
}
sheet.getRange("A1:N28").clear({ applyTo: "all" });

const rows = [
  ["", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", "", ""],
  [
    "Source / use",
    "Ticket\nID",
    "Dealer /\nYard",
    "Ticket\nType",
    "Created\nOn",
    "Quote\nAmount",
    "ERP\nInvoice\nNo.",
    "Billing\nDate",
    "Invoice\nAmount",
    "Internal /\nExternal",
    "Actual\nLabour\nHours",
    "Worker /\nTechnician",
    "Payroll\nPayable\nHours",
  ],
  [
    "C4C",
    "Yes",
    "WarrantyHandlingDealerID",
    "yes",
    "yes",
    "AmountIncludingTax",
    "from SAP",
    "no",
    "No finance amount from C4C",
    "Role_1001_InvolvedPartyName",
    "No actual worker time; only TotalLabourHours estimate",
    "No reliable assigned worker in this CSV",
    "No",
  ],
  [
    "SAP",
    "Short text",
    "bill to name",
    "no",
    "billing / GR / delivery dates",
    "No C4C quote amount",
    "yes",
    "yes",
    "yes",
    "no",
    "Usually no workshop actual labour unless labour confirmations exist",
    "Usually no assigned service worker unless another module / source",
    "No payroll hours",
  ],
  [
    "Service Center\nDashboard",
    "Use C4C",
    "Use C4C",
    "Use C4C",
    "Use C4C",
    "Use C4C",
    "Use SAP",
    "Use SAP",
    "Use SAP",
    "Use C4C",
    "Needed",
    "Needed",
    "Needed",
  ],
  [
    "Ticket\nexample",
    "38925",
    "Perth / Regent RV - Perth",
    "Z007 / Repair ticket",
    "2026-07-08",
    "$1,485",
    "0090046551",
    "2026-08-13",
    "$1,350",
    "REGENT RV PTY LTD",
    "Missing",
    "Missing",
    "Missing",
  ],
  [
    "Count /\nAmount impact",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Worker productivity / utilisation cannot be final",
    "Worker workload cannot be split",
    "Assigned vs payroll hours cannot be calculated",
  ],
];

sheet.getRange("A1:M9").values = rows;

sheet.getRange("A1:M1").merge();
sheet.getRange("A1").values = [["Service Center Dashboard - Horizontal Value Matrix"]];
sheet.getRange("A2:M2").merge();
sheet.getRange("A2").values = [["First sheet restyled to match the handwritten horizontal matrix: ticket example across columns, sources down rows."]];
sheet.getRange("B3:J3").merge();
sheet.getRange("B3").values = [["Fields already usable from C4C / SAP"]];
sheet.getRange("K3:M3").merge();
sheet.getRange("K3").values = [["Fields still missing / needed"]];

sheet.getRange("A1:M9").format.font = { name: "Aptos", size: 10, color: "#111827" };
sheet.getRange("A1").format.font = { name: "Aptos Display", size: 16, bold: true, color: "#111827" };
sheet.getRange("A2").format.font = { name: "Aptos", size: 9, italic: true, color: "#475569" };
sheet.getRange("A1:M2").format.fill = "#FFFFFF";

sheet.getRange("B3:J3").format = {
  fill: "#E8F1FF",
  font: { bold: true, color: "#1E3A8A" },
  borders: { preset: "outside", style: "medium", color: "#93C5FD" },
};
sheet.getRange("K3:M3").format = {
  fill: "#8BCB4A",
  font: { bold: true, color: "#12320A" },
  borders: { preset: "outside", style: "medium", color: "#6AA93B" },
};

// Re-apply group bands after merged-cell setup; this renders more consistently.
sheet.getRange("B3:J3").unmerge();
sheet.getRange("B3:J3").values = [["", "", "", "", "Fields already usable from C4C / SAP", "", "", "", ""]];
sheet.getRange("B3:J3").format = {
  fill: "#E8F1FF",
  font: { bold: true, color: "#1E3A8A" },
  borders: { preset: "all", style: "thin", color: "#93C5FD" },
};
sheet.getRange("K3:M3").unmerge();
sheet.getRange("K3:M3").values = [["", "Fields still missing / needed", ""]];
sheet.getRange("K3:M3").format = {
  fill: "#8BCB4A",
  font: { bold: true, color: "#12320A" },
  borders: { preset: "all", style: "thin", color: "#6AA93B" },
};

sheet.getRange("A4:M4").format = {
  fill: "#F8FAFC",
  font: { bold: true, color: "#111827" },
  borders: { preset: "all", style: "thin", color: "#CBD5E1" },
  wrapText: true,
};
sheet.getRange("B4:J4").format.fill = "#F8FAFC";
sheet.getRange("K4:M4").format.fill = "#8BCB4A";
sheet.getRange("K4:M9").format.fill = "#8BCB4A";
sheet.getRange("B4:B9").format.fill = "#FBE2D5";
sheet.getRange("A8:M8").format.fill = "#F3CBEF";

sheet.getRange("A5:A9").format = {
  font: { bold: true, color: "#111827" },
  fill: "#FFFFFF",
  borders: { preset: "all", style: "thin", color: "#CBD5E1" },
  wrapText: true,
};
sheet.getRange("B5:M9").format = {
  borders: { preset: "all", style: "thin", color: "#D7DEE8" },
  wrapText: true,
};
sheet.getRange("A4:M9").format.borders = { preset: "all", style: "thin", color: "#CBD5E1" };
sheet.getRange("A4:M9").format.borders = {
  top: { style: "medium", color: "#64748B" },
  bottom: { style: "medium", color: "#64748B" },
  left: { style: "medium", color: "#64748B" },
  right: { style: "medium", color: "#64748B" },
  insideHorizontal: { style: "thin", color: "#CBD5E1" },
  insideVertical: { style: "thin", color: "#CBD5E1" },
};
sheet.getRange("J4:J9").format.borders = {
  right: { style: "medium", color: "#64748B" },
};
sheet.getRange("K4:M9").format.borders = {
  top: { style: "medium", color: "#6AA93B" },
  bottom: { style: "medium", color: "#6AA93B" },
  left: { style: "medium", color: "#6AA93B" },
  right: { style: "medium", color: "#6AA93B" },
  insideHorizontal: { style: "thin", color: "#BDE69A" },
  insideVertical: { style: "thin", color: "#BDE69A" },
};

sheet.getRange("A1:M9").format.horizontalAlignment = "center";
sheet.getRange("A1:M9").format.verticalAlignment = "center";
sheet.getRange("A2").format.horizontalAlignment = "left";
sheet.getRange("C5:C9").format.horizontalAlignment = "left";
sheet.getRange("J5:J9").format.horizontalAlignment = "left";
sheet.getRange("K5:M9").format.horizontalAlignment = "center";

sheet.getRange("A:A").format.columnWidthPx = 155;
sheet.getRange("B:B").format.columnWidthPx = 118;
sheet.getRange("C:C").format.columnWidthPx = 205;
sheet.getRange("D:D").format.columnWidthPx = 145;
sheet.getRange("E:I").format.columnWidthPx = 130;
sheet.getRange("J:J").format.columnWidthPx = 185;
sheet.getRange("K:M").format.columnWidthPx = 210;

sheet.getRange("1:1").format.rowHeightPx = 34;
sheet.getRange("2:2").format.rowHeightPx = 26;
sheet.getRange("3:3").format.rowHeightPx = 28;
sheet.getRange("4:4").format.rowHeightPx = 96;
sheet.getRange("5:6").format.rowHeightPx = 78;
sheet.getRange("7:7").format.rowHeightPx = 58;
sheet.getRange("8:8").format.rowHeightPx = 92;
sheet.getRange("9:9").format.rowHeightPx = 78;

sheet.getRange("B5:B8").format.font = { bold: true, color: "#006100" };
sheet.getRange("K7:M8").format.font = { bold: true, color: "#111827" };

sheet.freezePanes.freezeRows(4);
sheet.freezePanes.freezeColumns(1);

const check = await workbook.inspect({
  kind: "region",
  sheetId: sheet.name,
  range: "A1:M9",
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

const firstPreview = await workbook.render({
  sheetName: sheet.name,
  range: "A1:M12",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/after_first_sheet.png`, new Uint8Array(await firstPreview.arrayBuffer()));

const second = workbook.worksheets.getItemAt(1);
const secondPreview = await workbook.render({
  sheetName: second.name,
  range: "A1:I18",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/after_second_sheet.png`, new Uint8Array(await secondPreview.arrayBuffer()));

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
