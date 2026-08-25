import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_dashboard_horizontal_value_matrix_updated.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/sheet_matrix_update";
const outputPath = `${outputDir}/service_center_dashboard_fields_grouped_by_source.xlsx`;

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItemAt(0);

sheet.showGridLines = false;

for (const table of sheet.tables.items) {
  table.delete();
}
sheet.getRange("A1:K30").clear({ applyTo: "all" });

sheet.getRange("A1:K1").merge();
sheet.getRange("A1").values = [["Service Center Dashboard Field Source Map"]];
sheet.getRange("A2:K2").merge();
sheet.getRange("A2").values = [["Fields are grouped by where they come from: final dashboard output, C4C, and SAP."]];

sheet.getRange("A4:C4").merge();
sheet.getRange("E4:G4").merge();
sheet.getRange("I4:K4").merge();
sheet.getRange("A4").values = [["Final Dashboard Fields"]];
sheet.getRange("E4").values = [["C4C Source Fields"]];
sheet.getRange("I4").values = [["SAP Source Fields"]];

sheet.getRange("A5:C5").values = [["Field wanted", "Example / value", "Purpose"]];
sheet.getRange("E5:G5").values = [["C4C field", "Example / value", "Used for"]];
sheet.getRange("I5:K5").values = [["SAP field", "Example / value", "Used for"]];

const dashboardRows = [
  ["Ticket ID", "38925", "Join key / ticket example"],
  ["Dealer / Yard", "Perth / Regent RV - Perth", "Dealer split"],
  ["Ticket Type", "Z007 / Repair ticket", "Repair vs PDI"],
  ["Created On", "2026-07-08", "Created month / trend"],
  ["Quote Amount", "$1,485", "Open quote amount"],
  ["ERP Invoice No.", "0090046551", "Invoice tracking"],
  ["Billing Date", "2026-08-13", "Invoice month"],
  ["Invoice Amount", "$1,350", "Revenue / invoice amount"],
  ["Internal / External", "REGENT RV PTY LTD", "Invoice scope split"],
  ["Actual Labour Hours", "Missing", "Productivity / utilisation"],
  ["Worker / Technician", "Missing", "Workload split"],
  ["Payroll Payable Hours", "Missing", "Payroll vs assigned hours"],
];

const c4cRows = [
  ["TicketID", "38925", "Ticket key"],
  ["WarrantyHandlingDealerID", "Perth dealer", "Dealer / yard mapping"],
  ["DealerName", "Regent RV - Perth", "Dealer / yard label"],
  ["TicketType", "Z007", "Repair / PDI code"],
  ["TicketTypeText", "Repair ticket", "Readable ticket type"],
  ["CreatedOn", "2026-07-08", "Created month / trend"],
  ["AmountIncludingTax", "$1,485", "Quote amount"],
  ["TicketStatusText", "Create invoice", "Workflow stage"],
  ["Role_1001_InvolvedPartyName", "REGENT RV PTY LTD", "Internal / external logic"],
  ["Role_40_InvolvedPartyName", "Missing / unreliable", "Worker if source is reliable"],
  ["TotalLabourHours / Z1Z8TimeConsumed", "Estimate only", "Labour estimate, not actual"],
  ["ResolvedOnDateTime / ChangeOnDateTime", "if populated", "Lifecycle timing"],
];

const sapRows = [
  ["Short text / ticket reference", "38925", "Join back to C4C ticket"],
  ["Bill-to name", "REGENT RV PTY LTD", "Customer / internal split cross-check"],
  ["ERPInvoiceNumber", "0090046551", "Invoice number"],
  ["Billing date", "2026-08-13", "Invoice month"],
  ["ERPInvoiceNumberPrice", "$1,350", "Invoice amount"],
  ["GR / delivery date", "if available", "PDI / delivery evidence"],
  ["Labour confirmation hours", "Missing", "Actual labour hours"],
  ["Assigned service worker", "Missing", "Technician / worker"],
  ["Payroll payable hours", "Missing", "Payroll comparison"],
  ["Purchase order / order no.", "if available", "SAP traceability"],
  ["Invoice status", "if available", "Billing validation"],
  ["Manual user / source", "if available", "Manual SAP exception checks"],
];

sheet.getRange("A6:C17").values = dashboardRows;
sheet.getRange("E6:G17").values = c4cRows;
sheet.getRange("I6:K17").values = sapRows;

sheet.getRange("A19:K19").merge();
sheet.getRange("A19").values = [["Quick read"]];
sheet.getRange("A20:K22").values = [
  ["Final dashboard fields are the output we want to show.", "", "", "", "C4C provides ticket identity, dealer, status, quote, and customer scope.", "", "", "", "SAP provides invoice, billing, labour/payroll, and operational confirmation fields.", "", ""],
  ["Keep each source grouped in its own block.", "", "", "", "Do not mix C4C field names into the SAP block.", "", "", "", "Fields that are missing should stay in the SAP block if SAP is expected to provide them.", "", ""],
  ["", "", "", "", "", "", "", "", "", "", ""],
];

const all = sheet.getRange("A1:K22");
all.format.font = { name: "Aptos", size: 10, color: "#111827" };
all.format.wrapText = true;
all.format.verticalAlignment = "center";

sheet.getRange("A1").format.font = { name: "Aptos Display", size: 16, bold: true, color: "#111827" };
sheet.getRange("A1").format.horizontalAlignment = "center";
sheet.getRange("A2").format.font = { italic: true, color: "#475569" };
sheet.getRange("A2").format.horizontalAlignment = "center";

sheet.getRange("A4:C4").format = {
  fill: "#111827",
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "medium", color: "#111827" },
};
sheet.getRange("E4:G4").format = {
  fill: "#2563EB",
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "medium", color: "#1D4ED8" },
};
sheet.getRange("I4:K4").format = {
  fill: "#16A34A",
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "medium", color: "#15803D" },
};

sheet.getRange("A5:C5").format = {
  fill: "#E5E7EB",
  font: { bold: true, color: "#111827" },
  borders: { preset: "all", style: "thin", color: "#CBD5E1" },
};
sheet.getRange("E5:G5").format = {
  fill: "#DBEAFE",
  font: { bold: true, color: "#1E3A8A" },
  borders: { preset: "all", style: "thin", color: "#93C5FD" },
};
sheet.getRange("I5:K5").format = {
  fill: "#DCFCE7",
  font: { bold: true, color: "#14532D" },
  borders: { preset: "all", style: "thin", color: "#86EFAC" },
};

sheet.getRange("A6:C17").format = {
  fill: "#F8FAFC",
  borders: { preset: "all", style: "thin", color: "#CBD5E1" },
};
sheet.getRange("E6:G17").format = {
  fill: "#EFF6FF",
  borders: { preset: "all", style: "thin", color: "#BFDBFE" },
};
sheet.getRange("I6:K17").format = {
  fill: "#F0FDF4",
  borders: { preset: "all", style: "thin", color: "#BBF7D0" },
};

sheet.getRange("A6:A17").format.font = { bold: true, color: "#111827" };
sheet.getRange("E6:E17").format.font = { bold: true, color: "#1D4ED8" };
sheet.getRange("I6:I17").format.font = { bold: true, color: "#15803D" };
sheet.getRange("A10:C10").format.fill = "#FEF3C7";
sheet.getRange("A15:C17").format.fill = "#FCE7F3";
sheet.getRange("I12:K17").format.fill = "#DCFCE7";

sheet.getRange("A19:K19").format = {
  fill: "#111827",
  font: { bold: true, color: "#FFFFFF" },
  borders: { preset: "outside", style: "medium", color: "#111827" },
};
sheet.getRange("A20:K22").format = {
  fill: "#F8FAFC",
  borders: { preset: "all", style: "thin", color: "#CBD5E1" },
};

sheet.getRange("A:K").format.horizontalAlignment = "left";
sheet.getRange("A4:K5").format.horizontalAlignment = "center";
sheet.getRange("A1:K2").format.horizontalAlignment = "center";

sheet.getRange("A:A").format.columnWidthPx = 190;
sheet.getRange("B:B").format.columnWidthPx = 185;
sheet.getRange("C:C").format.columnWidthPx = 230;
sheet.getRange("D:D").format.columnWidthPx = 28;
sheet.getRange("E:E").format.columnWidthPx = 230;
sheet.getRange("F:F").format.columnWidthPx = 180;
sheet.getRange("G:G").format.columnWidthPx = 230;
sheet.getRange("H:H").format.columnWidthPx = 28;
sheet.getRange("I:I").format.columnWidthPx = 230;
sheet.getRange("J:J").format.columnWidthPx = 180;
sheet.getRange("K:K").format.columnWidthPx = 230;

sheet.getRange("1:1").format.rowHeightPx = 34;
sheet.getRange("2:2").format.rowHeightPx = 24;
sheet.getRange("3:3").format.rowHeightPx = 12;
sheet.getRange("4:5").format.rowHeightPx = 34;
sheet.getRange("6:17").format.rowHeightPx = 42;
sheet.getRange("19:19").format.rowHeightPx = 30;
sheet.getRange("20:22").format.rowHeightPx = 46;

sheet.freezePanes.freezeRows(5);

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
  sheetName: sheet.name,
  range: "A1:K22",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/source_grouped_first_sheet.png`, new Uint8Array(await firstPreview.arrayBuffer()));

const second = workbook.worksheets.getItemAt(1);
const secondPreview = await workbook.render({
  sheetName: second.name,
  range: "A1:I18",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/source_grouped_second_sheet.png`, new Uint8Array(await secondPreview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
