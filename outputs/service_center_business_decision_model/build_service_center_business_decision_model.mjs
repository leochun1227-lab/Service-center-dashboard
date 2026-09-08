import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/service_center_business_decision_model";
const outputPath = `${outputDir}/service_center_business_category_decision_model.xlsx`;
const previewPath = `${outputDir}/service_center_business_category_decision_model_preview.png`;

const cases = [
  {
    category: "Customer Self-Serviced",
    ticketId: "SO-CUST-001",
    chassis: "SRV260101",
    invoiceNo: "0090046501",
    ticketType: "Z007",
    ticketTypeText: "Repair ticket",
    status: "Create invoice",
    customer: "Janelle Sheahan",
    dealer: "Geelong",
    warrantyDealer: "3128",
    sapBillTo: "Janelle Sheahan",
    sapPayerType: "Customer",
    sapDept: "",
    invoiceAmount: 1350,
    billingDate: "2026-09-03",
    expectedChargeTo: "Customer",
    actualChargeTo: "Customer",
    ruleId: "R010",
    businessCategory: "Customer Self-Serviced",
    scenario: "Customer paid service / repair",
    statusResult: "OK",
    countImpact: "+1 income",
    month: "Current Month",
    exceptionReason: "",
  },
  {
    category: "",
    ticketId: "SO-CUST-002",
    chassis: "SRV260102",
    invoiceNo: "",
    ticketType: "Z007",
    ticketTypeText: "Repair ticket",
    status: "Open",
    customer: "Ben Hill",
    dealer: "Perth",
    warrantyDealer: "3121",
    sapBillTo: "",
    sapPayerType: "",
    sapDept: "",
    invoiceAmount: "",
    billingDate: "",
    expectedChargeTo: "Customer",
    actualChargeTo: "",
    ruleId: "R060",
    businessCategory: "Customer Self-Serviced",
    scenario: "Expected customer-paid open repair",
    statusResult: "Pending invoice",
    countImpact: "+1 open order",
    month: "Created Month",
    exceptionReason: "",
  },
  {
    category: "",
    ticketId: "SO-CUST-003",
    chassis: "SRV260103",
    invoiceNo: "0090046503",
    ticketType: "Z007",
    ticketTypeText: "Repair ticket",
    status: "Cancel",
    customer: "David Waters",
    dealer: "Frankston",
    warrantyDealer: "3141",
    sapBillTo: "David Waters",
    sapPayerType: "Customer",
    sapDept: "",
    invoiceAmount: -420,
    billingDate: "2026-09-05",
    expectedChargeTo: "Customer",
    actualChargeTo: "Customer",
    ruleId: "R011",
    businessCategory: "Customer Self-Serviced",
    scenario: "Customer invoice cancelled / reversal",
    statusResult: "Review",
    countImpact: "-1 income",
    month: "Billing Month",
    exceptionReason: "Cancelled ticket has billing activity",
  },
  {
    category: "Internal",
    ticketId: "SO-INT-001",
    chassis: "SRV260201",
    invoiceNo: "0090046601",
    ticketType: "Z010",
    ticketTypeText: "PDI",
    status: "Claim time ticket",
    customer: "REGENT RV PTY LTD",
    dealer: "Perth",
    warrantyDealer: "3121",
    sapBillTo: "REGENT RV PTY LTD",
    sapPayerType: "Department",
    sapDept: "PDI Department",
    invoiceAmount: 0,
    billingDate: "2026-09-04",
    expectedChargeTo: "PDI Department",
    actualChargeTo: "PDI Department",
    ruleId: "R030",
    businessCategory: "Internal",
    scenario: "Inbound / outbound PDI",
    statusResult: "OK",
    countImpact: "+1 internal payment",
    month: "Billing Month",
    exceptionReason: "",
  },
  {
    category: "",
    ticketId: "SO-INT-002",
    chassis: "SRV260202",
    invoiceNo: "0090046602",
    ticketType: "Z007",
    ticketTypeText: "Repair ticket",
    status: "Create invoice",
    customer: "REGENT RV - PERTH",
    dealer: "Perth",
    warrantyDealer: "3121",
    sapBillTo: "Sales Department",
    sapPayerType: "Department",
    sapDept: "Sales Department",
    invoiceAmount: 680,
    billingDate: "2026-09-07",
    expectedChargeTo: "Sales Department",
    actualChargeTo: "Sales Department",
    ruleId: "R031",
    businessCategory: "Internal",
    scenario: "Dealer fitted accessory / option",
    statusResult: "OK",
    countImpact: "+1 internal payment",
    month: "Billing Month",
    exceptionReason: "",
  },
  {
    category: "",
    ticketId: "SO-INT-003",
    chassis: "SRV260203",
    invoiceNo: "",
    ticketType: "Z010",
    ticketTypeText: "PDI",
    status: "Open",
    customer: "REGENT RV PTY LTD",
    dealer: "Traralgon",
    warrantyDealer: "3123",
    sapBillTo: "",
    sapPayerType: "",
    sapDept: "",
    invoiceAmount: "",
    billingDate: "",
    expectedChargeTo: "PDI Department",
    actualChargeTo: "",
    ruleId: "R050",
    businessCategory: "Internal",
    scenario: "Expected internal PDI, no SAP invoice yet",
    statusResult: "Pending invoice",
    countImpact: "+1 open internal",
    month: "Created Month",
    exceptionReason: "",
  },
  {
    category: "Warranty",
    ticketId: "SO-WAR-001",
    chassis: "SRV260301",
    invoiceNo: "0090046701",
    ticketType: "Z007",
    ticketTypeText: "Repair ticket",
    status: "Create invoice",
    customer: "Robert Otto",
    dealer: "Geelong",
    warrantyDealer: "3128",
    sapBillTo: "Warranty Department",
    sapPayerType: "Department",
    sapDept: "Warranty Department",
    invoiceAmount: 980,
    billingDate: "2026-09-08",
    expectedChargeTo: "Warranty Department",
    actualChargeTo: "Warranty Department",
    ruleId: "R020",
    businessCategory: "Warranty",
    scenario: "Warranty claim paid by warranty department",
    statusResult: "OK",
    countImpact: "+1 warranty payment",
    month: "Billing Month",
    exceptionReason: "",
  },
  {
    category: "",
    ticketId: "SO-WAR-002",
    chassis: "SRV260302",
    invoiceNo: "",
    ticketType: "Z007",
    ticketTypeText: "Warranty repair",
    status: "Repair completed",
    customer: "Linda McLellan",
    dealer: "Frankston",
    warrantyDealer: "3141",
    sapBillTo: "",
    sapPayerType: "",
    sapDept: "",
    invoiceAmount: "",
    billingDate: "",
    expectedChargeTo: "Warranty Department",
    actualChargeTo: "",
    ruleId: "R040",
    businessCategory: "Warranty",
    scenario: "Expected warranty, no SAP payment yet",
    statusResult: "Pending invoice",
    countImpact: "+1 open warranty",
    month: "Created Month",
    exceptionReason: "",
  },
  {
    category: "",
    ticketId: "SO-WAR-003",
    chassis: "SRV260303",
    invoiceNo: "0090046703",
    ticketType: "Z007",
    ticketTypeText: "Warranty repair",
    status: "Create invoice",
    customer: "REGENT RV - PERTH",
    dealer: "Perth",
    warrantyDealer: "3121",
    sapBillTo: "Customer",
    sapPayerType: "Customer",
    sapDept: "",
    invoiceAmount: 760,
    billingDate: "2026-09-09",
    expectedChargeTo: "Warranty Department",
    actualChargeTo: "Customer",
    ruleId: "R900",
    businessCategory: "Warranty",
    scenario: "Warranty expectation but customer billed",
    statusResult: "Exception",
    countImpact: "Hold from KPI",
    month: "Billing Month",
    exceptionReason: "Expected and actual charge-to conflict",
  },
];

const ruleRows = [
  [10, "SAP payer type = Customer and invoice exists", "Customer Self-Serviced", "Customer", "Customer paid service / repair", "SAP billing", "Final actual"],
  [20, "SAP department = Warranty Department and invoice exists", "Warranty", "Warranty Department", "Warranty claim paid", "SAP billing", "Final actual"],
  [30, "SAP department in Sales Department / PDI Department and invoice exists", "Internal", "Sales or PDI Department", "Internal work paid by department", "SAP billing", "Final actual"],
  [40, "No SAP invoice and C4C ticket has warranty type / warranty flag", "Warranty", "Warranty Department", "Expected warranty", "C4C intent", "Expected"],
  [50, "No SAP invoice and C4C ticket is PDI / inbound / outbound / accessory / option", "Internal", "Sales or PDI Department", "Expected internal", "C4C intent", "Expected"],
  [60, "No SAP invoice and C4C repair/service has real customer", "Customer Self-Serviced", "Customer", "Expected customer paid repair", "C4C intent", "Expected"],
  [900, "ExpectedChargeTo <> ActualChargeTo", "Keep expected category, flag exception", "Review", "Mismatch between systems", "Cross-system check", "Exception"],
  [999, "No rule matched", "Unknown", "Unknown", "Need manual review", "Fallback", "Exception"],
];

await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Normal Cases");
sheet.showGridLines = false;

const headers = [
  "",
  "TicketID",
  "Chassis",
  "Invoice No.",
  "Ticket Type",
  "Ticket Type Text",
  "Status",
  "Customer",
  "Dealer",
  "Warranty Dealer ID",
  "Bill-to / Payer",
  "Payer Type",
  "Department",
  "Invoice Amount",
  "Billing Date",
  "Expected Charge To",
  "Actual Charge To",
  "Rule ID",
  "Business Category",
  "Scenario",
  "Status",
  "Count / Amount",
  "Month",
  "Exception Reason",
];

sheet.getRange("B1:D1").merge();
sheet.getRange("E1:J1").merge();
sheet.getRange("K1:O1").merge();
sheet.getRange("P1:Q1").merge();
sheet.getRange("R1:R1").merge();
sheet.getRange("S1:X1").merge();
sheet.getRange("B1:X1").values = [[
  "Main Key", "", "",
  "C4C Ticket", "", "", "", "", "",
  "SAP Finance", "", "", "", "",
  "Decision Input", "",
  "????",
  "Result", "", "", "", "", "",
]];
sheet.getRange("A2:X2").values = [headers];

const blankRow = Array(headers.length).fill("");
const caseValues = cases.flatMap((row, index) => {
  const values = [
    row.category,
    row.ticketId,
    row.chassis,
    row.invoiceNo,
    row.ticketType,
    row.ticketTypeText,
    row.status,
    row.customer,
    row.dealer,
    row.warrantyDealer,
    row.sapBillTo,
    row.sapPayerType,
    row.sapDept,
    row.invoiceAmount,
    row.billingDate,
    row.expectedChargeTo,
    row.actualChargeTo,
    row.ruleId,
    row.businessCategory,
    row.scenario,
    row.statusResult,
    row.countImpact,
    row.month,
    row.exceptionReason,
  ];
  return index === 2 || index === 5 ? [values, blankRow] : [values];
});
sheet.getRange(`A3:X${caseValues.length + 2}`).values = caseValues;

const ruleSheet = workbook.worksheets.add("Rule Priority");
ruleSheet.showGridLines = false;
ruleSheet.getRange("A1:G1").merge();
ruleSheet.getRange("A1").values = [["Service Center Unique Business Classification Rules"]];
ruleSheet.getRange("A2:G2").merge();
ruleSheet.getRange("A2").values = [["Apply rules from lowest Priority to highest. The first matched rule owns the final Business Category; mismatches are exported for review."]];
ruleSheet.getRange("A4:G4").values = [["Priority", "Condition", "BusinessCategory", "ChargeTo", "Scenario", "Source", "Output Type"]];
ruleSheet.getRange(`A5:G${ruleRows.length + 4}`).values = ruleRows;

function applyBaseStyle(targetSheet, range) {
  targetSheet.getRange(range).format = {
    font: { name: "Aptos", size: 10, color: "#17324D" },
    wrapText: true,
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#D9E2EA" },
  };
}

applyBaseStyle(sheet, "A1:X13");
sheet.getRange("A1:X1").format = {
  fill: "#D9EAF7",
  font: { name: "Aptos", size: 10, color: "#17324D" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sheet.getRange("R1").format.fill = "#FFFF00";
sheet.getRange("B1:D2").format.font = { name: "Aptos", size: 10, color: "#FF0000", bold: true };
sheet.getRange("A2:X2").format.font = { name: "Aptos", size: 10, bold: true, color: "#17324D" };
sheet.getRange("A2:X2").format.horizontalAlignment = "right";
sheet.getRange("A3:A13").format.horizontalAlignment = "center";
sheet.getRange("A3:A13").format.font = { name: "Aptos", size: 10, bold: true, color: "#17324D" };
sheet.getRange("A6:X6").format.fill = "#FFFFFF";
sheet.getRange("A10:X10").format.fill = "#FFFFFF";
sheet.getRange("T3:T13").format.columnWidthPx = 175;
sheet.getRange("X3:X13").format.columnWidthPx = 210;
sheet.getRange("N3:N13").format.numberFormat = "$#,##0";
sheet.getRange("O3:O13").format.numberFormat = "yyyy-mm-dd";
sheet.getRange("R3:R13").format.horizontalAlignment = "center";
sheet.getRange("U3:U13").format.horizontalAlignment = "center";

const columnWidths = {
  A: 155, B: 112, C: 112, D: 112, E: 90, F: 132, G: 126, H: 160,
  I: 105, J: 126, K: 155, L: 105, M: 145, N: 105, O: 105, P: 160,
  Q: 150, R: 72, S: 150, T: 195, U: 92, V: 130, W: 120, X: 220,
};
for (const [col, width] of Object.entries(columnWidths)) {
  sheet.getRange(`${col}:${col}`).format.columnWidthPx = width;
}
sheet.getRange("1:1").format.rowHeightPx = 30;
sheet.getRange("2:2").format.rowHeightPx = 42;
sheet.getRange("3:13").format.rowHeightPx = 46;
sheet.getRange("6:6").format.rowHeightPx = 16;
sheet.getRange("10:10").format.rowHeightPx = 16;
sheet.freezePanes.freezeRows(2);

applyBaseStyle(ruleSheet, "A1:G12");
ruleSheet.getRange("A1:G1").format = {
  fill: "#D9EAF7",
  font: { name: "Aptos", size: 12, bold: true, color: "#17324D" },
  horizontalAlignment: "center",
};
ruleSheet.getRange("A2:G2").format = {
  fill: "#FFFFFF",
  font: { name: "Aptos", size: 10, italic: true, color: "#17324D" },
  horizontalAlignment: "center",
};
ruleSheet.getRange("A4:G4").format = {
  fill: "#D9EAF7",
  font: { name: "Aptos", size: 10, bold: true, color: "#17324D" },
  horizontalAlignment: "right",
};
ruleSheet.getRange("A:A").format.columnWidthPx = 82;
ruleSheet.getRange("B:B").format.columnWidthPx = 310;
ruleSheet.getRange("C:C").format.columnWidthPx = 170;
ruleSheet.getRange("D:D").format.columnWidthPx = 165;
ruleSheet.getRange("E:E").format.columnWidthPx = 210;
ruleSheet.getRange("F:F").format.columnWidthPx = 125;
ruleSheet.getRange("G:G").format.columnWidthPx = 120;
ruleSheet.getRange("1:1").format.rowHeightPx = 30;
ruleSheet.getRange("2:2").format.rowHeightPx = 38;
ruleSheet.getRange("4:12").format.rowHeightPx = 38;
ruleSheet.freezePanes.freezeRows(4);

const inspectNormal = await workbook.inspect({
  kind: "region",
  sheetId: "Normal Cases",
  range: "A1:X13",
  maxChars: 8000,
});
console.log(inspectNormal.ndjson);

const inspectRules = await workbook.inspect({
  kind: "region",
  sheetId: "Rule Priority",
  range: "A1:G12",
  maxChars: 5000,
});
console.log(inspectRules.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Normal Cases",
  range: "A1:X13",
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const rulesPreview = await workbook.render({
  sheetName: "Rule Priority",
  range: "A1:G12",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/rule_priority_preview.png`, new Uint8Array(await rulesPreview.arrayBuffer()));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath }, null, 2));
