import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/workflow_audit_data.json";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/01a01c4f-25b7-7e91-8ffd-41a61b3fc906";
const outputPath = path.join(outputDir, "workflow_audit_latest.xlsx");
const previewPath = path.join(outputDir, "workflow_audit_summary_preview.png");

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

function valuesFromRows(rows) {
  const headers = rows.length ? Object.keys(rows[0]) : ["No Data"];
  const values = rows.length ? rows.map((row) => headers.map((header) => row[header] ?? "")) : [[""]];
  return { headers, matrix: [headers, ...values] };
}

function writeTableSheet(workbook, name, rows, tableName) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  const { headers, matrix } = valuesFromRows(rows);
  const range = sheet.getRangeByIndexes(0, 0, matrix.length, headers.length);
  range.values = matrix;
  const lastCol = colLetter(headers.length);
  const header = sheet.getRange(`A1:${lastCol}1`);
  header.format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF" },
  };
  range.format.borders = {
    insideHorizontal: { style: "thin", color: "#D9E2F3" },
    top: { style: "thin", color: "#B4C6E7" },
    bottom: { style: "thin", color: "#B4C6E7" },
  };
  sheet.freezePanes.freezeRows(1);
  sheet.tables.add(`A1:${lastCol}${matrix.length}`, true, tableName);
  range.format.autofitColumns();
  range.format.autofitRows();
  return sheet;
}

const workbook = Workbook.create();

const currentPeriod = data.meta.dashboardLastUpdated.includes("2026") ? "Aug 2026" : "";
const spotlight = data.workflowSummary.filter(
  (row) =>
    row.Dealer === "Perth" &&
    ["Aug 2026", "2026"].includes(row.Period) &&
    ["Repair ticket", "PDI ticket"].includes(row.TicketType),
);

const summarySheet = workbook.worksheets.add("Summary");
summarySheet.showGridLines = false;
summarySheet.getRange("A1:F1").merge();
summarySheet.getRange("A1").values = [["Workflow Audit Export"]];
summarySheet.getRange("A1").format = {
  fill: "#17365D",
  font: { bold: true, color: "#FFFFFF", size: 16 },
};
summarySheet.getRange("A3:B11").values = [
  ["Source workbook", data.meta.sourceWorkbook],
  ["Dashboard data", data.meta.dashboardData],
  ["Dashboard last updated", data.meta.dashboardLastUpdated],
  ["Source rows", data.meta.sourceRows],
  ["Rows included by current Workflow base logic", data.meta.normalRowsIncludedInWorkflow],
  ["AbnormalParty rows excluded by current Workflow base logic", data.meta.abnormalPartyRowsExcluded],
  ["Workflow summary rows", data.meta.summaryRows],
  ["Current Workflow period rule", "Month = CreatedMonth; Year = CreatedYear"],
  ["Important note", "WorkflowSummary compares current dashboard-data.js against source recalculation."],
];
summarySheet.getRange("A3:A11").format = {
  fill: "#D9EAF7",
  font: { bold: true, color: "#17365D" },
};
summarySheet.getRange("A13:N13").values = [[
  "Period",
  "Dealer",
  "TicketType",
  "WorkflowStage",
  "DashboardQty",
  "RecalcQtyFromSource",
  "Difference",
  "QuoteAmount",
  "RawStatuses",
  "SourceRule",
  "DashboardTotalCreated",
  "RecalcTotalCreated",
  "DashboardOpen",
  "RecalcOpen",
]];
const spotlightMatrix = spotlight.map((row) => [
  row.Period,
  row.Dealer,
  row.TicketType,
  row.WorkflowStage,
  row.DashboardQty,
  row.RecalcQtyFromSource,
  row.Difference,
  row.QuoteAmount,
  row.RawStatuses,
  row.SourceRule,
  row.DashboardTotalCreated,
  row.RecalcTotalCreated,
  row.DashboardOpen,
  row.RecalcOpen,
]);
if (spotlightMatrix.length) {
  summarySheet.getRangeByIndexes(13, 0, spotlightMatrix.length, 14).values = spotlightMatrix;
}
summarySheet.getRange("A13:N13").format = {
  fill: "#1F4E79",
  font: { bold: true, color: "#FFFFFF" },
};
summarySheet.getRange(`A13:N${13 + spotlightMatrix.length}`).format.borders = {
  insideHorizontal: { style: "thin", color: "#D9E2F3" },
  top: { style: "thin", color: "#B4C6E7" },
  bottom: { style: "thin", color: "#B4C6E7" },
};
summarySheet.getRange("D:D").format.columnWidth = 34;
summarySheet.getRange("I:J").format.columnWidth = 42;
summarySheet.freezePanes.freezeRows(13);
summarySheet.getRange("A:N").format.autofitColumns();

const workflowSheet = writeTableSheet(workbook, "WorkflowSummary", data.workflowSummary, "WorkflowSummaryTable");
workflowSheet.getRange("H:H").format.numberFormat = "$#,##0.00";
workflowSheet.getRange("J:J").format.columnWidth = 44;

const detailSheet = writeTableSheet(workbook, "WorkflowDetail", data.workflowDetail, "WorkflowDetailTable");
detailSheet.getRange("Q:R").format.numberFormat = "$#,##0.00";
detailSheet.getRange("Z:AA").format.numberFormat = "#,##0.0";
detailSheet.getRange("AD:AE").format.columnWidth = 48;

writeTableSheet(workbook, "StatusMapping", data.statusMapping, "StatusMappingTable");

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errorScan.ndjson);

const preview = await workbook.render({
  sheetName: "Summary",
  range: "A1:N32",
  scale: 1,
  format: "png",
});

await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath }, null, 2));
