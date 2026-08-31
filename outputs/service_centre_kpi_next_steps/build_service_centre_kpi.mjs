import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/service_centre_kpi_next_steps";
const outputPath = `${outputDir}/service_centre_kpi_next_steps.xlsx`;
const previewPath = `${outputDir}/service_centre_kpi_next_steps_preview.png`;

const functions = [
  {
    name: "Technician Utilisation",
    formula: "Productive / Available Hours",
    field: "Missing\nProductive hours",
    target: ">90%",
    nextStep: "Add a new field\nProductive hours",
    stage: "",
    owner: "C4C",
    dueDate: "",
  },
  {
    name: "Service Orders Trend",
    formula: "Status Change Date (Missing)",
    field: "Missing\nStatus Change Date",
    target: "Month To Month",
    nextStep: "Add a new field\nStatus Change Date",
    stage: "",
    owner: "C4C",
    dueDate: "",
  },
  {
    name: "Comeback / Rework Rate",
    formula: "Rework Tickets / Completed Repair Orders",
    field: "Rework can tell by same chassis number and description but looking forward a mark in system",
    target: "<5%",
    nextStep: "Add a new field to mark rework.",
    stage: "",
    owner: "C4C",
    dueDate: "",
  },
  {
    name: "Fixed First Time Right",
    formula: "1 - Rework Rate",
    field: "Once Rework Rate is done",
    target: ">95%",
    nextStep: "Use rework tickets as the failure signal.",
    stage: "",
    owner: "C4C",
    dueDate: "",
  },
  {
    name: "Quotation Approval Rate",
    formula: "Approved Quotes / Raised Quotes",
    field: "Missing quote rejected, having quote approved from C4C using Repair Total Amount",
    target: ">70%",
    nextStep: "Add a new field for rejection in system instead of sharing with cancel.\nProvide Repair Total Amount Trigger",
    stage: "",
    owner: "C4C",
    dueDate: "",
  },
  {
    name: "Service Appointment Lead Time",
    formula: "Appointment Date - Booking Request Date",
    field: "Missing appointment date, having CreatedOn as Booking Request Date",
    target: "3 to 7 days",
    nextStep: "Add a new field for Appointment Date in system.",
    stage: "",
    owner: "C4C",
    dueDate: "",
  },
  {
    name: "Customer Retention Rate",
    formula: "Returning Sold Vans / Sold Vans in Service Scope",
    field: "Chassis number in C4C",
    target: ">90%",
    nextStep: "Define customer return logic using sold van / local service records in system. Chassis? Rego?",
    stage: "",
    owner: "",
    dueDate: "",
  },
  {
    name: "Retail Repair Order Growth Rate",
    formula: "(Current Period Retail ROs - Prior Period Retail ROs) / Prior Period Retail ROs",
    field: "RO ID in C4C, but missing status change time by date",
    target: "3% to 6% YOY",
    nextStep: "Add a new field to record the status last change time.",
    stage: "",
    owner: "C4C",
    dueDate: "",
  },
  {
    name: "Average Repair Order Amount",
    formula: "Total Repair Order Revenue / Number of Repair Orders",
    field: "Total Repair Order Revenue: ?\nNumber of Repair Orders: C4C OrderID sum up",
    target: "Benchmark: confirm with Finance",
    nextStep: "",
    stage: "",
    owner: "",
    dueDate: "",
  },
  {
    name: "Effective Labor Rate",
    formula: "Total Labor Revenue / Labor Hours Sold",
    field: "Total Labour Revenue: __?__\nLabor Hours Sold: ClaimHour from C4C (TotalHour)",
    target: ">90% of posted rate",
    nextStep: "Confirm labour revenue and labour hours sold source.",
    stage: "",
    owner: "",
    dueDate: "",
  },
  {
    name: "Gross Profit Margin",
    formula: "(Parts Revenue + Labor Revenue - Parts Cost - Labor Cost) / Revenue",
    field: "",
    target: "58% to 65%",
    nextStep: "",
    stage: "",
    owner: "",
    dueDate: "",
  },
  {
    name: "Service Gross %",
    formula: "(Labor Sales - Direct Labor Cost) / Labor Sales",
    field: "",
    target: "65% to 75%",
    nextStep: "",
    stage: "",
    owner: "",
    dueDate: "",
  },
  {
    name: "Parts Gross %",
    formula: "(Parts Sales - Parts Cost) / Parts Sales",
    field: "",
    target: "35% to 45%",
    nextStep: "Confirm parts sales and parts cost fields.",
    stage: "",
    owner: "",
    dueDate: "",
  },
  {
    name: "Absorption Rate",
    formula: "(Service Gross Profit + Parts Gross Profit) / Operating Expenses",
    field: "",
    target: "100% to 130%",
    nextStep: "Confirm service operating expense allocation and cost centre split.",
    stage: "",
    owner: "",
    dueDate: "",
  },
];

const rowLabels = [
  ["Formula", "formula"],
  ["Field", "field"],
  ["Target", "target"],
  ["Next Step", "nextStep"],
  ["Stage", "stage"],
  ["Owner", "owner"],
  ["Due Date", "dueDate"],
];

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("KPI Next Steps");
sheet.showGridLines = false;

sheet.getRange("A1:H1").merge();
sheet.getRange("A1").values = [["Service Centre KPI Next Steps"]];

const matrix = [
  ["Function", ...rowLabels.map(([label]) => label)],
  ...functions.map((item) => [
    item.name,
    item.formula,
    item.field,
    item.target,
    item.nextStep,
    item.stage,
    item.owner,
    item.dueDate,
  ]),
];

sheet.getRangeByIndexes(1, 0, matrix.length, matrix[0].length).values = matrix;

sheet.freezePanes.freezeRows(2);
sheet.freezePanes.freezeColumns(1);

sheet.getRange("A1:H1").format = {
  fill: "#174E63",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  horizontalAlignment: "left",
  verticalAlignment: "middle",
};

sheet.getRange("A2:H2").format = {
  fill: "#256B7C",
  font: { bold: true, color: "#FFFFFF", size: 10 },
  horizontalAlignment: "center",
  verticalAlignment: "middle",
  wrapText: true,
};

sheet.getRange("A3:A16").format = {
  fill: "#E8F0F3",
  font: { bold: true, color: "#16313B" },
  horizontalAlignment: "left",
  verticalAlignment: "middle",
};

sheet.getRange("B3:H16").format = {
  fill: "#FFFFFF",
  font: { color: "#1F2933", size: 9 },
  horizontalAlignment: "left",
  verticalAlignment: "top",
  wrapText: true,
};

sheet.getRange("A2:H16").format.borders = {
  insideHorizontal: { style: "thin", color: "#CFD8DC" },
  insideVertical: { style: "thin", color: "#E1E7EA" },
  top: { style: "thin", color: "#B7C5CB" },
  bottom: { style: "thin", color: "#B7C5CB" },
  left: { style: "thin", color: "#B7C5CB" },
  right: { style: "thin", color: "#B7C5CB" },
};

sheet.getRange("A1").format.rowHeight = 28;
sheet.getRange("A2:H2").format.rowHeight = 32;
sheet.getRange("A3:H16").format.rowHeight = 58;

sheet.getRange("A:A").format.columnWidth = 28;
sheet.getRange("B:B").format.columnWidth = 36;
sheet.getRange("C:C").format.columnWidth = 42;
sheet.getRange("D:D").format.columnWidth = 18;
sheet.getRange("E:E").format.columnWidth = 44;
sheet.getRange("F:H").format.columnWidth = 16;
sheet.getRange("A1:H16").format.font = { name: "Aptos" };

const check = await workbook.inspect({
  kind: "table",
  range: "KPI Next Steps!A1:H16",
  include: "values,formulas",
  tableMaxRows: 16,
  tableMaxCols: 8,
  tableMaxCellChars: 120,
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
  sheetName: "KPI Next Steps",
  range: "A1:H16",
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

console.log(`Saved ${outputPath}`);
console.log(`Preview ${previewPath}`);
