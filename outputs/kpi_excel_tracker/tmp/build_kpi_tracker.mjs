import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/kpi_excel_tracker";
const outputPath = `${outputDir}/service_centre_kpi_tracker.xlsx`;

const headers = [
  "Category",
  "KPI",
  "Target / Benchmark",
  "Formula / Definition",
  "Source / System",
  "Field in System",
  "Next Step",
  "Priority",
  "Alternative",
  "Due date",
  "Responsible by",
  "Status",
  "Notes"
];

const rows = [
  {
    category: "Financial",
    kpi: "Average Repair Order Amount",
    target: "",
    formula: "Total Repair Order Revenue / Number of Repair Orders",
    source: "SAP Invoice + C4C",
    field: "",
    next: "Confirm SAP invoice revenue field and match it with C4C repair order count.",
    priority: "High",
    alt: "Use C4C quote / invoice amount as a temporary proxy.",
    due: "",
    owner: "Finance / SAP + C4C owner",
    status: "TBC",
    notes: "Repair Order Revenue, Labor Revenue and Parts Revenue should be confirmed before dashboard build."
  },
  {
    category: "Financial",
    kpi: "Effective Labor Rate",
    target: ">90% of posted rate",
    formula: "Total Labor Revenue / Labor Hours Sold",
    source: "SAP Invoice + C4C Labour",
    field: "",
    next: "Confirm labour revenue and labour hours sold source.",
    priority: "High",
    alt: "Use claimed hours first if sold hours are not available.",
    due: "",
    owner: "Finance / Service Managers",
    status: "TBC",
    notes: "People ID / technician-level labour linkage still needs confirmation."
  },
  {
    category: "Financial",
    kpi: "Gross Profit Margin",
    target: "58% to 65%",
    formula: "(Parts Revenue + Labor Revenue - Parts Cost - Labor Cost) / (Parts Revenue + Labor Revenue)",
    source: "SAP Invoice + SAP PO + HR / Finance",
    field: "",
    next: "Confirm revenue, cost and expense allocation rules.",
    priority: "High",
    alt: "Report gross margin before full operating expense allocation.",
    due: "",
    owner: "Finance / Mary",
    status: "TBC",
    notes: "Expense allocation may be phase 2 if cost-centre split is not ready."
  },
  {
    category: "Financial",
    kpi: "Service Gross %",
    target: "65% to 75%",
    formula: "(Labor Sales - Direct Labor Cost) / Labor Sales",
    source: "SAP Invoice + HR / Payroll",
    field: "",
    next: "Confirm labour sales and direct labour cost source.",
    priority: "High",
    alt: "Show overview-only service gross until technician people ID is available.",
    due: "",
    owner: "Finance / HR",
    status: "TBC",
    notes: "Person-level detail depends on technician / people ID mapping."
  },
  {
    category: "Financial",
    kpi: "Parts Gross %",
    target: "35% to 45%",
    formula: "(Parts Sales - Parts Cost) / Parts Sales",
    source: "SAP SO Invoice + SAP PO Invoice",
    field: "",
    next: "Confirm parts sales and parts cost fields.",
    priority: "Medium",
    alt: "Use summary parts revenue / cost until line-level mapping is stable.",
    due: "",
    owner: "Finance / SAP owner",
    status: "TBC",
    notes: ""
  },
  {
    category: "Financial",
    kpi: "Absorption Rate",
    target: "100% to 130%",
    formula: "(Service Gross Profit + Parts Gross Profit) / Operating Expenses",
    source: "Finance P&L / SAP",
    field: "",
    next: "Confirm service operating expense allocation and cost centre split.",
    priority: "Medium",
    alt: "Track after-service gross profit against estimated operating cost.",
    due: "",
    owner: "Finance",
    status: "TBC",
    notes: "Current notes say dealers only use a single cost centre; full split may be phase 2."
  },
  {
    category: "Productivity and Efficiency",
    kpi: "Technician Productivity",
    target: "110% to 130%",
    formula: "Sold / Claimed Hours / Available Hours",
    source: "C4C Labour + Technician roster",
    field: "",
    next: "Set up technician list, team and available hours.",
    priority: "High",
    alt: "Use location-level claimed hours until technician-level setup is complete.",
    due: "",
    owner: "Service Managers / Andy He",
    status: "TBC",
    notes: "Apprentice efficiency benchmarks can be added as target dimensions."
  },
  {
    category: "Productivity and Efficiency",
    kpi: "Technician Efficiency",
    target: ">110%",
    formula: "Sold / Claimed Hours / Actual Productive Hours",
    source: "C4C Work Orders + Labour records",
    field: "",
    next: "Assign every work order to a specific technician and record actual work hours.",
    priority: "High",
    alt: "Use claim hours by location if actual productive hours are not stable.",
    due: "",
    owner: "Service Managers",
    status: "TBC",
    notes: ""
  },
  {
    category: "Productivity and Efficiency",
    kpi: "Technician Utilisation",
    target: ">90%",
    formula: "Productive / Claimed Hours / Available Hours",
    source: "C4C Labour + Technician calendar",
    field: "",
    next: "Confirm roster / calendar available hours and productive hour logic.",
    priority: "High",
    alt: "Use claimed hours / working-day capacity until calendar data is ready.",
    due: "",
    owner: "Service Managers / C4C owner",
    status: "TBC",
    notes: "Calendar solution is needed for more accurate available capacity."
  },
  {
    category: "Operations & Customer",
    kpi: "Comeback Rate / Rework Rate",
    target: "<5%",
    formula: "Rework Tickets / Completed Repair Orders",
    source: "C4C Tickets",
    field: "",
    next: "Rename comeback to rework and define rework ticket creation rule.",
    priority: "High",
    alt: "Detect repeat repair by same vehicle / same issue as a temporary proxy.",
    due: "",
    owner: "Service Managers / Lee",
    status: "TBC",
    notes: "Notes recommend creating a separate rework order for repeat repair."
  },
  {
    category: "Operations & Customer",
    kpi: "Fixed First Time Right",
    target: ">95%",
    formula: "1 - Rework Rate",
    source: "C4C Tickets",
    field: "",
    next: "Use rework tickets as the failure signal.",
    priority: "High",
    alt: "Use same-issue repeat repair detection until rework ticket rule is live.",
    due: "",
    owner: "Dashboard owner / Service Managers",
    status: "TBC",
    notes: ""
  },
  {
    category: "Operations & Customer",
    kpi: "Quotation Approval Rate",
    target: ">70%",
    formula: "Approved Quotes / Raised Quotes",
    source: "C4C Quote / Repair Orders",
    field: "",
    next: "Standardise quote statuses into raised, approved and rejected.",
    priority: "High",
    alt: "Use open repair order quote status as interim measure.",
    due: "",
    owner: "C4C owner / Service Managers",
    status: "TBC",
    notes: "The notes say quote tracking can use standard line item status."
  },
  {
    category: "Operations & Customer",
    kpi: "Customer Retention Rate",
    target: ">90%",
    formula: "Returning Sold Vans / Sold Vans in Service Scope",
    source: "C4C Service + Sales / Customer data",
    field: "",
    next: "Define customer return logic using sold van / local service records in C4C.",
    priority: "Medium",
    alt: "Use return service rate for sold vans as a first proxy.",
    due: "",
    owner: "Business / Service",
    status: "TBC",
    notes: "Tie this to growing car park once sales/service linkage is confirmed."
  },
  {
    category: "Operations & Customer",
    kpi: "Retail Repair Order Growth Rate",
    target: "3% to 6% YOY",
    formula: "(Current Period Retail ROs - Prior Period Retail ROs) / Prior Period Retail ROs",
    source: "C4C Repair Orders",
    field: "",
    next: "Use month-on-month or quarterly trend before full YOY baseline is available.",
    priority: "Medium",
    alt: "Use MoM / QoQ trend until 2027 YOY comparison is reliable.",
    due: "",
    owner: "BI / Dashboard owner",
    status: "TBC",
    notes: "Notes say YoY growth becomes available from 2027."
  },
  {
    category: "Operations & Customer",
    kpi: "Service Appointment Lead Time",
    target: "3 to 7 days",
    formula: "Appointment Date - Booking Request Date",
    source: "C4C Calendar / Tickets",
    field: "",
    next: "Confirm requested service date capture; short term use ticket created date as proxy.",
    priority: "High",
    alt: "Ticket created date to appointment calendar date.",
    due: "",
    owner: "Andy He / Service Managers",
    status: "TBC",
    notes: "Current notes say requested service date is not recorded consistently."
  }
];

const dictionary = [
  ["Column", "Purpose"],
  ["Category", "KPI grouping used for dashboard / sheet filters."],
  ["KPI", "Metric name."],
  ["Target / Benchmark", "Target threshold or expected range."],
  ["Formula / Definition", "Business formula, not final system formula."],
  ["Source / System", "Likely system source such as SAP, C4C, Finance P&L, HR."],
  ["Field in System", "Exact field name to confirm and fill later."],
  ["Next Step", "Action needed before dashboard automation."],
  ["Priority", "High / Medium / Low."],
  ["Alternative", "Fallback metric if exact source is not available."],
  ["Due date", "Owner-entered target date."],
  ["Responsible by", "Person or team accountable for confirmation."],
  ["Status", "TBC / Not Started / In Progress / Confirmed / Blocked."],
  ["Notes", "Open assumptions or implementation comments."]
];

function matrix(data) {
  return data.map((r) => [
    r.category,
    r.kpi,
    r.target,
    r.formula,
    r.source,
    r.field,
    r.next,
    r.priority,
    r.alt,
    r.due,
    r.owner,
    r.status,
    r.notes
  ]);
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - m) / 26);
  }
  return s;
}

function setupSheet(sheet, title, data) {
  const width = headers.length;
  const lastCol = colLetter(width);
  sheet.showGridLines = false;
  sheet.getRange(`A1:${lastCol}1`).merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: "#0F766E",
    font: { bold: true, color: "#FFFFFF", size: 16 },
  };
  sheet.getRange("A2").values = [[`Updated: ${new Date().toISOString().slice(0, 10)} | editable KPI implementation tracker`]];
  sheet.getRange(`A2:${lastCol}2`).merge();
  sheet.getRange("A2").format = {
    fill: "#ECFDF5",
    font: { color: "#065F46", italic: true },
  };
  sheet.getRangeByIndexes(3, 0, 1, width).values = [headers];
  sheet.getRangeByIndexes(4, 0, data.length, width).values = matrix(data);
  const tableRange = `A4:${lastCol}${4 + data.length}`;
  const table = sheet.tables.add(tableRange, true, `${title.replace(/[^A-Za-z0-9]/g, "").slice(0, 20)}Table`);
  table.style = "TableStyleMedium2";
  sheet.freezePanes.freezeRows(4);
  sheet.getRange(tableRange).format.wrapText = true;
  sheet.getRange(`A4:${lastCol}4`).format = {
    fill: "#134E4A",
    font: { bold: true, color: "#FFFFFF" },
  };
  sheet.getRange(`A5:${lastCol}${4 + data.length}`).format = {
    font: { color: "#111827" },
    borders: {
      insideHorizontal: { style: "thin", color: "#E5E7EB" },
      insideVertical: { style: "thin", color: "#E5E7EB" },
    },
  };
  const widths = [24, 30, 18, 58, 24, 28, 48, 14, 46, 16, 28, 16, 42];
  widths.forEach((w, i) => {
    sheet.getRange(`${colLetter(i + 1)}:${colLetter(i + 1)}`).format.columnWidth = w;
  });
  sheet.getRange(`A5:${lastCol}${4 + data.length}`).format.rowHeight = 58;
  sheet.getRange(`H5:H${4 + data.length}`).dataValidation = {
    rule: { type: "list", values: ["High", "Medium", "Low"] },
  };
  sheet.getRange(`L5:L${4 + data.length}`).dataValidation = {
    rule: { type: "list", values: ["TBC", "Not Started", "In Progress", "Confirmed", "Blocked"] },
  };
}

const workbook = Workbook.create();
const all = workbook.worksheets.add("KPI Tracker");
setupSheet(all, "Service Centre KPI Tracker", rows);

const groups = [
  ["Financial", "Financial"],
  ["Productivity", "Productivity and Efficiency"],
  ["Operations", "Operations & Customer"],
];
for (const [sheetName, category] of groups) {
  const sheet = workbook.worksheets.add(sheetName);
  setupSheet(sheet, sheetName, rows.filter((r) => r.category === category));
}

const dict = workbook.worksheets.add("Data Dictionary");
dict.showGridLines = false;
dict.getRange("A1:B1").merge();
dict.getRange("A1").values = [["Data Dictionary"]];
dict.getRange("A1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF", size: 16 },
};
dict.getRange(`A3:B${dictionary.length + 2}`).values = dictionary;
dict.getRange("A3:B3").format = {
  fill: "#134E4A",
  font: { bold: true, color: "#FFFFFF" },
};
dict.getRange(`A3:B${dictionary.length + 2}`).format.wrapText = true;
dict.getRange("A:A").format.columnWidth = 24;
dict.getRange("B:B").format.columnWidth = 90;
dict.getRange(`A4:B${dictionary.length + 2}`).format.rowHeight = 34;
dict.freezePanes.freezeRows(3);

for (const sheetName of ["KPI Tracker", "Financial", "Productivity", "Operations", "Data Dictionary"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName.replace(/[^A-Za-z0-9]/g, "_")}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
await fs.writeFile(`${outputDir}/verification_formula_errors.ndjson`, errors.ndjson, "utf8");

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
