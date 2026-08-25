import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_dashboard.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/field_supplement_work";
const outputPath = `${outputDir}/service_center_dashboard_field_supplement.xlsx`;

await fs.mkdir(outputDir, { recursive: true });

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

function styleHeader(range, fill = "#0F4C81") {
  range.format = {
    fill,
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#D9E2F3" },
  };
}

function styleBody(range, fill = "#FFFFFF") {
  range.format = {
    fill,
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "all", style: "thin", color: "#D9E2F3" },
  };
}

function setColumnWidths(sheet) {
  const widths = [26, 34, 18, 34, 24, 34, 40, 34, 34];
  widths.forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });
}

const featureSheet = workbook.worksheets.getOrAdd("Feature Field Needs");
featureSheet.showGridLines = false;
featureSheet.getRange("A1:I80").clear({ applyTo: "all" });
featureSheet.getRange("A1:I1").merge();
featureSheet.getRange("A1").values = [["Dashboard Feature Field Needs"]];
featureSheet.getRange("A1").format = {
  fill: "#102A43",
  font: { bold: true, color: "#FFFFFF", size: 16 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
featureSheet.getRange("A2:I2").merge();
featureSheet.getRange("A2").values = [[
  "Updated from user confirmation: C4C CreatedOn, status/dealer/ticket type exist; Priority and technician/task/labour detail tables are not available yet.",
]];
featureSheet.getRange("A2").format = {
  fill: "#EAF2F8",
  font: { color: "#102A43" },
  wrapText: true,
  verticalAlignment: "center",
};

const headers = [[
  "Feature / Card",
  "Dashboard metric",
  "Build readiness",
  "Fields confirmed available",
  "Fields still needed",
  "Calculation / rule",
  "Source table / field",
  "User confirmation",
  "Implementation note",
]];
featureSheet.getRange("A4:I4").values = headers;
styleHeader(featureSheet.getRange("A4:I4"));

const rows = [
  [
    "Cycle time",
    "Average waiting days",
    "Can start after date-field check",
    "TicketID; C4C CreatedOn; TicketStatusText; DealerName / DealerID; TicketType",
    "BookingDate and/or ArrivalDateTime and/or WorkStartDateTime. Need final business definition.",
    "Preferred: WorkStartDateTime - CreatedOn. Alternative: ArrivalDateTime - BookingDate.",
    "FACT_SERVICE_ORDER.CreatedDate = C4C CreatedOn; date milestone fields to be checked",
    "CreatedDate = C4C CreatedOn; status/dealer/ticket type exist; metric needs checking",
    "Add as KPI plus dealer split once the date milestone fields are confirmed.",
  ],
  [
    "Cycle time",
    "Average repair cycle",
    "Can start after date-field check",
    "TicketID; C4C CreatedOn; TicketStatusText; DealerName / DealerID; TicketType",
    "WorkStartDateTime; CompletedDateTime",
    "CompletedDateTime - WorkStartDateTime",
    "FACT_SERVICE_ORDER.WorkStartDateTime; FACT_SERVICE_ORDER.CompletedDateTime",
    "Metric needs checking",
    "Exclude cancelled tickets; optionally split Repair ticket vs PDI ticket.",
  ],
  [
    "Cycle time",
    "Overdue unfinished orders",
    "Can build with current status/date fields",
    "TicketID; C4C CreatedOn; TicketStatusText; DealerName / DealerID; TicketType",
    "Overdue threshold rule. Priority field is missing if threshold depends on priority.",
    "Today - CreatedOn > threshold and status not in completed/invoiced/cancelled statuses.",
    "C4C CreatedOn; TicketStatusText; DealerName / DealerID; TicketType",
    "Metric needs checking; priority not available",
    "Add a threshold setting such as 7/14/30 days, or separate thresholds by ticket type.",
  ],
  [
    "Priority",
    "High priority tickets card",
    "Blocked until field is added",
    "TicketID; C4C CreatedOn; TicketStatusText; DealerName / DealerID; TicketType",
    "Priority",
    "Count open high-priority tickets; optionally show oldest high-priority ticket age.",
    "FACT_SERVICE_ORDER.Priority",
    "High priority should have its own card, but field is not currently available",
    "When Priority is added, use a small KPI card and an exception table.",
  ],
  [
    "Technician performance",
    "Actual hours by technician",
    "Blocked",
    "DealerName / DealerID; TicketID if joined from current C4C export",
    "TechnicianID or TechnicianName; WorkDate; ActualHours; TaskID preferred",
    "Sum ActualHours by Technician and period.",
    "DIM_TECHNICIAN; FACT_LABOUR; FACT_SERVICE_TASK",
    "Technician/task/labour detail tables are not available",
    "Minimum version requires TicketID, TechnicianName, ActualHours, WorkDate.",
  ],
  [
    "Technician performance",
    "Assigned tasks / completed tasks / average task hours",
    "Blocked",
    "No reliable task-level source confirmed",
    "TaskID; AssignedTechnicianID; TaskStatus; ActualHours; EstimatedHours",
    "Task count, completed task count, ActualHours / completed task count.",
    "FACT_SERVICE_TASK.TaskID; AssignedTechnicianID; TaskStatus; ActualHours",
    "Technician/task/labour detail tables are not available",
    "Do not calculate from ticket-level data only, because one ticket can contain multiple tasks/workers.",
  ],
  [
    "Technician performance",
    "Team workload and skill type mix",
    "Blocked",
    "No technician dimension confirmed",
    "Team; SkillType; ActiveFlag; TechnicianID",
    "Group technician actual hours and task count by Team and SkillType.",
    "DIM_TECHNICIAN.Team; SkillType; ActiveFlag",
    "DIM_TECHNICIAN is not available",
    "Useful later as a separate Technician Performance page.",
  ],
  [
    "Labour variance",
    "Dealer actual-hours variance",
    "Blocked until estimated and actual hours are both available",
    "DealerName / DealerID; TicketID; TicketType",
    "EstimatedHours; ActualHours; TaskID preferred; Component/TaskCategory optional",
    "ActualHours - EstimatedHours; variance rate = variance / EstimatedHours.",
    "FACT_SERVICE_TASK.EstimatedHours; FACT_SERVICE_TASK.ActualHours; DealerName / DealerID",
    "Dealer labour variance can be added",
    "If EstimatedHours is missing, show actual-hours ranking only, not variance.",
  ],
];

featureSheet.getRange(`A5:I${rows.length + 4}`).values = rows;
styleBody(featureSheet.getRange(`A5:I${rows.length + 4}`));
featureSheet.getRange("C5:C12").format = {
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: "#D9E2F3" },
};
featureSheet.getRange("C5:C7").format.fill = "#FFF7D6";
featureSheet.getRange("C8:C12").format.fill = "#FDE2E2";
featureSheet.getRange("H5:H12").format.fill = "#E8F5E9";
featureSheet.getRange("A13:I13").merge();
featureSheet.getRange("A13").values = [[
  "Decision needed: define overdue threshold and confirm which date fields exist in the C4C export for booking/arrival/work start/completion/departure.",
]];
featureSheet.getRange("A13").format = {
  fill: "#FFF2CC",
  font: { bold: true, color: "#7A4B00" },
  wrapText: true,
  verticalAlignment: "center",
  borders: { preset: "outside", style: "thin", color: "#D6B656" },
};
featureSheet.getRange("A1:I13").format.autofitRows();
setColumnWidths(featureSheet);
featureSheet.freezePanes.freezeRows(4);

const matrixSheet = workbook.worksheets.getItem("Value Matrix");
matrixSheet.getRange("A23:M23").merge();
matrixSheet.getRange("A23").values = [["Requested Dashboard Additions - Field Supplements"]];
matrixSheet.getRange("A23").format = {
  fill: "#102A43",
  font: { bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
matrixSheet.getRange("A24:C24").values = [["Final Dashboard Fields", null, null]];
matrixSheet.getRange("E24:G24").values = [["C4C Source Fields", null, null]];
matrixSheet.getRange("I24:K24").values = [["Extra Source Fields Needed", null, null]];
matrixSheet.getRange("L24:M24").values = [["Readiness", null]];
styleHeader(matrixSheet.getRange("A24:C24"), "#1F4E79");
styleHeader(matrixSheet.getRange("E24:G24"), "#1F4E79");
styleHeader(matrixSheet.getRange("I24:K24"), "#70AD47");
styleHeader(matrixSheet.getRange("L24:M24"), "#C00000");
matrixSheet.getRange("A25:C25").values = [["Field wanted", "Example / value", "Purpose"]];
matrixSheet.getRange("E25:G25").values = [["C4C field", "Current status", "Used for"]];
matrixSheet.getRange("I25:K25").values = [["Needed field", "Current status", "Used for"]];
matrixSheet.getRange("L25:M25").values = [["Build status", "Note"]];
styleHeader(matrixSheet.getRange("A25:C25"), "#5B9BD5");
styleHeader(matrixSheet.getRange("E25:G25"), "#5B9BD5");
styleHeader(matrixSheet.getRange("I25:K25"), "#A9D18E");
styleHeader(matrixSheet.getRange("L25:M25"), "#E06666");

const supplementRows = [
  ["Average waiting days", "e.g. 6.4 days", "Waiting time KPI", null, "CreatedOn", "Available as C4C CreatedOn", "Start date", null, "WorkStartDateTime or ArrivalDateTime", "Need to check", "End date for waiting", "Partial", "Need date-field confirmation"],
  ["Average repair cycle", "e.g. 2.1 days", "Repair duration KPI", null, "TicketStatusText", "Available", "Exclude cancelled/completed logic", null, "WorkStartDateTime + CompletedDateTime", "Need to check", "Repair duration", "Partial", "Can build after date fields confirmed"],
  ["Overdue unfinished orders", "e.g. 18 open > 14 days", "Open-ticket risk", null, "DealerName / DealerID", "Available", "Dealer split", null, "Overdue threshold", "Business rule needed", "Define overdue", "Ready with rule", "Priority-based threshold blocked until Priority exists"],
  ["High priority ticket card", "e.g. 4 high priority open", "Exception KPI", null, "TicketType", "Available", "Repair vs PDI split", null, "Priority", "Missing", "High-priority filter", "Blocked", "Add field when source provides Priority"],
  ["Dealer labour variance", "Actual - estimated hours", "Dealer productivity variance", null, "DealerName / DealerID", "Available", "Dealer grouping", null, "EstimatedHours + ActualHours", "Missing / not reliable yet", "Variance calculation", "Blocked", "Can show actual-hours ranking only if EstimatedHours is missing"],
  ["Technician performance", "Hours/tasks by worker", "Worker productivity", null, "TicketID", "Available", "Join key", null, "TechnicianID/Name, TaskID, WorkDate, ActualHours", "Missing", "Technician workload", "Blocked", "DIM_TECHNICIAN / FACT_TASK / FACT_LABOUR not available"],
];
matrixSheet.getRange(`A26:M${supplementRows.length + 25}`).values = supplementRows;
styleBody(matrixSheet.getRange(`A26:M${supplementRows.length + 25}`));
matrixSheet.getRange("L26:L28").format.fill = "#FFF2CC";
matrixSheet.getRange("L29:L31").format.fill = "#F4CCCC";
matrixSheet.getRange("M26:M31").format.fill = "#FCE4D6";
matrixSheet.getRange("A23:M31").format.autofitRows();

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

for (const sheet of workbook.worksheets.items) {
  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    `${outputDir}/${sheet.name.replace(/[\\/:*?"<>|]/g, "_")}_updated.png`,
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(outputPath);
