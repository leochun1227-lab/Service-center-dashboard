import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/field_source_workbook";
const outputPath = `${outputDir}/service_center_data_center_field_sources.xlsx`;

await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();

const columns = [
  "Picture field",
  "Can get?",
  "Source",
  "Current available field",
  "Source detail",
  "Dashboard use",
  "Note",
];

const sheets = [
  {
    name: "1_DIM_VEHICLE",
    title: "1. DIM_VEHICLE",
    rows: [
      ["VehicleID", "Yes", "C4C", "SerialID", "Tickets sheet", "Vehicle key / join proxy", "Current export has SerialID, not a clean VehicleID column. Use SerialID as vehicle identifier unless a true VehicleID source is added."],
      ["VIN", "Missing", "Missing", "Missing", "Missing", "Vehicle lookup / warranty history", "No VIN column found in current C4C/SAP export."],
      ["ChassisNo", "Yes", "C4C", "ChassisNumber", "Tickets sheet", "Vehicle identification", "Field exists as ChassisNumber."],
      ["Model", "Missing", "Missing", "Missing", "Missing", "Model reliability analysis", "No model field found in current export."],
      ["ModelYear", "Missing", "Missing", "Missing", "Missing", "Age/model-year split", "No model year field found in current export."],
    ],
  },
  {
    name: "2_FACT_SERVICE_ORDER",
    title: "2. FACT_SERVICE_ORDER",
    rows: [
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Primary service order key", "Use TicketID as ServiceOrderID."],
      ["VehicleID", "Yes", "C4C", "SerialID", "Tickets sheet", "Join to DIM_VEHICLE", "Use SerialID as VehicleID proxy."],
      ["ServiceType", "Yes", "C4C", "TicketType / TicketTypeText", "Tickets sheet", "Repair vs PDI / service type split", "Z007/Z010 and text are available."],
      ["Source", "Yes", "C4C", "TicketType / TicketTypeText", "Tickets sheet", "Classify source/workflow", "No separate Source column found; can derive from ticket type or set as C4C."],
      ["CreatedDate", "Yes", "C4C", "CreatedOn", "Tickets sheet", "Created trend / aging start", "Confirmed by user: CreatedDate = C4C CreatedOn."],
      ["BookingDate", "Missing", "Missing", "Missing", "Missing", "Waiting days definition", "Not found in current C4C/SAP export."],
      ["ArrivalDateTime", "Missing", "Missing", "Missing", "Missing", "Arrival / waiting analysis", "Not found in current C4C/SAP export."],
      ["WorkStartDateTime", "Missing", "Missing", "Missing", "Missing", "Repair cycle start", "Not found in current C4C/SAP export."],
      ["CompletedDateTime", "Partial", "C4C", "ResolvedOnDateTime / ClaimApprovedOnDateTime / Billing date", "Tickets sheet / SAP billing lookup", "Completion proxy", "No exact CompletedDateTime. Dashboard currently can use workflow-derived completion proxy."],
      ["DepartedDateTime", "Missing", "Missing", "Missing", "Missing", "Departed / close-out analysis", "Not found in current C4C/SAP export."],
      ["Status", "Yes", "C4C", "TicketStatus / TicketStatusText", "Tickets sheet", "Workflow status / backlog", "Confirmed available."],
      ["Priority", "Partial", "C4C", "TicketSeverity", "Tickets sheet", "Priority card / high-priority exceptions", "No field named Priority. TicketSeverity exists and may be usable if business confirms it maps to priority."],
    ],
  },
  {
    name: "3_FACT_SERVICE_TASK",
    title: "3. FACT_SERVICE_TASK",
    rows: [
      ["TaskID", "Missing", "Missing", "Missing", "Missing", "Task-level key", "No task-level table in current export."],
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join back to service order", "Ticket-level only; task-level child rows are not available."],
      ["TaskCategory", "Missing", "Missing", "Missing", "Missing", "Task category analysis", "No task category field found."],
      ["Component", "Missing", "Missing", "Missing", "Missing", "Problem/component ranking", "No component field found."],
      ["ProblemDescription", "Partial", "C4C", "TicketName", "Tickets sheet", "Problem text / keyword analysis", "TicketName may be a weak proxy, not a structured problem description."],
      ["RepairAction", "Missing", "Missing", "Missing", "Missing", "Repair action analysis", "No repair action field found."],
      ["AssignedTechnicianID", "Missing", "Missing", "Missing", "Missing", "Assigned worker / team performance", "No reliable assigned technician ID found."],
      ["EstimatedHours", "Partial", "C4C", "Z1Z8TimeConsumed", "Tickets sheet", "Estimated/consumed time proxy", "Current logic treats this as estimate/proxy, not reliable actual task hours."],
      ["ActualHours", "Partial", "SAP", "TotalLabourHours", "SAP labour patch / Tickets sheet", "Actual labour hours", "Available at ticket level after SAP labour patch, not task-level."],
      ["TaskStatus", "Missing", "Missing", "Missing", "Missing", "Task completion split", "No task status field found."],
      ["WaitingReason", "Missing", "Missing", "Missing", "Missing", "Delay reason analysis", "No waiting reason field found."],
      ["ReworkFlag", "Missing", "Missing", "Missing", "Missing", "Rework tracking", "No rework flag field found."],
    ],
  },
  {
    name: "4_FACT_LABOUR",
    title: "4. FACT_LABOUR",
    rows: [
      ["LabourID", "Missing", "Missing", "Missing", "Missing", "Labour row key", "No labour transaction table in current export."],
      ["TaskID", "Missing", "Missing", "Missing", "Missing", "Join to service task", "No task-level labour rows."],
      ["TechnicianID", "Missing", "Missing", "Missing", "Missing", "Join to technician dimension", "No technician ID found."],
      ["WorkDate", "Missing", "Missing", "Missing", "Missing", "Daily labour trend", "No work-date field found in current labour patch."],
      ["StartTime", "Missing", "Missing", "Missing", "Missing", "Time clock analysis", "No start time field found."],
      ["EndTime", "Missing", "Missing", "Missing", "Missing", "Time clock analysis", "No end time field found."],
      ["ActualHours", "Partial", "SAP", "TotalLabourHours", "SAP labour patch / Tickets sheet", "Actual labour hours", "Available grouped by TicketID, not by LabourID/TaskID/TechnicianID."],
      ["WorkType", "Missing", "Missing", "Missing", "Missing", "Work type split", "No work type field found."],
      ["Comment", "Missing", "Missing", "Missing", "Missing", "Labour note detail", "No labour comment field found."],
    ],
  },
  {
    name: "5_DIM_TECHNICIAN",
    title: "5. DIM_TECHNICIAN",
    rows: [
      ["TechnicianID", "Missing", "Missing", "Missing", "Missing", "Technician key", "No technician dimension in current export."],
      ["TechnicianName", "Partial", "C4C", "Role_40_InvolvedPartyName", "Tickets sheet", "Worker display name", "Exists, but current notes say it is not reliable as assigned service worker."],
      ["Team", "Missing", "Missing", "Missing", "Missing", "Team workload", "No team field found."],
      ["SkillType", "Missing", "Missing", "Missing", "Missing", "Skill mix analysis", "No skill type field found."],
      ["ActiveFlag", "Missing", "Missing", "Missing", "Missing", "Active technician filter", "No active flag field found."],
    ],
  },
  {
    name: "6_FACT_INVOICE",
    title: "6. FACT_INVOICE",
    rows: [
      ["InvoiceID", "Yes", "C4C", "ERPInvoiceNumber", "Tickets sheet", "Invoice key", "Invoice number exists in C4C export."],
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join invoice to service order", "Use TicketID as ServiceOrderID."],
      ["InvoiceNo", "Yes", "C4C", "ERPInvoiceNumber", "Tickets sheet", "Invoice tracking", "Same as InvoiceID unless a separate SAP invoice ID is added."],
      ["InvoiceType", "Partial", "C4C", "Role_1001_InvolvedPartyName / Role_1001_InvolvedPartyID", "Tickets sheet", "Internal vs external logic", "Can derive invoice scope, but no explicit InvoiceType field."],
      ["InvoiceDate", "Yes", "SAP", "Billing date", "SAPInvoiceLookup / Tickets sheet after HANA enrichment", "Invoice month / trend", "Fetched from SAP by ERPInvoiceNumber."],
      ["DueDate", "Missing", "Missing", "Missing", "Missing", "Overdue receivables", "No due date field found."],
      ["InvoiceStatus", "Partial", "C4C", "TicketStatusText", "Tickets sheet", "Invoice workflow status", "No dedicated invoice status; can infer from ticket status."],
      ["LabourAmount", "Missing", "Missing", "Missing", "Missing", "Labour revenue split", "No labour amount field found."],
      ["PartsAmount", "Missing", "Missing", "Missing", "Missing", "Parts revenue split", "No parts amount field found."],
      ["GSTAmount", "Missing", "Missing", "Missing", "Missing", "Tax split", "No GST amount field found."],
      ["TotalAmount", "Yes", "SAP", "ERPInvoiceNumberPrice", "SAPInvoiceLookup / Tickets sheet after HANA enrichment", "Invoice amount", "Fetched from SAP by ERPInvoiceNumber."],
      ["PaidAmount", "Missing", "Missing", "Missing", "Missing", "Cash collected", "No paid amount field found."],
      ["BalanceAmount", "Missing", "Missing", "Missing", "Missing", "Outstanding balance", "No balance amount field found."],
      ["PaymentDate", "Missing", "Missing", "Missing", "Missing", "Payment timing", "No payment date field found."],
      ["PaymentStatus", "Missing", "Missing", "Missing", "Missing", "Payment status", "No payment status field found."],
    ],
  },
];

const sapRows = [
  ["SAP field currently available", "Mapped dashboard field", "How it is obtained", "Available in workbook", "Note"],
  ["Billing date", "FACT_INVOICE.InvoiceDate", "Fetched from SAP/HANA by ERPInvoiceNumber in sap_invoice_enrich_from_hana.py", "Yes", "Used for invoice month trend."],
  ["ERPInvoiceNumberPrice", "FACT_INVOICE.TotalAmount", "Fetched from SAP/HANA by ERPInvoiceNumber in sap_invoice_enrich_from_hana.py", "Yes", "Used as invoice amount."],
  ["TotalLabourHours", "FACT_LABOUR.ActualHours / FACT_SERVICE_TASK.ActualHours proxy", "Patched from SAP labour source and grouped by TicketID in apply_sap_labour_hours.py", "Partial", "Ticket-level total only, not task/technician transaction detail."],
  ["Worker", "DIM_TECHNICIAN.TechnicianName proxy", "SAP labour loader has Worker internally before grouping", "Partial", "Not retained as a reliable technician dimension in the final dashboard export."],
  ["ERPInvoiceNumber", "FACT_INVOICE.InvoiceID / InvoiceNo lookup key", "Present in C4C export and used to query SAP invoice data", "C4C key used for SAP lookup", "Not counted as SAP-created in current pipeline, but required for SAP join."],
];

function addSheet({ name, title, rows }) {
  const sheet = workbook.worksheets.add(name);
  sheet.showGridLines = false;
  sheet.getRange("A1:G1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: "#102A43",
    font: { bold: true, color: "#FFFFFF", size: 15 },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sheet.getRange("A2:G2").merge();
  sheet.getRange("A2").values = [["Source judgement based on current C4C exports, SAP invoice enrichment, and SAP labour patch scripts in the dashboard repo."]];
  sheet.getRange("A2").format = {
    fill: "#EAF2F8",
    font: { color: "#102A43" },
    wrapText: true,
  };
  sheet.getRange("A4:G4").values = [columns];
  sheet.getRange("A4:G4").format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF" },
    wrapText: true,
    horizontalAlignment: "center",
    verticalAlignment: "center",
    borders: { preset: "all", style: "thin", color: "#D9E2F3" },
  };
  sheet.getRange(`A5:G${rows.length + 4}`).values = rows;
  sheet.getRange(`A5:G${rows.length + 4}`).format = {
    wrapText: true,
    verticalAlignment: "top",
    borders: { preset: "all", style: "thin", color: "#D9E2F3" },
  };
  for (let i = 0; i < rows.length; i += 1) {
    const rowNumber = i + 5;
    const status = rows[i][1];
    const source = rows[i][2];
    const statusFill = status === "Yes" ? "#E2F0D9" : status === "Partial" ? "#FFF2CC" : "#F4CCCC";
    const sourceFill = source === "C4C" ? "#D9EAF7" : source === "SAP" ? "#E2F0D9" : source === "Missing" ? "#F4CCCC" : "#FFF2CC";
    sheet.getRange(`B${rowNumber}`).format.fill = statusFill;
    sheet.getRange(`C${rowNumber}`).format.fill = sourceFill;
  }
  [24, 14, 14, 28, 28, 28, 46].forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
  });
  sheet.getRange(`A1:G${rows.length + 4}`).format.autofitRows();
  sheet.freezePanes.freezeRows(4);
}

for (const sheetSpec of sheets) {
  addSheet(sheetSpec);
}

const sapSheet = workbook.worksheets.add("7_SAP_FIELDS");
sapSheet.showGridLines = false;
sapSheet.getRange("A1:E1").merge();
sapSheet.getRange("A1").values = [["7. SAP Fields Currently Available"]];
sapSheet.getRange("A1").format = {
  fill: "#375623",
  font: { bold: true, color: "#FFFFFF", size: 15 },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
sapSheet.getRange("A3:E8").values = sapRows;
sapSheet.getRange("A3:E3").format = {
  fill: "#70AD47",
  font: { bold: true, color: "#FFFFFF" },
  wrapText: true,
  horizontalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#D9E2F3" },
};
sapSheet.getRange("A4:E8").format = {
  wrapText: true,
  verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: "#D9E2F3" },
};
sapSheet.getRange("D4:D8").format.fill = "#E2F0D9";
[30, 32, 56, 22, 50].forEach((width, index) => {
  sapSheet.getRangeByIndexes(0, index, 1, 1).format.columnWidth = width;
});
sapSheet.getRange("A1:E8").format.autofitRows();
sapSheet.freezePanes.freezeRows(3);

for (const sheet of workbook.worksheets.items) {
  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    `${outputDir}/${sheet.name.replace(/[\\/:*?"<>|]/g, "_")}.png`,
    new Uint8Array(await preview.arrayBuffer()),
  );
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);
console.log(outputPath);
