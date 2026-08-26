import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/data_center_excel_update";
const outputPath = `${outputDir}/service_center_data_center_ro_finance_model.xlsx`;

const note =
  "Aligned to the current 6-table Canva model. Source color guide: SAP = red, C4C = dark green, black = no confirmed source in current export.";

const headers = [
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
    name: "1_FACT_RO_FINANCE",
    title: "1. FACT_RO_FINANCE",
    rows: [
      ["ROFinanceID", "Missing", "", "", "", "RO finance row key", "No confirmed source key yet; can be generated later if needed."],
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join financial KPI inputs to service order", "Use TicketID as ServiceOrderID."],
      ["InvoiceID", "Yes", "SAP", "ERPInvoiceNumber", "C4C key enriched from SAP", "Join to invoice / finance source", "ERPInvoiceNumber is used to query SAP invoice data."],
      ["RepairOrderRevenue", "Yes", "SAP", "ERPInvoiceNumberPrice", "SAPInvoiceLookup / HANA enrichment", "Average Repair Order Amount numerator", "Current best proxy for total repair order revenue."],
      ["LaborRevenue", "Missing", "SAP", "", "Target SAP invoice line or GL split", "Effective labor rate and service gross %", "Need labour/labor revenue split from invoice or GL detail."],
      ["PartsRevenue", "Missing", "SAP", "", "Target SAP invoice line or GL split", "Parts gross % and gross profit margin", "Need parts revenue split from invoice or GL detail."],
      ["LaborCost", "Missing", "", "", "", "Gross profit margin and service gross %", "Source not confirmed. Likely needs payroll/costing source, not safe to mark SAP/C4C yet."],
      ["PartsCost", "Missing", "SAP", "", "Target SAP PO invoice / parts cost data", "Parts gross % and gross profit margin", "Need parts cost by repair order or period."],
      ["OperatingExpenses", "Missing", "SAP", "", "Target SAP finance / GL data", "Absorption rate denominator", "May be period-level rather than RO-level."],
      ["LaborHoursSold", "Partial", "C4C", "LabourHours", "Tickets sheet", "Effective labor rate denominator", "Current C4C labour hours can be a proxy until sold hours are confirmed."],
      ["PostedLaborRate", "Missing", "", "", "", "Effective labor rate benchmark", "No confirmed source; keep black until business confirms rate table/source."],
      ["FinancialPeriod", "Yes", "SAP", "Billing date", "SAPInvoiceLookup / HANA enrichment", "Monthly financial KPI grouping", "Can derive period from SAP billing date."],
    ],
  },
  {
    name: "2_FACT_SERVICE_ORDER",
    title: "2. FACT_SERVICE_ORDER",
    rows: [
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Primary service order key", "Use TicketID as ServiceOrderID."],
      ["VehicleID", "Yes", "C4C", "SerialID", "Tickets sheet", "Vehicle-level filter if needed", "Kept as an order attribute; no separate vehicle dimension in current model."],
      ["DealerYard", "Partial", "C4C", "Dealer / involved party field", "Tickets sheet", "Dealer / yard filter and ranking", "Confirm exact C4C dealer-yard field name."],
      ["ServiceType", "Yes", "C4C", "TicketType / TicketTypeText", "Tickets sheet", "Warranty / PDI / customer paid split", ""],
      ["Source", "Yes", "C4C", "TicketType / workflow logic", "Tickets sheet", "Classify workflow source", "Can be derived if no explicit Source column exists."],
      ["CreatedDate", "Yes", "C4C", "CreatedOn", "Tickets sheet", "Created trend / aging start", ""],
      ["BookingDate", "Missing", "", "", "", "Appointment and waiting days", "No confirmed field in current export."],
      ["ArrivalDateTime", "Missing", "", "", "", "Turnaround start / arrival analysis", "No confirmed field in current export."],
      ["CompletedDateTime", "Partial", "C4C", "ResolvedOnDateTime / ClaimApprovedOnDateTime", "Tickets sheet", "Completion trend and turnaround proxy", "Exact completed timestamp still needs business confirmation."],
      ["DepartedDateTime", "Missing", "", "", "", "Departed / close-out analysis", "No confirmed field in current export."],
      ["Status", "Yes", "C4C", "TicketStatus / TicketStatusText", "Tickets sheet", "Open, completed and backlog KPIs", ""],
      ["Priority", "Partial", "C4C", "TicketSeverity", "Tickets sheet", "Priority and exception filters", "Use only if business confirms severity maps to priority."],
      ["QuoteAmount", "Missing", "C4C", "", "Target C4C quote data", "Quotation approval and ARO context", "Source expected from C4C, but no confirmed current field."],
      ["QuoteStatus", "Missing", "C4C", "", "Target C4C quote data", "Quotation approval rate", "Source expected from C4C, but no confirmed current field."],
      ["AppointmentLeadTimeDays", "Missing", "C4C", "", "Target C4C appointment data", "Service appointment lead time", "Needs request/booked/appointment dates from C4C."],
      ["ComebackFlag", "Missing", "C4C", "", "Target C4C service history", "Comeback rate / fixed-first-time", "Can later be derived from repeat repair logic if data exists."],
    ],
  },
  {
    name: "3_FACT_SERVICE_TASK",
    title: "3. FACT_SERVICE_TASK",
    rows: [
      ["TaskID", "Missing", "C4C", "", "Target C4C task data", "Task-level key", "No confirmed task-level child table in current export."],
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join back to service order", "Ticket-level only in current data."],
      ["TaskCategory", "Missing", "C4C", "", "Target C4C task data", "Task category analysis", "No confirmed current field."],
      ["Component", "Missing", "C4C", "", "Target C4C task data", "Problem/component ranking", "No confirmed current field."],
      ["ProblemDescription", "Partial", "C4C", "TicketName", "Tickets sheet", "Problem text / keyword analysis", "TicketName is a weak proxy, not structured root cause."],
      ["RepairAction", "Missing", "C4C", "", "Target C4C task data", "Repair action analysis", "No confirmed current field."],
      ["AssignedTechnicianID", "Missing", "C4C", "", "Target C4C technician assignment", "Technician workload", "No reliable assigned technician ID in current export."],
      ["EstimatedHours", "Partial", "C4C", "Z1Z8TimeConsumed", "Tickets sheet", "Estimated or consumed hour proxy", "Needs confirmation before using as estimate."],
      ["ActualHours", "Partial", "C4C", "LabourHours / task proxy", "Tickets sheet", "Task effort proxy", "Actual paid/posted labour is handled in FACT_LABOUR."],
      ["TaskStatus", "Missing", "C4C", "", "Target C4C task data", "Task completion split", "No confirmed current field."],
      ["WaitingReason", "Missing", "C4C", "", "Target C4C task data", "Delay reason analysis", "No confirmed current field."],
      ["ReworkFlag", "Missing", "C4C", "", "Target C4C task data", "Rework signal", "No confirmed current field."],
    ],
  },
  {
    name: "4_FACT_LABOUR",
    title: "4. FACT_LABOUR",
    rows: [
      ["LabourID", "Missing", "", "", "", "Labour row key", "No confirmed row-level labour key."],
      ["TaskID", "Missing", "", "", "", "Join to service task", "No confirmed task-level labour link."],
      ["TechnicianID", "Missing", "", "", "", "Join to technician dimension", "No reliable technician ID retained in current export."],
      ["LabourDate", "Missing", "C4C", "", "Target C4C labour data", "Daily labour trend", "Need labour date from source data."],
      ["WorkerName", "Partial", "C4C", "Role_40_InvolvedPartyName / Worker", "Tickets sheet / labour source", "Technician display name", "Current worker name exists as a proxy but may not be reliable."],
      ["ClaimHours", "Yes", "C4C", "LabourHours", "Tickets sheet", "Claim labour hours", ""],
      ["ActualWorkHours", "Partial", "C4C", "Z1Z8TimeConsumed / labour proxy", "Tickets sheet", "Actual work hour proxy", "Needs business confirmation of the exact hour definition."],
      ["InvoicePaidHours", "Partial", "SAP", "TotalLabourHours", "SAP labour patch", "Paid labour hours", "Currently ticket-level after SAP labour patch, not row-level."],
      ["AvailableHours", "Missing", "", "", "", "Productivity denominator", "No confirmed source."],
      ["ClockedHours", "Missing", "", "", "", "Utilisation denominator", "No confirmed source."],
      ["SoldHours", "Missing", "", "", "", "Efficiency numerator", "No confirmed source."],
      ["NonProductiveHours", "Missing", "", "", "", "Non-productive time analysis", "No confirmed source."],
      ["WorkType", "Missing", "C4C", "", "Target C4C labour data", "Repair / diagnosis / inspection split", "No confirmed current field."],
    ],
  },
  {
    name: "5_DIM_TECHNICIAN",
    title: "5. DIM_TECHNICIAN",
    rows: [
      ["TechnicianID", "Missing", "C4C", "", "Target C4C technician master", "Technician key", "No confirmed technician dimension in current export."],
      ["TechnicianName", "Partial", "C4C", "Role_40_InvolvedPartyName", "Tickets sheet", "Worker display name", "Useful as a proxy only after assignment logic is confirmed."],
      ["Team", "Missing", "C4C", "", "Target C4C technician master", "Team workload", "No confirmed current field."],
      ["SkillType", "Missing", "", "", "", "Skill mix analysis", "No confirmed source."],
      ["ActiveFlag", "Missing", "", "", "", "Active technician filter", "No confirmed source."],
    ],
  },
  {
    name: "6_FACT_INVOICE",
    title: "6. FACT_INVOICE",
    rows: [
      ["InvoiceID", "Yes", "SAP", "ERPInvoiceNumber", "C4C key enriched from SAP", "Invoice key", "ERP invoice number is used to query SAP invoice data."],
      ["ServiceOrderID", "Yes", "SAP", "TicketID / ERPInvoiceNumber join", "C4C to SAP enrichment", "Join invoice to service order", "TicketID remains the service order key."],
      ["InvoiceNo", "Yes", "SAP", "ERPInvoiceNumber", "C4C key enriched from SAP", "Invoice tracking", ""],
      ["InvoiceType", "Partial", "SAP", "Invoice scope / derived type", "SAP or derived logic", "Internal vs external invoice split", "Needs explicit invoice type if available."],
      ["InvoiceDate", "Yes", "SAP", "Billing date", "SAPInvoiceLookup / HANA enrichment", "Invoice month and trend", ""],
      ["DueDate", "Missing", "SAP", "", "Target SAP receivables data", "Overdue receivables", "No confirmed current field."],
      ["InvoiceStatus", "Partial", "SAP", "Payment / billing status", "Target SAP invoice status", "Invoice workflow status", "No dedicated current field confirmed."],
      ["InvoiceAmount", "Yes", "SAP", "ERPInvoiceNumberPrice", "SAPInvoiceLookup / HANA enrichment", "Invoice amount / ARO", ""],
      ["LabourRevenueAmount", "Missing", "SAP", "", "Target SAP invoice line or GL split", "Labour gross margin", "Needs labour revenue split."],
      ["LabourCostAmount", "Missing", "SAP", "", "Target SAP cost split", "Effective labour rate / labour margin", "Needs labour cost split."],
      ["PartsRevenueAmount", "Missing", "SAP", "", "Target SAP invoice line or GL split", "Parts gross margin", "Needs parts revenue split."],
      ["PartsCostAmount", "Missing", "SAP", "", "Target SAP cost split", "Parts gross margin", "Needs parts cost split."],
      ["HoursBilled", "Partial", "SAP", "TotalLabourHours", "SAP labour patch", "Effective labour rate", "Current source is ticket-level labour hours."],
      ["GSTAmount", "Missing", "SAP", "", "Target SAP tax data", "Tax split", "No confirmed current field."],
      ["PaidAmount", "Missing", "SAP", "", "Target SAP payment data", "Cash collected", "No confirmed current field."],
      ["BalanceAmount", "Missing", "SAP", "", "Target SAP payment data", "Outstanding balance", "No confirmed current field."],
      ["PaymentStatus", "Missing", "SAP", "", "Target SAP payment data", "Payment status", "No confirmed current field."],
    ],
  },
];

function styleSourceCell(cell, source) {
  const color = source === "SAP" ? "#C7251A" : source === "C4C" ? "#075F43" : "#111827";
  cell.format = { font: { bold: Boolean(source), color } };
}

function styleAvailabilityCell(cell, value) {
  const color = value === "Yes" ? "#075F43" : value === "Partial" ? "#B45309" : "#111827";
  cell.format = { font: { bold: value !== "Missing", color } };
}

await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();

for (const spec of sheets) {
  const sheet = workbook.worksheets.add(spec.name);
  sheet.showGridLines = false;

  const rowCount = spec.rows.length + 4;
  sheet.getRange("A1:G1").merge();
  sheet.getRange("A1").values = [[spec.title]];
  sheet.getRange("A2:G2").merge();
  sheet.getRange("A2").values = [[note]];
  sheet.getRange("A4:G4").values = [headers];
  sheet.getRange(`A5:G${rowCount}`).values = spec.rows;

  sheet.getRange("A1:G1").format = {
    fill: spec.name.includes("DIM") ? "#0B62C5" : "#D92D20",
    font: { bold: true, color: "#FFFFFF", size: 16 },
  };
  sheet.getRange("A2:G2").format = {
    fill: "#EEF6FF",
    font: { italic: true, color: "#475569", size: 10 },
    wrapText: true,
  };
  sheet.getRange("A4:G4").format = {
    fill: "#E5E7EB",
    font: { bold: true, color: "#111827" },
    wrapText: true,
  };
  sheet.getRange(`A5:G${rowCount}`).format = {
    borders: { preset: "inside", style: "thin", color: "#D9E2EC" },
    wrapText: true,
  };
  sheet.getRange(`A4:G${rowCount}`).format.borders = {
    preset: "all",
    style: "thin",
    color: "#CBD5E1",
  };

  sheet.getRange("A:A").format.columnWidthPx = 190;
  sheet.getRange("B:B").format.columnWidthPx = 95;
  sheet.getRange("C:C").format.columnWidthPx = 90;
  sheet.getRange("D:D").format.columnWidthPx = 220;
  sheet.getRange("E:E").format.columnWidthPx = 225;
  sheet.getRange("F:F").format.columnWidthPx = 235;
  sheet.getRange("G:G").format.columnWidthPx = 330;
  sheet.getRange("1:1").format.rowHeightPx = 30;
  sheet.getRange("2:2").format.rowHeightPx = 34;
  sheet.getRange("4:4").format.rowHeightPx = 28;
  sheet.getRange(`5:${rowCount}`).format.rowHeightPx = 34;

  for (let i = 0; i < spec.rows.length; i += 1) {
    const row = i + 5;
    styleAvailabilityCell(sheet.getRange(`B${row}`), spec.rows[i][1]);
    styleSourceCell(sheet.getRange(`C${row}`), spec.rows[i][2]);
  }

  const table = sheet.tables.add(`A4:G${rowCount}`, true, `${spec.name.replace(/[^A-Za-z0-9]/g, "")}Table`);
  table.showFilterButton = true;
  sheet.freezePanes.freezeRows(4);
}

const scan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(scan.ndjson);

for (const spec of sheets) {
  const preview = await workbook.render({
    sheetName: spec.name,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(`${outputDir}/${spec.name}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const check = await workbook.inspect({
  kind: "table",
  sheetId: "2_FACT_SERVICE_ORDER",
  range: "A1:G12",
  include: "values",
  tableMaxRows: 12,
  tableMaxCols: 7,
  maxChars: 5000,
});
console.log(check.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`output=${outputPath}`);
