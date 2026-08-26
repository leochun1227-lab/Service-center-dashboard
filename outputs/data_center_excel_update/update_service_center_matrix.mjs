import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_data_center_.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/data_center_excel_update";
const outputPath = `${outputDir}/service_center_data_center_kpi_expanded.xlsx`;

const sourceNote = "Source judgement based on current C4C exports, SAP invoice enrichment, and SAP labour patch scripts in the dashboard repo.";
const columns = ["Picture field", "Can get?", "Source", "Current available field", "Source detail", "Dashboard use", "Note"];

const sheets = [
  {
    name: "7_FACT_INVOICE_LINE",
    title: "7. FACT_INVOICE_LINE",
    rows: [
      ["InvoiceLineID", "Missing", "Missing", "", "", "Line-level invoice key", "Need invoice line interface."],
      ["InvoiceID", "Yes", "C4C / SAP", "ERPInvoiceNumber", "Tickets sheet / SAP invoice lookup", "Join invoice line back to invoice header", "Use invoice number as InvoiceID until SAP line ID is available."],
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join revenue and margin back to service order", "Available now."],
      ["LineType", "Missing", "Missing", "", "", "Labour vs parts split", "Required for Service Gross % and Parts Gross %."],
      ["RevenueAmount", "Partial", "SAP", "ERPInvoiceNumberPrice", "Invoice header amount only", "ARO and total invoice revenue", "Need line revenue for labour / parts margin split."],
      ["CostAmount", "Missing", "Missing", "", "", "Gross profit and margin", "Need labour cost and parts cost."],
      ["HoursBilled", "Partial", "SAP / C4C", "TotalLabourHours / LabourHours", "Ticket-level labour hours", "Effective labour rate", "Need invoice-paid/billed hours by line."],
      ["Quantity", "Missing", "Missing", "", "", "Parts quantity / line validation", "Needed only when parts line detail is available."],
      ["PartID", "Missing", "Missing", "", "", "Part-level margin and usage", "Future parts interface."],
    ],
  },
  {
    name: "8_FACT_TECHNICIAN_TIME",
    title: "8. FACT_TECHNICIAN_TIME",
    rows: [
      ["TimeEntryID", "Missing", "Missing", "", "", "Time row key", "Need clocking / timesheet interface."],
      ["TechnicianID", "Missing", "Missing", "", "", "Join to technician dimension", "Current worker name exists, but no reliable technician ID."],
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join time to service order", "Available now."],
      ["WorkDate", "Missing", "Missing", "", "", "Daily productivity and utilisation", "Need labour date / timesheet date."],
      ["AvailableHours", "Missing", "Missing", "", "", "Utilisation denominator", "Need roster/capacity source."],
      ["ClockedHours", "Missing", "Missing", "", "", "Utilisation and attendance", "Need time clock source."],
      ["SoldHours", "Partial", "SAP / C4C", "TotalLabourHours / LabourHours", "Ticket-level labour hours", "Technician productivity", "Need sold/billed hours at worker-day level."],
      ["ActualWorkHours", "Partial", "SAP", "TotalLabourHours", "SAP labour patch", "Technician efficiency", "Available at ticket level, not true time-entry level."],
      ["NonProductiveHours", "Missing", "Missing", "", "", "Productivity bridge", "Need reason-coded non-productive time."],
    ],
  },
  {
    name: "9_FACT_QUOTE",
    title: "9. FACT_QUOTE",
    rows: [
      ["QuoteID", "Missing", "Missing", "", "", "Quote key", "Need quote object/interface."],
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join quote to service order", "Available now."],
      ["QuoteDate", "Missing", "Missing", "", "", "Quote cycle timing", "No dedicated quote date found."],
      ["QuoteAmount", "Yes", "C4C", "AmountIncludingTax", "Tickets sheet", "Quote value and ARO context", "Available as ticket quote amount."],
      ["QuoteStatus", "Partial", "C4C", "TicketStatusText", "Tickets sheet", "Quotation approval rate", "Can infer some statuses, but no clean quote status field."],
      ["ApprovedDate", "Partial", "C4C", "ApprovalDate", "Tickets sheet", "Approval timing", "Available where populated."],
      ["DeclinedReason", "Missing", "Missing", "", "", "Lost quote analysis", "Need quote decline reason."],
    ],
  },
  {
    name: "10_FACT_APPOINTMENT",
    title: "10. FACT_APPOINTMENT",
    rows: [
      ["AppointmentID", "Missing", "Missing", "", "", "Appointment key", "Need scheduler / booking interface."],
      ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join appointment to service order", "Available now."],
      ["RequestDate", "Partial", "C4C", "CreatedOn", "Tickets sheet", "Lead-time start proxy", "Use only as proxy until request date exists."],
      ["BookedDate", "Missing", "Missing", "", "", "Booking conversion timing", "No booking date in current export."],
      ["AppointmentDate", "Missing", "Missing", "", "", "Appointment lead time", "Need scheduled appointment date."],
      ["ArrivalDate", "Missing", "Missing", "", "", "Arrival / no-show / lead-time close", "No arrival date in current export."],
      ["LeadTimeDays", "Missing", "Derived", "AppointmentDate - RequestDate", "Needs both dates", "Target KPI: 3 to 7 days; can derive when source dates exist."],
    ],
  },
  {
    name: "11_FACT_COMEBACK",
    title: "11. FACT_COMEBACK",
    rows: [
      ["ComebackID", "Missing", "Missing", "", "", "Comeback key", "Need comeback/rework interface or rule."],
      ["OriginalServiceOrderID", "Missing", "Missing", "", "", "Original repair order", "Need link to original RO."],
      ["ComebackServiceOrderID", "Missing", "Missing", "", "", "Repeat repair order", "Need link to comeback RO."],
      ["VehicleID", "Yes", "C4C", "SerialID / ChassisNumber", "Tickets sheet", "Detect repeat issues by vehicle", "Available as vehicle proxy."],
      ["ComebackDate", "Missing", "Missing", "", "", "Comeback period", "Need comeback date or repeat repair date rule."],
      ["ComebackReason", "Missing", "Missing", "", "", "Root-cause analysis", "Need reason coding."],
      ["IsComeback", "Missing", "Derived", "", "Needs business rule", "Comeback Rate / Fixed First Time Right", "Typical rule: repeat repair within X days for same vehicle/component."],
    ],
  },
];

const labourRows = [
  ["LabourID", "Missing", "Missing", "", "", "Labour row key", "Need labour line/time-entry source."],
  ["ServiceOrderID", "Yes", "C4C", "TicketID", "Tickets sheet", "Join labour to service order", "Available now."],
  ["LabourDate", "Missing", "Missing", "", "", "Daily labour view", "Need labour date / work date."],
  ["WorkerName", "Partial", "C4C", "Role_40_InvolvedPartyName", "Tickets sheet", "Worker display", "Name exists but technician assignment should be confirmed."],
  ["ClaimHours", "Yes", "C4C", "LabourHours", "Tickets sheet", "Claimed labour hours", "Available now."],
  ["ActualWorkHours", "Partial", "SAP", "TotalLabourHours", "SAP labour patch / Tickets sheet", "Actual work-hour comparison", "Available at ticket level."],
  ["InvoicePaidHours", "Missing", "Missing", "", "", "Paid/billed labour hour comparison", "Need invoice labour-hours source."],
  ["WorkType", "Partial", "C4C", "TicketType / TicketTypeText", "Tickets sheet", "Repair vs PDI labour split", "Can derive broad work type only."],
  ["Comment", "Missing", "Missing", "", "", "Labour note detail", "Need labour line comments."],
];

function matrix(title, rows) {
  return [
    [title, "", "", "", "", "", ""],
    [sourceNote, "", "", "", "", "", ""],
    ["", "", "", "", "", "", ""],
    columns,
    ...rows,
  ];
}

function applyMatrix(sheet, title, rows) {
  const data = matrix(title, rows);
  const range = sheet.getRangeByIndexes(0, 0, data.length, 7);
  range.values = data;

  sheet.mergeCells("A1:G1");
  sheet.mergeCells("A2:G2");
  sheet.getRange("A1:G1").format = {
    fill: "#102A43",
    font: { bold: true, color: "#FFFFFF", fontSize: 15 },
    horizontalAlignment: "center",
  };
  sheet.getRange("A2:G2").format = {
    fill: "#EAF2F8",
    font: { color: "#102A43", fontSize: 11 },
    wrapText: true,
  };
  sheet.getRange("A4:G4").format = {
    fill: "#1F4E79",
    font: { bold: true, color: "#FFFFFF", fontSize: 11 },
    horizontalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#D9E2F3" },
  };

  const body = sheet.getRangeByIndexes(4, 0, rows.length, 7);
  body.format = {
    font: { fontSize: 11 },
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#D9E2F3" },
  };
  sheet.getRangeByIndexes(4, 1, rows.length, 1).format.fill = "#FCE4D6";
  rows.forEach((row, index) => {
    const canGet = row[1];
    const fill = canGet === "Yes" ? "#E2F0D9" : canGet === "Partial" ? "#FFF2CC" : "#FCE4D6";
    sheet.getRangeByIndexes(4 + index, 1, 1, 1).format.fill = fill;
  });
  sheet.getRangeByIndexes(4, 2, rows.length, 1).format.fill = "#D9EAF7";
  sheet.getRangeByIndexes(4, 0, rows.length, 1).format.font = { bold: true, fontSize: 11 };

  sheet.getRange("A:A").format.columnWidth = 24;
  sheet.getRange("B:B").format.columnWidth = 12;
  sheet.getRange("C:C").format.columnWidth = 16;
  sheet.getRange("D:D").format.columnWidth = 28;
  sheet.getRange("E:E").format.columnWidth = 34;
  sheet.getRange("F:F").format.columnWidth = 34;
  sheet.getRange("G:G").format.columnWidth = 34;
  sheet.getRange("A2:G2").format.rowHeight = 34;
  sheet.getRangeByIndexes(4, 0, rows.length, 7).format.rowHeight = 33;
  sheet.freezePanes.freezeRows(4);
  sheet.showGridLines = false;
}

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const labour = workbook.worksheets.getItem("4_FACT_LABOUR");
labour.getRange("A1:G40").clear({ applyTo: "all" });
applyMatrix(labour, "4. FACT_LABOUR", labourRows);

try {
  workbook.worksheets.getItem("7_SAP_FIELDS").name = "12_SOURCE_FIELDS";
} catch {
  // If rename is not supported in this runtime, leave the source sheet as-is.
}

for (const spec of sheets) {
  const sheet = workbook.worksheets.add(spec.name);
  applyMatrix(sheet, spec.title, spec.rows);
}

await fs.mkdir(outputDir, { recursive: true });

for (const spec of ["4_FACT_LABOUR", ...sheets.map(s => s.name)]) {
  const preview = await workbook.render({ sheetName: spec, range: "A1:G16", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${spec}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const verify = await workbook.inspect({
  kind: "table",
  sheetId: "7_FACT_INVOICE_LINE",
  range: "A1:G14",
  include: "values",
  tableMaxRows: 14,
  tableMaxCols: 7,
});
console.log(verify.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
