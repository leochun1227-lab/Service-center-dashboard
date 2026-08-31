import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/service-center-erd-editable-clean.pptx";
const PREVIEW = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/service-center-erd-editable-clean.png";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const W = 2400;
const H = 1580;
const red = "#dd2229";
const blue = "#1976c9";
const bg = "#eef7fd";
const text = "#172033";
const grid = "#dbe3ea";
const sap = "#bd3a2d";
const c4c = "#16834d";

const presentation = Presentation.create({ slideSize: { width: W, height: H } });
const slide = presentation.slides.add();
slide.background.fill = bg;

function box(name, left, top, width, height, fill = "white", line = "#d9e4ee") {
  return slide.shapes.add({
    geometry: "roundRect",
    name,
    position: { left, top, width, height },
    fill,
    line: { style: "solid", fill: line, width: 1.5 },
    borderRadius: 10,
  });
}

function textBox(name, content, left, top, width, height, size, options = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position: { left, top, width, height },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = content;
  shape.text.style = {
    fontSize: size,
    bold: Boolean(options.bold),
    color: options.color || text,
    alignment: options.alignment || "left",
  };
  return shape;
}

function label(name, content, left, top, width, color = red) {
  const circle = slide.shapes.add({
    geometry: "ellipse",
    name: `${name}-badge`,
    position: { left, top: top + 3, width: 30, height: 30 },
    fill: color,
    line: { style: "solid", fill: color, width: 1 },
  });
  circle.text = content.split(" ")[0];
  circle.text.style = { fontSize: 13, bold: true, color: "white", alignment: "center" };
  textBox(name, content.replace(/^\S+\s+/, ""), left + 40, top, width, 36, 18, { color, bold: true });
}

function addRelationship(name, from, to, fromSide, toSide, color = red, labelText = "1:N", labelLeft = 0, labelTop = 0, badge = "") {
  slide.shapes.connect(from, to, {
    kind: "elbow",
    fromSide,
    toSide,
    line: { style: "solid", fill: color, width: 3 },
    head: { type: "triangle", width: "sm", length: "sm" },
  });
  const circle = slide.shapes.add({
    geometry: "ellipse",
    name: `${name}-number`,
    position: { left: labelLeft, top: labelTop, width: 34, height: 28 },
    fill: color,
    line: { style: "solid", fill: color, width: 1 },
  });
  circle.text = badge || name;
  circle.text.style = { fontSize: 10, bold: true, color: "white", alignment: "center" };
  textBox(`${name}-cardinality`, labelText, labelLeft + 42, labelTop + 1, 48, 28, 17, { color, bold: true });
}

function addTable({ id, left, top, width, header, badge, color, rows, note }) {
  const rowH = id === "quote" ? 27 : id === "task" ? 26 : 30;
  const headerH = 52;
  const noteH = note ? 72 : 0;
  const height = headerH + rows.length * rowH + 28;
  const card = box(`${id}-card`, left, top, width, height, "white", color);
  const head = slide.shapes.add({
    geometry: "rect",
    name: `${id}-header`,
    position: { left, top, width, height: headerH },
    fill: color,
    line: { style: "solid", fill: color, width: 1 },
  });
  const badgeShape = slide.shapes.add({
    geometry: "ellipse",
    name: `${id}-badge`,
    position: { left: left + 14, top: top + 12, width: 28, height: 28 },
    fill: "white",
    line: { style: "solid", fill: "white", width: 1 },
  });
  badgeShape.text = badge;
  badgeShape.text.style = { fontSize: 13, bold: true, color, alignment: "center" };
  textBox(`${id}-title`, header, left + 54, top + 12, width - 70, 30, 22, { color: "white", bold: true });

  const values = [["Field", "Source", "Key"], ...rows.map((r) => [r[0], r[1] || "", r[2] || ""])];
  const table = slide.tables.add({
    rows: values.length,
    columns: 3,
    left: left + 14,
    top: top + headerH + 14,
    width: width - 28,
    height: values.length * rowH,
    columnWidths: [250, width - 360, 58],
    values,
  });
  table.borders.assign({ style: "solid", fill: grid, width: 1 });
  table.cells.block({ row: 0, column: 0, rowCount: 1, columnCount: 3 }).assign({
    fill: "#f8fafc",
    textStyle: { fontSize: 13, bold: true, color: "#475569" },
    margins: { left: 6, right: 6, top: 2, bottom: 2 },
  });
  table.cells.block({ row: 1, column: 0, rowCount: rows.length, columnCount: 3 }).assign({
    fill: "white",
    textStyle: { fontSize: 14, color: text },
    margins: { left: 6, right: 6, top: 1, bottom: 1 },
  });
  for (let i = 0; i < rows.length; i += 1) {
    const src = rows[i][1] || "";
    const key = rows[i][2] || "";
    if (src.includes("SAP")) table.getCell(i + 1, 1).text.style = { fontSize: 13, color: sap, bold: true };
    if (src.includes("C4C")) table.getCell(i + 1, 1).text.style = { fontSize: 13, color: c4c, bold: true };
    if (key === "PK") table.getCell(i + 1, 2).fill = "#f6b53f";
    if (key === "FK") table.getCell(i + 1, 2).fill = "#32b56d";
    if (key) table.getCell(i + 1, 2).text.style = { fontSize: 12, bold: true, color: "white", alignment: "center" };
  }

  if (note) {
    const n = box(`${id}-note`, left, top + height + 14, width, noteH, color === blue ? "#eef6ff" : "#fff1f1", "none");
    textBox(`${id}-note-text`, note, left + 16, top + height + 28, width - 32, noteH - 24, 15, { color: color === blue ? "#27557e" : "#8a2525" });
    n.sendToBack();
  }
  return card;
}

textBox("title", "Service Center Data Center", 78, 50, 1100, 56, 44, { bold: true });
textBox("subtitle", "7 tables · dashboard serving model with repair order finance KPI inputs", 80, 112, 1200, 32, 20, { color: "#35506e" });

const legend = box("legend", 1740, 50, 580, 140, "white", "#dce8f2");
textBox("legend-pk", "PK = Primary Key", 1768, 72, 250, 24, 17);
textBox("legend-fk", "FK = Foreign Key", 2032, 72, 250, 24, 17);
textBox("legend-dim", "DIM = Dimension", 1768, 110, 250, 24, 17);
textBox("legend-fact", "FACT = Fact Table", 2032, 110, 250, 24, 17);
textBox("legend-sap", "SAP = finance source", 1768, 150, 250, 24, 17, { color: sap, bold: true });
textBox("legend-c4c", "C4C = service source", 2032, 150, 250, 24, 17, { color: c4c, bold: true });
legend.sendToBack();

const ro = addTable({
  id: "ro",
  left: 80,
  top: 220,
  width: 580,
  header: "1. FACT_RO_FINANCE",
  badge: "F",
  color: red,
  rows: [
    ["ROFinanceID", "", "PK"],
    ["ServiceOrderID", "C4C", "FK"],
    ["InvoiceID", "SAP", "FK"],
    ["OrderRevenue", "", ""],
    ["LaborRevenue", "", ""],
    ["PartsRevenue", "", ""],
    ["LabourCostAmount", "HR Payslip", ""],
    ["PartsCost", "", ""],
    ["SoldHours", "C4C ClaimHours", ""],
    ["PostedLaborRate", "C4C", ""],
  ],
  note: "RO finance fact: finance KPI inputs for ARO, labor rate, gross margin and absorption.",
});

const order = addTable({
  id: "order",
  left: 760,
  top: 220,
  width: 580,
  header: "2. FACT_SERVICE_ORDER",
  badge: "O",
  color: red,
  rows: [
    ["ServiceOrderID", "C4C Ticket ID", "PK"],
    ["VehicleID", "C4C Chassis number", "FK"],
    ["DealerYard", "C4C Warranty Dealer", ""],
    ["ServiceType", "C4C Type", ""],
    ["CreatedDate", "C4C CreatedOn", ""],
    ["BookingRequestDate", "", ""],
    ["AppointmentDateTime", "missing", ""],
    ["StartDate", "missing", ""],
    ["CompletionTime", "missing", ""],
    ["Status", "C4C", ""],
    ["Total QuoteAmount", "C4C Repair Total Amount", ""],
  ],
  note: "Service order fact: main grain for volume, status, lead time, total quote amount and comeback KPIs.",
});

const invoice = addTable({
  id: "invoice",
  left: 1440,
  top: 220,
  width: 580,
  header: "6. FACT_INVOICE",
  badge: "I",
  color: red,
  rows: [
    ["InvoiceID", "SAP", "PK"],
    ["ServiceOrderID", "C4C", "FK"],
    ["InvoiceNo", "SAP / C4C", ""],
    ["InvoiceType", "C4C external/internal", ""],
    ["InvoiceDate", "SAP Billing date", ""],
    ["InvoiceStatus", "SAP not decline/cancel", ""],
    ["InvoiceAmount", "SAP", ""],
    ["LabourRevenueAmount", "tbc", ""],
    ["PartsRevenueAmount", "", ""],
    ["HoursBilled", "SAP", ""],
    ["PaidAmount", "SAP", ""],
  ],
  note: "Invoice fact: header plus margin split for ARO, gross margin, paid amount and absorption KPIs.",
});

const tech = addTable({
  id: "tech",
  left: 80,
  top: 700,
  width: 580,
  header: "5. DIM_TECHNICIAN",
  badge: "T",
  color: blue,
  rows: [
    ["TechnicianID", "C4C assign to", "PK"],
    ["TechnicianName", "C4C assign to", ""],
    ["ServiceCentre", "missing", ""],
  ],
  note: "Technician master: supports productivity, efficiency and utilisation KPI grouping.",
});

const quote = addTable({
  id: "quote",
  left: 760,
  top: 700,
  width: 580,
  header: "2A. FACT_QUOTE",
  badge: "Q",
  color: red,
  rows: [
    ["QuoteID", "C4C", "PK"],
    ["ServiceOrderID", "C4C Ticket ID", "FK"],
    ["IssueID", "C4C Issue ID", ""],
    ["IssuesPosition", "", ""],
    ["IssuesDescription", "", ""],
    ["Subcategory", "", ""],
    ["SubcategoryReason", "", ""],
    ["RepairItem", "", ""],
    ["RepairItemDescription", "", ""],
    ["HoursOfLabour", "", ""],
    ["UnitPrice", "", ""],
    ["PartTotalAmount", "", ""],
    ["QuoteAmount", "", ""],
  ],
  note: "Quote fact: one Service Order can have multiple quote rows.",
});

const labour = addTable({
  id: "labour",
  left: 1440,
  top: 700,
  width: 580,
  header: "4. FACT_LABOUR",
  badge: "L",
  color: red,
  rows: [
    ["LabourID", "", "PK"],
    ["TaskID", "", "FK"],
    ["TechnicianID", "", "FK"],
    ["SoldHours", "C4C ClaimHours", ""],
    ["ProductivityHours", "ActualWorkHours", ""],
    ["AvailableHours", "HR Payslip", ""],
    ["NonProductiveHours", "missing PH", ""],
  ],
  note: "Labour fact: labour hours plus capacity fields for productivity, efficiency and utilisation.",
});

const task = addTable({
  id: "task",
  left: 760,
  top: 1190,
  width: 580,
  header: "3. FACT_SERVICE_TASK",
  badge: "R",
  color: red,
  rows: [
    ["TaskID", "", "PK"],
    ["ServiceOrderID", "C4C", "FK"],
    ["TaskCategory", "C4C internal/external", ""],
    ["RepairItem", "C4C not into use", ""],
    ["ProblemDescription", "C4C", ""],
    ["AssignedTechnicianID", "missing", "FK"],
    ["SoldHours", "C4C", ""],
    ["ActualHours", "missing", ""],
    ["WaitingReason", "", ""],
    ["ReworkFlag", "same chassis + fix desc.", ""],
  ],
  note: "",
});

const merged = box("merged-fields", 80, 1060, 580, 330, "#eef6ff", "#bfd7ee");
textBox("merged-title", "Merged Fields", 112, 1086, 500, 34, 24, { bold: true, color: "#27557e" });
textBox(
  "merged-body",
  "Quote outcome is summarized on Service Order as Total QuoteAmount.\nQuote detail is stored in 2A. FACT_QUOTE with 1:N from Service Order.\nTechnician links directly to FACT_SERVICE_TASK through AssignedTechnicianID.\nFinancial KPI inputs stay in FACT_RO_FINANCE.",
  112,
  1136,
  500,
  190,
  17,
  { color: "#27557e" },
);

addRelationship("1", order, ro, "left", "right", red, "1:N", 674, 360);
addRelationship("6", order, invoice, "right", "left", red, "1:N", 1352, 360);
addRelationship("2A", order, quote, "bottom", "top", red, "1:N", 1350, 650, "2A");
addRelationship("3", quote, task, "bottom", "top", red, "1:N", 1350, 1140);
addRelationship("T", tech, task, "right", "left", blue, "1:N", 674, 1010, "T");
addRelationship("4", task, labour, "right", "bottom", red, "1:N", 1352, 1080);

await fs.mkdir("C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs", { recursive: true });
await writeBlob(PREVIEW, await presentation.export({ slide, format: "png", scale: 0.5 }));
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(OUT);
console.log(JSON.stringify({ OUT, PREVIEW }));
