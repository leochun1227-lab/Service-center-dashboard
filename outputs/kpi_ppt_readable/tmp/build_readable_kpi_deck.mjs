import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const outDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/kpi_ppt_readable";
const finalPptx = `${outDir}/service_centre_kpi_next_steps_readable.pptx`;

const W = 1600;
const H = 900;
const C = {
  bg: "#F8FAFC",
  ink: "#111827",
  muted: "#64748B",
  teal: "#0F766E",
  tealSoft: "#CCFBF1",
  line: "#CBD5E1",
  pale: "#FFFFFF",
  amber: "#F59E0B",
  red: "#DC2626",
  green: "#047857",
  blue: "#2563EB",
};

const data = [
  {
    section: "Financial",
    kpi: "Average Repair Order Amount",
    target: "Benchmark: confirm with Finance",
    formula: "Total Repair Order Revenue / Number of Repair Orders",
    source: "SAP Invoice + C4C",
    field: "Field in System: ____________________",
    next: "Confirm SAP invoice revenue field and match it with C4C repair order count.",
    alt: "Alternative: use C4C quote / invoice amount as proxy.",
    owner: "Owner: Finance / SAP + C4C owner",
    priority: "High",
  },
  {
    section: "Financial",
    kpi: "Effective Labor Rate",
    target: "Target: >90% of posted rate",
    formula: "Total Labor Revenue / Labor Hours Sold",
    source: "SAP Invoice + C4C Labour",
    field: "Field in System: ____________________",
    next: "Confirm labour revenue and labour hours sold source.",
    alt: "Alternative: use claimed hours if sold hours are not ready.",
    owner: "Owner: Finance / Service Managers",
    priority: "High",
  },
  {
    section: "Financial",
    kpi: "Gross Profit Margin",
    target: "Target: 58% to 65%",
    formula: "(Parts Revenue + Labor Revenue - Parts Cost - Labor Cost) / Revenue",
    source: "SAP Invoice + SAP PO + HR / Finance",
    field: "Field in System: ____________________",
    next: "Confirm revenue, cost and expense allocation rules.",
    alt: "Alternative: report gross margin before full expense allocation.",
    owner: "Owner: Finance / Mary",
    priority: "High",
  },
  {
    section: "Financial",
    kpi: "Service Gross %",
    target: "Target: 65% to 75%",
    formula: "(Labor Sales - Direct Labor Cost) / Labor Sales",
    source: "SAP Invoice + HR / Payroll",
    field: "Field in System: ____________________",
    next: "Confirm labour sales and direct labour cost source.",
    alt: "Alternative: overview-only service gross until People ID is mapped.",
    owner: "Owner: Finance / HR",
    priority: "High",
  },
  {
    section: "Financial",
    kpi: "Parts Gross %",
    target: "Target: 35% to 45%",
    formula: "(Parts Sales - Parts Cost) / Parts Sales",
    source: "SAP SO Invoice + SAP PO Invoice",
    field: "Field in System: ____________________",
    next: "Confirm parts sales and parts cost fields.",
    alt: "Alternative: use summary parts revenue / cost before line-level mapping.",
    owner: "Owner: Finance / SAP owner",
    priority: "Medium",
  },
  {
    section: "Financial",
    kpi: "Absorption Rate",
    target: "Target: 100% to 130%",
    formula: "(Service Gross Profit + Parts Gross Profit) / Operating Expenses",
    source: "Finance P&L / SAP",
    field: "Field in System: ____________________",
    next: "Confirm service operating expense allocation and cost centre split.",
    alt: "Alternative: track after-service gross profit against estimated cost.",
    owner: "Owner: Finance",
    priority: "Medium",
  },
  {
    section: "Productivity",
    kpi: "Technician Productivity",
    target: "Target: 110% to 130%",
    formula: "Sold / Claimed Hours / Available Hours",
    source: "C4C Labour + Technician roster",
    field: "Field in System: ____________________",
    next: "Set up technician list, team and available hours.",
    alt: "Alternative: location-level claimed hours until technician setup is live.",
    owner: "Owner: Service Managers / Andy He",
    priority: "High",
  },
  {
    section: "Productivity",
    kpi: "Technician Efficiency",
    target: "Target: >110%",
    formula: "Sold / Claimed Hours / Actual Productive Hours",
    source: "C4C Work Orders + Labour records",
    field: "Field in System: ____________________",
    next: "Assign every work order to a technician and record actual work hours.",
    alt: "Alternative: claim hours by location if productive hours are unstable.",
    owner: "Owner: Service Managers",
    priority: "High",
  },
  {
    section: "Productivity",
    kpi: "Technician Utilisation",
    target: "Target: >90%",
    formula: "Productive / Claimed Hours / Available Hours",
    source: "C4C Labour + Technician calendar",
    field: "Field in System: ____________________",
    next: "Confirm roster / calendar available hours and productive hour logic.",
    alt: "Alternative: claimed hours / working-day capacity.",
    owner: "Owner: Service Managers / C4C owner",
    priority: "High",
  },
  {
    section: "Operations",
    kpi: "Comeback / Rework Rate",
    target: "Target: <5%",
    formula: "Rework Tickets / Completed Repair Orders",
    source: "C4C Tickets",
    field: "Field in System: ____________________",
    next: "Rename comeback to rework and define the rework ticket creation rule.",
    alt: "Alternative: detect same vehicle / same issue repeat repairs.",
    owner: "Owner: Service Managers / Lee",
    priority: "High",
  },
  {
    section: "Operations",
    kpi: "Fixed First Time Right",
    target: "Target: >95%",
    formula: "1 - Rework Rate",
    source: "C4C Tickets",
    field: "Field in System: ____________________",
    next: "Use rework tickets as the failure signal.",
    alt: "Alternative: same-issue repeat repair detection.",
    owner: "Owner: Dashboard owner / Service Managers",
    priority: "High",
  },
  {
    section: "Operations",
    kpi: "Quotation Approval Rate",
    target: "Target: >70%",
    formula: "Approved Quotes / Raised Quotes",
    source: "C4C Quote / Repair Orders",
    field: "Field in System: ____________________",
    next: "Standardise quote statuses into raised, approved and rejected.",
    alt: "Alternative: use open repair order quote status.",
    owner: "Owner: C4C owner / Service Managers",
    priority: "High",
  },
  {
    section: "Operations",
    kpi: "Service Appointment Lead Time",
    target: "Target: 3 to 7 days",
    formula: "Appointment Date - Booking Request Date",
    source: "C4C Calendar / Tickets",
    field: "Field in System: ____________________",
    next: "Confirm requested service date capture; short term use ticket created date.",
    alt: "Alternative: created date to appointment calendar date.",
    owner: "Owner: Andy He / Service Managers",
    priority: "High",
  },
  {
    section: "Operations",
    kpi: "Customer Retention Rate",
    target: "Target: >90%",
    formula: "Returning Sold Vans / Sold Vans in Service Scope",
    source: "C4C Service + Sales / Customer data",
    field: "Field in System: ____________________",
    next: "Define customer return logic using sold van / local service records.",
    alt: "Alternative: return service rate for sold vans.",
    owner: "Owner: Business / Service",
    priority: "Medium",
  },
  {
    section: "Operations",
    kpi: "Retail Repair Order Growth Rate",
    target: "Target: 3% to 6% YOY",
    formula: "(Current Period Retail ROs - Prior Period Retail ROs) / Prior Period Retail ROs",
    source: "C4C Repair Orders",
    field: "Field in System: ____________________",
    next: "Use month-on-month or quarterly trend before full YOY baseline is available.",
    alt: "Alternative: MoM / QoQ trend until 2027.",
    owner: "Owner: BI / Dashboard owner",
    priority: "Medium",
  },
];

function addBox(slide, cfg) {
  return slide.shapes.add({
    geometry: cfg.geometry ?? "rect",
    name: cfg.name,
    position: cfg.position,
    fill: cfg.fill ?? "none",
    line: cfg.line ?? { style: "solid", fill: "none", width: 0 },
  });
}

function addText(slide, text, x, y, w, h, style = {}) {
  const shape = addBox(slide, {
    geometry: "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    fontSize: style.fontSize ?? 18,
    bold: style.bold ?? false,
    color: style.color ?? C.ink,
    alignment: style.alignment ?? "left",
    italic: style.italic ?? false,
  };
  return shape;
}

function addHeader(slide, title, kicker, pageNo) {
  slide.background.fill = C.bg;
  addBox(slide, { position: { left: 0, top: 0, width: W, height: 18 }, fill: C.teal });
  addText(slide, kicker.toUpperCase(), 80, 50, 500, 28, { fontSize: 16, bold: true, color: C.teal });
  addText(slide, title, 80, 78, 1200, 62, { fontSize: 44, bold: true, color: C.ink });
  addBox(slide, { position: { left: 80, top: 150, width: 1440, height: 2 }, fill: C.line });
  addText(slide, String(pageNo).padStart(2, "0"), 1460, 62, 70, 34, { fontSize: 20, bold: true, color: C.muted, alignment: "right" });
}

function priorityColor(priority) {
  return priority === "High" ? C.red : priority === "Medium" ? C.amber : C.green;
}

function addMetric(slide, item, x, y, w, h) {
  addBox(slide, { position: { left: x, top: y, width: 7, height: h }, fill: priorityColor(item.priority) });
  addText(slide, item.kpi, x + 22, y - 2, w - 24, 34, { fontSize: 24, bold: true });
  addText(slide, item.target, x + 22, y + 34, 360, 24, { fontSize: 16, color: C.teal, bold: true });
  addText(slide, `Source: ${item.source}`, x + 400, y + 34, w - 420, 24, { fontSize: 16, color: C.muted });

  const midY = y + 72;
  addBox(slide, { position: { left: x + 22, top: midY, width: w - 44, height: 1 }, fill: "#E5E7EB" });
  addText(slide, "Formula", x + 22, midY + 14, 130, 24, { fontSize: 16, bold: true, color: C.muted });
  addText(slide, item.formula, x + 155, midY + 14, w - 180, 48, { fontSize: 17, color: C.ink });

  addText(slide, "Field", x + 22, midY + 72, 130, 24, { fontSize: 16, bold: true, color: C.muted });
  addText(slide, item.field, x + 155, midY + 72, w - 180, 30, { fontSize: 17, color: C.ink });

  addText(slide, "Next Step", x + 22, midY + 116, 130, 24, { fontSize: 16, bold: true, color: C.muted });
  addText(slide, item.next, x + 155, midY + 116, w - 180, 50, { fontSize: 17, color: C.ink });

  addText(slide, item.alt, x + 155, y + h - 52, Math.floor((w - 190) * 0.6), 34, { fontSize: 15, color: C.muted });
  addText(slide, item.owner, x + w - 430, y + h - 52, 385, 34, { fontSize: 15, color: C.muted });
}

function addMetricSlide(presentation, title, kicker, items, pageNo) {
  const slide = presentation.slides.add();
  addHeader(slide, title, kicker, pageNo);
  const usableX = 90;
  const usableW = 1420;
  const h = items.length === 1 ? 420 : 300;
  const startY = items.length === 1 ? 245 : 205;
  const gap = items.length === 1 ? 0 : 325;
  items.forEach((item, i) => addMetric(slide, item, usableX, startY + i * gap, usableW, h));
  return slide;
}

function addSummarySlide(presentation) {
  const slide = presentation.slides.add();
  addHeader(slide, "KPI build needs system fields before automation", "Service Centre Dashboard", 2);
  addText(
    slide,
    "This version separates formula definition from implementation readiness, so each KPI has enough room for the field name, next action, owner and fallback.",
    90,
    180,
    1330,
    70,
    { fontSize: 24, color: C.ink },
  );
  const groups = [
    ["Financial", "SAP invoice, SAP PO, HR/payroll and P&L allocation need confirmation before margin KPIs can be automated."],
    ["Productivity", "Technician list, assignment, roster and actual / claimed hours are the core blockers."],
    ["Operations", "Rework rules, quote statuses and requested service date capture unlock the customer and workflow KPIs."],
  ];
  groups.forEach(([head, body], i) => {
    const x = 100 + i * 475;
    addBox(slide, { position: { left: x, top: 320, width: 405, height: 300 }, fill: C.pale, line: { style: "solid", fill: C.line, width: 1 } });
    addBox(slide, { position: { left: x, top: 320, width: 405, height: 8 }, fill: i === 0 ? C.red : i === 1 ? C.blue : C.teal });
    addText(slide, head, x + 28, 360, 340, 38, { fontSize: 28, bold: true });
    addText(slide, body, x + 28, 420, 345, 130, { fontSize: 20, color: C.ink });
  });
}

function addTitleSlide(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = C.bg;
  addBox(slide, { position: { left: 0, top: 0, width: W, height: 26 }, fill: C.teal });
  addText(slide, "SERVICE CENTRE KPI READINESS", 90, 120, 600, 30, { fontSize: 18, bold: true, color: C.teal });
  addText(slide, "KPI Formulas and Next Steps", 90, 172, 1120, 68, { fontSize: 56, bold: true });
  addText(slide, "Readable PPT version with larger type, clearer spacing and Field in System before each implementation step.", 92, 260, 1060, 78, { fontSize: 25, color: C.muted });
  addBox(slide, { position: { left: 90, top: 390, width: 1240, height: 2 }, fill: C.line });
  const items = ["Financial source confirmation", "Technician setup and labour capture", "Operations rules for rework, quote and appointment KPIs"];
  items.forEach((item, i) => {
    addBox(slide, { position: { left: 110, top: 445 + i * 78, width: 16, height: 16 }, fill: i === 0 ? C.red : i === 1 ? C.blue : C.teal });
    addText(slide, item, 150, 428 + i * 78, 1000, 46, { fontSize: 28, color: C.ink });
  });
}

function addClosingSlide(presentation, pageNo) {
  const slide = presentation.slides.add();
  addHeader(slide, "Decision points before dashboard build", "Implementation checklist", pageNo);
  const checklist = [
    "Fill exact Field in System names for each KPI.",
    "Confirm whether missing financial fields are SAP, C4C, HR/payroll or manual P&L inputs.",
    "Finish technician setup: people ID, team, assigned work orders and available hours.",
    "Approve interim alternatives for metrics that cannot be automated at launch.",
    "Use traffic-light status on the dashboard once targets and sources are signed off.",
  ];
  checklist.forEach((text, i) => {
    const y = 210 + i * 96;
    addBox(slide, { position: { left: 110, top: y + 8, width: 20, height: 20 }, fill: C.teal });
    addText(slide, text, 160, y, 1220, 52, { fontSize: 27, color: C.ink });
  });
}

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

await fs.mkdir(outDir, { recursive: true });
const presentation = Presentation.create({ slideSize: { width: W, height: H } });

addTitleSlide(presentation);
addSummarySlide(presentation);
addMetricSlide(presentation, "Financial revenue KPIs need invoice linkage", "Financial 1", data.slice(0, 2), 3);
addMetricSlide(presentation, "Financial margin KPIs need labour cost mapping", "Financial 2", data.slice(2, 4), 4);
addMetricSlide(presentation, "Parts and absorption need cost allocation", "Financial 3", data.slice(4, 6), 5);
addMetricSlide(presentation, "Technician productivity depends on setup", "Productivity 1", data.slice(6, 8), 6);
addMetricSlide(presentation, "Technician utilisation needs calendar capacity", "Productivity 2", data.slice(8, 9), 7);
addMetricSlide(presentation, "Rework KPIs need a clear ticket rule", "Operations 1", data.slice(9, 11), 8);
addMetricSlide(presentation, "Quote and appointment KPIs need standard dates", "Operations 2", data.slice(11, 13), 9);
addMetricSlide(presentation, "Customer and growth KPIs need stable history", "Operations 3", data.slice(13, 15), 10);
addClosingSlide(presentation, 11);

for (const [index, slide] of presentation.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(`${outDir}/${stem}.png`, await presentation.export({ slide, format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${outDir}/${stem}.layout.json`, await layout.text(), "utf8");
}

await writeBlob(`${outDir}/montage.webp`, await presentation.export({ format: "webp", montage: true, scale: 1 }));
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(finalPptx);
console.log(finalPptx);
