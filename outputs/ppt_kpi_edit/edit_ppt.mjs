import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const source = "C:/Users/Leo.Li/Desktop/service_centre_kpi_formulas_3_pages.pptx";
const outDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/ppt_kpi_edit";
const finalPath = `${outDir}/service_centre_kpi_formulas_3_pages_next_steps.pptx`;
await fs.mkdir(outDir, { recursive: true });

const presentation = await PresentationFile.importPptx(await FileBlob.load(source));

const snapshot = await presentation.inspect({
  kind: "slide,textbox,shape",
  include: "id,slide,name,bbox,text",
  maxChars: 60000,
});
const records = snapshot.ndjson.trim().split(/\n+/).map(line => JSON.parse(line));

const slideByNumber = new Map();
for (const rec of records) {
  if (rec.kind === "slide") slideByNumber.set(rec.slide, rec.id);
}

const columns = {
  kpi: { left: 86.4, width: 185 },
  formula: { left: 285, width: 460 },
  extras: [
    { title: "Next Step", left: 760, width: 100 },
    { title: "Priority", left: 864, width: 62 },
    { title: "Alternative", left: 930, width: 112 },
    { title: "Due date", left: 1046, width: 74 },
    { title: "Responsible by", left: 1124, width: 88 },
  ],
};
const gridRight = 1212;
const gridLineColor = "#D8DEE8";
const headerColor = "#18233A";
const accentColor = "#0F8F86";

function textStyle(size, opts = {}) {
  return {
    fontSize: size,
    color: opts.color || "#172033",
    bold: !!opts.bold,
    alignment: opts.alignment || "left",
  };
}

function addText(slide, text, left, top, width, height, style) {
  const box = slide.shapes.add({
    geometry: "textbox",
    position: { left, top, width, height },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  box.text = text;
  box.text.style = style;
  return box;
}

function addLine(slide, x1, y1, x2, y2, color = gridLineColor, width = 1) {
  slide.shapes.add({
    geometry: "line",
    position: { left: x1, top: y1, width: x2 - x1, height: y2 - y1 },
    fill: "none",
    line: { style: "solid", fill: color, width },
  });
}

for (const slideNo of [2, 3, 4]) {
  const slide = presentation.resolve(slideByNumber.get(slideNo));
  const textBoxes = records.filter(r => r.kind === "textbox" && r.slide === slideNo && r.bbox && r.bbox[1] > 120 && r.bbox[1] < 660);
  const labels = textBoxes.filter(r => r.bbox[0] < 360).sort((a, b) => a.bbox[1] - b.bbox[1]);
  const formulas = textBoxes.filter(r => r.bbox[0] >= 360).sort((a, b) => a.bbox[1] - b.bbox[1]);
  const allRows = [...labels, ...formulas];
  const rowTops = labels.map(r => r.bbox[1]);
  const headerY = Math.max(110, Math.min(...rowTops) - 25);
  const gridTop = headerY + 22;
  const gridBottom = Math.max(...allRows.map(r => r.bbox[1] + r.bbox[3]), 630);

  for (const rec of labels) {
    const shape = presentation.resolve(rec.id);
    shape.position = { left: columns.kpi.left, top: rec.bbox[1], width: columns.kpi.width, height: rec.bbox[3] };
    shape.text.style = textStyle(13, { bold: true });
  }
  for (const rec of formulas) {
    const shape = presentation.resolve(rec.id);
    shape.position = { left: columns.formula.left, top: rec.bbox[1], width: columns.formula.width, height: rec.bbox[3] };
    shape.text.style = textStyle(10.7, { color: "#172033" });
  }

  addText(slide, "KPI", columns.kpi.left, headerY, columns.kpi.width, 18, textStyle(9.5, { bold: true, color: headerColor }));
  addText(slide, "Formula / Source", columns.formula.left, headerY, columns.formula.width, 18, textStyle(9.5, { bold: true, color: headerColor }));
  for (const col of columns.extras) {
    addText(slide, col.title, col.left, headerY, col.width, 20, textStyle(8.3, { bold: true, color: headerColor, alignment: "center" }));
  }

  const separators = [
    columns.formula.left - 12,
    columns.extras[0].left - 10,
    ...columns.extras.slice(1).map(c => c.left - 4),
    gridRight,
  ];
  for (const x of separators) addLine(slide, x, gridTop, x, gridBottom, gridLineColor, 0.65);
  addLine(slide, 69.12, headerY + 20, gridRight, headerY + 20, accentColor, 1.1);
}

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(finalPath);

const finalPresentation = await PresentationFile.importPptx(await FileBlob.load(finalPath));
for (const [index, slide] of finalPresentation.slides.items.entries()) {
  await writeBlob(`${outDir}/final-slide-${index + 1}.png`, await finalPresentation.export({ slide, format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${outDir}/final-slide-${index + 1}.layout.json`, await layout.text(), "utf8");
}
await writeBlob(`${outDir}/final-montage.webp`, await finalPresentation.export({ format: "webp", montage: true, scale: 1 }));
console.log(finalPath);
