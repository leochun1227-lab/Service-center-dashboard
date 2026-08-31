import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "C:/Users/Leo.Li/Desktop/service_centre_kpi_formulas_3_pages_next_steps.pptx";
const outDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/ppt_kpi_field_system";
const finalPptx = `${outDir}/service_centre_kpi_formulas_3_pages_next_steps_field_in_system.pptx`;

const actionSlides = new Set([2, 3, 4]);
const boundaries = [700, 788, 880, 942, 1054, 1126, 1212];
const columns = [
  { label: "Field in System", left: 702, width: 84 },
  { label: "Next Step", left: 790, width: 88 },
  { label: "Priority", left: 882, width: 58 },
  { label: "Alternative", left: 944, width: 108 },
  { label: "Due date", left: 1056, width: 68 },
  { label: "Responsible by", left: 1128, width: 82 },
];

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function recordsFrom(ndjson) {
  return ndjson
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

function slideHeaderTop(records, slideNumber) {
  const next = records.find(
    (r) => r.kind === "textbox" && r.slide === slideNumber && r.text === "Next Step",
  );
  return next?.bbox?.[1] ?? 118;
}

function addHeader(slide, top) {
  const header = slide.shapes.add({
    geometry: "textbox",
    name: "Field in System header",
    position: { left: columns[0].left, top, width: columns[0].width, height: 26 },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  header.text = "Field in System";
  header.text.style = {
    fontSize: 8.5,
    bold: true,
    color: "#1F2937",
    alignment: "center",
  };
}

function styleHeader(shape, fontSize = 9.2) {
  shape.text.style = {
    fontSize,
    bold: true,
    color: "#1F2937",
    alignment: "center",
  };
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
await fs.mkdir(outDir, { recursive: true });

const inspect = await presentation.inspect({
  kind: "slide,textbox,shape",
  include: "id,slide,name,bbox,text",
  maxChars: 80000,
});
const records = recordsFrom(inspect.ndjson);

for (const slideNumber of actionSlides) {
  const slideRecord = records.find((r) => r.kind === "slide" && r.slide === slideNumber);
  const slide = presentation.resolve(slideRecord.id);
  const top = slideHeaderTop(records, slideNumber);

  for (const record of records) {
    if (record.slide !== slideNumber) continue;

    const [left, y, width, height] = record.bbox ?? [];
    if (
      record.kind === "shape" &&
      record.name?.startsWith("Straight Connector") &&
      left > 680 &&
      width < 8 &&
      height > 400
    ) {
      presentation.resolve(record.id).delete();
    }

    if (
      record.kind === "textbox" &&
      left >= 280 &&
      left <= 290 &&
      width >= 440 &&
      width <= 480 &&
      y > 140
    ) {
      const formula = presentation.resolve(record.id);
      formula.frame = { left, top: y, width: 405, height };
    }
  }

  addHeader(slide, top);

  for (const col of columns.slice(1)) {
    const record = records.find(
      (r) => r.kind === "textbox" && r.slide === slideNumber && r.text === col.label,
    );
    if (!record) continue;
    const shape = presentation.resolve(record.id);
    shape.frame = { left: col.left, top, width: col.width, height: 26 };
    styleHeader(shape, col.label === "Responsible by" ? 8.4 : 9.2);
  }

  const gridTop = slideNumber === 2 ? 130 : slideNumber === 3 ? 160 : 140;
  const gridHeight = slideNumber === 2 ? 496 : slideNumber === 3 ? 470 : 490;
  for (const x of boundaries) {
    const line = slide.shapes.add({
      geometry: "rect",
      name: "Action column divider",
      position: { left: x, top: gridTop, width: 1, height: gridHeight },
      fill: "#E2E8F0",
      line: { style: "solid", fill: "#E2E8F0", width: 0 },
    });
    line.sendToBack();
  }
}

for (const [index, slide] of presentation.slides.items.entries()) {
  const n = String(index + 1).padStart(2, "0");
  await writeBlob(`${outDir}/final-slide-${n}.png`, await presentation.export({ slide, format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${outDir}/final-slide-${n}.layout.json`, await layout.text(), "utf8");
}

await writeBlob(`${outDir}/final-montage.webp`, await presentation.export({ format: "webp", montage: true, scale: 1 }));
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(finalPptx);
console.log(finalPptx);
