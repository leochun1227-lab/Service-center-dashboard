import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const outDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/ppt_kpi_edit";
const finalPath = `${outDir}/service_centre_kpi_formulas_3_pages_next_steps.pptx`;
const presentation = await PresentationFile.importPptx(await FileBlob.load(finalPath));
const inspect = await presentation.inspect({ kind: "slide,textbox,shape", search: "Relationship", maxChars: 12000 });
const records = inspect.ndjson.trim().split(/\n+/).filter(Boolean).map(line => JSON.parse(line));
const slideRec = records.find(r => r.kind === "slide" && r.slide === 3) || (await presentation.inspect({ kind: "slide", maxChars: 4000 })).ndjson.trim().split(/\n+/).map(line => JSON.parse(line)).find(r => r.slide === 3);
const slide = presentation.resolve(slideRec.id);
for (const rec of records) {
  if (rec.slide === 3 && rec.kind === "textbox" && /Relationship check/.test(rec.text || "")) {
    const t = presentation.resolve(rec.id);
    t.position = { left: 86.4, top: 515, width: 640, height: 48 };
    t.text.style = { fontSize: 11, color: "#172033", bold: true };
  }
}
const all = await presentation.inspect({ kind: "shape", maxChars: 30000 });
const shapeRecords = all.ndjson.trim().split(/\n+/).filter(Boolean).map(line => JSON.parse(line));
const whiteBox = shapeRecords.find(r => r.slide === 3 && r.bbox && Math.abs(r.bbox[0] - 120) < 2 && Math.abs(r.bbox[1] - 513.6) < 5);
if (whiteBox) {
  const box = presentation.resolve(whiteBox.id);
  box.position = { left: 78, top: 506, width: 660, height: 68 };
}
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save(finalPath);
const finalPresentation = await PresentationFile.importPptx(await FileBlob.load(finalPath));
for (const [index, slide] of finalPresentation.slides.items.entries()) {
  await writeBlob(`${outDir}/final-slide-${index + 1}.png`, await finalPresentation.export({ slide, format: "png", scale: 1 }));
}
await writeBlob(`${outDir}/final-montage.webp`, await finalPresentation.export({ format: "webp", montage: true, scale: 1 }));
