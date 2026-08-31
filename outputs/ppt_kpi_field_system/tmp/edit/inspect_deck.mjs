import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "C:/Users/Leo.Li/Desktop/service_centre_kpi_formulas_3_pages_next_steps.pptx";
const outDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/ppt_kpi_field_system/tmp/edit";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
await fs.mkdir(outDir, { recursive: true });

const inspect = await presentation.inspect({
  kind: "slide,textbox,shape,table,layout",
  include: "id,slide,name,bbox,textPreview,text,isPlaceholder",
  maxChars: 50000,
});
await fs.writeFile(`${outDir}/inspect.ndjson`, inspect.ndjson, "utf8");

for (const [index, slide] of presentation.slides.items.entries()) {
  const n = String(index + 1).padStart(2, "0");
  await writeBlob(`${outDir}/slide-${n}.png`, await presentation.export({ slide, format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${outDir}/slide-${n}.layout.json`, await layout.text(), "utf8");
}

await writeBlob(`${outDir}/montage.webp`, await presentation.export({ format: "webp", montage: true, scale: 1 }));
