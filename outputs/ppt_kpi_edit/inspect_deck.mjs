import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

const source = "C:/Users/Leo.Li/Desktop/service_centre_kpi_formulas_3_pages.pptx";
const out = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/ppt_kpi_edit/inspect";
await fs.mkdir(out, { recursive: true });
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const inspect = await presentation.inspect({ kind: "deck,slide,textbox,shape,table,layout", maxChars: 20000 });
await fs.writeFile(`${out}/inspect.ndjson`, inspect.ndjson, "utf8");
console.log(inspect.ndjson);
let i = 0;
for (const slide of presentation.slides.items) {
  i += 1;
  await writeBlob(`${out}/slide-${i}.png`, await presentation.export({ slide, format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(`${out}/slide-${i}.layout.json`, await layout.text(), "utf8");
}
