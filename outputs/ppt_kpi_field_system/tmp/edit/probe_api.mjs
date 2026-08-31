import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const source = "C:/Users/Leo.Li/Desktop/service_centre_kpi_formulas_3_pages_next_steps.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(source));
const inspect = await presentation.inspect({
  kind: "textbox,shape",
  search: "Next Step",
  include: "id,slide,name,bbox,textPreview",
  maxChars: 10000,
});
console.log(inspect.ndjson);
const first = inspect.ndjson.split(/\r?\n/).find((line) => line.includes('"text":"Next Step"'));
const id = JSON.parse(first).id;
const shape = presentation.resolve(id);
console.log(Object.keys(shape));
console.log("position", shape.position);
console.log("text", String(shape.text));
console.log("data keys", Object.keys(shape.data ?? {}));
console.log(JSON.stringify(shape.data, null, 2).slice(0, 2000));
console.log("placement", JSON.stringify(shape.data.placement, null, 2));
console.log("frame", shape.frame);
console.log("pixelRect", shape.pixelRect);
console.log("methods", Object.getOwnPropertyNames(Object.getPrototypeOf(shape)));
console.log("text methods", Object.getOwnPropertyNames(Object.getPrototypeOf(shape.text)));
