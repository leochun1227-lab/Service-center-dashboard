import { FileBlob, PresentationFile } from "@oai/artifact-tool";
const path = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/ppt_kpi_edit/service_centre_kpi_formulas_3_pages_next_steps.pptx";
const presentation = await PresentationFile.importPptx(await FileBlob.load(path));
const inspect = await presentation.inspect({ kind: "slide,textbox", search: "Next Step", maxChars: 12000 });
console.log(inspect.ndjson);
