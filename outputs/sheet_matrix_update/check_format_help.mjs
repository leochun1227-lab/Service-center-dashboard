import { Workbook } from "@oai/artifact-tool";

const workbook = Workbook.create();
console.log(workbook.help("*", {
  search: "orientation|textRotation|rotation|vertical|wrapText|rowHeight|columnWidth",
  include: "index,examples,notes",
  maxChars: 6000,
}).ndjson);
