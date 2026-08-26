import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_data_center_.xlsx";
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const overview = await workbook.inspect({
  kind: "sheet,table,computedStyle",
  range: "A1:G12",
  maxChars: 7000,
  tableMaxRows: 12,
  tableMaxCols: 7,
});
console.log(overview.ndjson);
