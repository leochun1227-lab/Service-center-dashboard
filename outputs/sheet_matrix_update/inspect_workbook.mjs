import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_dashboard_horizontal_value_matrix_updated.xlsx";
const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/sheet_matrix_update";

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheets = await workbook.inspect({
  kind: "workbook,sheet,table,region",
  maxChars: 9000,
  tableMaxRows: 12,
  tableMaxCols: 14,
  tableMaxCellChars: 120,
});
console.log(sheets.ndjson);

const first = workbook.worksheets.getItemAt(0);
const firstName = first.name;
const region = await workbook.inspect({
  kind: "region",
  sheetId: firstName,
  range: "A1:Z40",
  maxChars: 8000,
});
console.log(region.ndjson);

const preview = await workbook.render({
  sheetName: firstName,
  range: "A1:Z35",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/before_first_sheet.png`, new Uint8Array(await preview.arrayBuffer()));
console.log(`${outputDir}/before_first_sheet.png`);
