import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "C:/Users/Leo.Li/Desktop/service_center_dashboard.xlsx";
const outputDir = "outputs/field_supplement_work";

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table,region",
  maxChars: 10000,
  tableMaxRows: 8,
  tableMaxCols: 12,
  tableMaxCellChars: 120,
});
console.log(summary.ndjson);

await fs.mkdir(outputDir, { recursive: true });
for (const sheet of workbook.worksheets.items) {
  try {
    const preview = await workbook.render({
      sheetName: sheet.name,
      autoCrop: "all",
      scale: 1,
      format: "png",
    });
    await fs.writeFile(
      `${outputDir}/${sheet.name.replace(/[\\/:*?"<>|]/g, "_")}.png`,
      new Uint8Array(await preview.arrayBuffer()),
    );
  } catch (error) {
    console.error(`Render failed for ${sheet.name}: ${error.message}`);
  }
}
