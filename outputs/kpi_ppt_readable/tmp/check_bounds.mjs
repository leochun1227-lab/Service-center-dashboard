import fs from "node:fs/promises";

const dir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/kpi_ppt_readable";
const files = (await fs.readdir(dir)).filter((f) => /^slide-\d+\.layout\.json$/.test(f));
const issues = [];
for (const file of files) {
  const json = JSON.parse(await fs.readFile(`${dir}/${file}`, "utf8"));
  const frame = json.slide.frame;
  for (const el of json.elements ?? []) {
    const b = el.bbox;
    if (!b) continue;
    const [x, y, w, h] = b;
    if (x < -1 || y < -1 || x + w > frame.width + 1 || y + h > frame.height + 1) {
      issues.push({ file, name: el.name, text: el.textPreview ?? el.text, bbox: b });
    }
  }
}
console.log(JSON.stringify({ files: files.length, issues }, null, 2));
