import fs from "node:fs/promises";
const dir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/ppt_kpi_edit";
let issues = [];
for (let i = 1; i <= 4; i++) {
  const path = `${dir}/final-slide-${i}.layout.json`;
  let data;
  try { data = JSON.parse(await fs.readFile(path, "utf8")); } catch { continue; }
  const nodes = [];
  const walk = (obj) => {
    if (!obj || typeof obj !== "object") return;
    if (obj.bbox || obj.bounds || obj.position) nodes.push(obj);
    for (const value of Object.values(obj)) {
      if (Array.isArray(value)) value.forEach(walk);
      else walk(value);
    }
  };
  walk(data);
  for (const node of nodes) {
    const b = node.bbox || node.bounds || node.position;
    if (!b) continue;
    const left = Array.isArray(b) ? b[0] : b.left;
    const top = Array.isArray(b) ? b[1] : b.top;
    const width = Array.isArray(b) ? b[2] : b.width;
    const height = Array.isArray(b) ? b[3] : b.height;
    if ([left, top, width, height].some(v => typeof v !== "number")) continue;
    if (left < -1 || top < -1 || left + width > 1281 || top + height > 721) {
      issues.push({ slide: i, name: node.name, left, top, width, height });
    }
  }
}
console.log(JSON.stringify({ issues: issues.length, sample: issues.slice(0, 10) }, null, 2));
