import fs from "node:fs/promises";
import { chromium } from "playwright";

const outputDir = "C:/Users/Leo.Li/Documents/GitHub/Service-center-dashboard/outputs/service_centre_kpi_next_steps/nav_check";
await fs.mkdir(outputDir, { recursive: true });

const browser = await chromium.launch({
  executablePath: "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  headless: true,
});
const page = await browser.newPage({ viewport: { width: 1440, height: 950 } });
const errors = [];
page.on("pageerror", error => errors.push(error.message));
page.on("console", message => {
  if (message.type() === "error") errors.push(message.text());
});

await page.goto("http://127.0.0.1:8780/overview.html", { waitUntil: "networkidle" });
await page.screenshot({ path: `${outputDir}/overview.png`, fullPage: true });

for (const pageId of ["service-orders", "technician-labour", "financial-performance", "quality-customer"]) {
  await page.click(`[data-page="${pageId}"]`);
  await page.waitForTimeout(250);
  await page.screenshot({ path: `${outputDir}/${pageId}.png`, fullPage: true });
}

await page.setViewportSize({ width: 390, height: 900 });
await page.click('[data-page="service-orders"]');
await page.waitForTimeout(250);
await page.screenshot({ path: `${outputDir}/service-orders-mobile.png`, fullPage: true });

const state = await page.evaluate(() => ({
  active: document.querySelector(".nav-item.active")?.textContent.trim(),
  title: document.getElementById("pageTitle")?.textContent,
  moduleHidden: document.getElementById("moduleView")?.hidden,
  overviewHidden: document.getElementById("overviewView")?.hidden,
  cardCount: document.querySelectorAll(".module-card").length,
  tbcBoxes: document.querySelectorAll(".tbc-box, .module-empty").length,
}));

await browser.close();

console.log(JSON.stringify({ state, errors }, null, 2));
