// Render assets/icon.svg to the PNG sizes the manifest and install page use.
// Usage: node scripts/icons.mjs [--preview <file.png>]
import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";

const root = fileURLToPath(new URL("..", import.meta.url));
const SIZES = [16, 20, 24, 32, 40, 48, 64, 80, 128];

const svg = await readFile(`${root}assets/icon.svg`, "utf8");
const browser = await chromium.launch();
try {
  const page = await browser.newPage({ deviceScaleFactor: 1 });
  for (const size of SIZES) {
    await page.setViewportSize({ width: size, height: size });
    const sized = svg.replace(/width="\d+" height="\d+"/, `width="${size}" height="${size}"`);
    await page.setContent(
      `<html><body style="margin:0;background:transparent">${sized}</body></html>`,
    );
    const png = await page.locator("svg").screenshot({ omitBackground: true });
    await writeFile(`${root}assets/icon-${size}.png`, png);
  }

  const previewIndex = process.argv.indexOf("--preview");
  if (previewIndex !== -1) {
    let cells = "";
    for (const s of [16, 32, 80, 128]) {
      const data = (await readFile(`${root}assets/icon-${s}.png`)).toString("base64");
      cells += `<img src="data:image/png;base64,${data}" width="${s}" height="${s}">`;
    }
    await page.setViewportSize({ width: 560, height: 300 });
    await page.setContent(
      `<body style="margin:0;font:12px sans-serif">
        <div style="display:flex;gap:24px;align-items:end;padding:20px;background:#fff">${cells}</div>
        <div style="display:flex;gap:24px;align-items:end;padding:20px;background:#292929">${cells}</div>
      </body>`,
    );
    await page.screenshot({ path: process.argv[previewIndex + 1] });
  }
} finally {
  await browser.close();
}
console.log(`Wrote ${SIZES.length} icons.`);
