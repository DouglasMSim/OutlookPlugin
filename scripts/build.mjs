// Build the add-in into dist/. Usage: node scripts/build.mjs [--env pilot|dev]
import { cp, mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { build } from "esbuild";
import { renderManifest, root } from "./lib/manifest.mjs";

const envIndex = process.argv.indexOf("--env");
const envName = envIndex === -1 ? "pilot" : process.argv[envIndex + 1];
const dist = `${root}dist`;

await rm(dist, { recursive: true, force: true });
await mkdir(`${dist}/assets`, { recursive: true });

await build({
  entryPoints: {
    commands: `${root}src/commands/commands.ts`,
    dialog: `${root}src/dialog/dialog.ts`,
    taskpane: `${root}src/pages/taskpane.ts`,
  },
  outdir: dist,
  bundle: true,
  format: "iife",
  // Outlook runs add-ins in WebView2 / Edge or the browser; ES2017 keeps older webviews working.
  target: "es2017",
  minify: envName !== "dev",
  sourcemap: envName === "dev",
  legalComments: "none",
  logLevel: "warning",
});

for (const file of await readdir(`${root}src/pages`)) {
  if (file.endsWith(".html") || file.endsWith(".css")) await cp(`${root}src/pages/${file}`, `${dist}/${file}`);
}
for (const file of await readdir(`${root}assets`)) {
  if (file.endsWith(".png")) await cp(`${root}assets/${file}`, `${dist}/assets/${file}`);
}

await writeFile(`${dist}/manifest.xml`, await renderManifest(envName));
// GitHub Pages: serve files as-is.
await writeFile(`${dist}/.nojekyll`, "");

console.log(`Built dist/ for "${envName}".`);
