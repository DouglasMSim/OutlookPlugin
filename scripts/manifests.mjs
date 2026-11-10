// Generate manifests/manifest.dev.xml and manifests/manifest.pilot.xml.
// With --check, fail instead of writing if a committed manifest is out of date.
import { readFile, writeFile } from "node:fs/promises";
import { loadEnvironments, renderManifest, root } from "./lib/manifest.mjs";

const check = process.argv.includes("--check");
const environments = await loadEnvironments();
let stale = 0;

for (const name of Object.keys(environments)) {
  const path = `${root}manifests/manifest.${name}.xml`;
  const expected = await renderManifest(name);
  if (check) {
    const actual = await readFile(path, "utf8").catch(() => "");
    if (actual !== expected) {
      console.error(`manifests/manifest.${name}.xml is out of date. Run: npm run manifests`);
      stale += 1;
    }
  } else {
    await writeFile(path, expected);
    console.log(`Wrote manifests/manifest.${name}.xml`);
  }
}

if (stale > 0) process.exit(1);
if (check) console.log("Manifests are up to date.");
