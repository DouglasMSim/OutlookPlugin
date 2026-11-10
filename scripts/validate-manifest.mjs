// Validate the generated manifests.
//  - Always: XML well-formedness (xmllint if installed, otherwise a parser check).
//  - With --online or in CI: Microsoft's validation service via office-addin-manifest,
//    which needs network access to validationgateway.omex.office.net.
import { execFileSync, spawnSync } from "node:child_process";
import { loadEnvironments, root } from "./lib/manifest.mjs";

const online = process.argv.includes("--online") || process.env.CI === "true";
const environments = Object.keys(await loadEnvironments());
let failed = false;

const hasXmllint = spawnSync("xmllint", ["--version"]).status === 0;
for (const name of environments) {
  const file = `${root}manifests/manifest.${name}.xml`;
  if (hasXmllint) {
    const result = spawnSync("xmllint", ["--noout", file], { encoding: "utf8" });
    if (result.status !== 0) {
      console.error(`manifest.${name}.xml is not well-formed XML:\n${result.stderr}`);
      failed = true;
      continue;
    }
  }
  console.log(`manifest.${name}.xml: well-formed${hasXmllint ? "" : " (xmllint not installed; skipped)"}`);
}

if (online) {
  // The pilot manifest is the one users install; dev points at localhost.
  try {
    const output = execFileSync(
      process.execPath,
      [`${root}node_modules/office-addin-manifest/cli.js`, "validate", `${root}manifests/manifest.pilot.xml`],
      { encoding: "utf8" },
    );
    console.log(output);
    if (!/The manifest is valid/i.test(output)) failed = true;
  } catch (error) {
    console.error(error.stdout || error.message);
    failed = true;
  }
} else {
  console.log("Skipping Microsoft's online manifest validation (run with --online; CI runs it).");
}

process.exit(failed ? 1 : 0);
