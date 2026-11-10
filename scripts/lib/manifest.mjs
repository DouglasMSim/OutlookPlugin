import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";

export const root = fileURLToPath(new URL("../..", import.meta.url));

export async function loadEnvironments() {
  return JSON.parse(await readFile(`${root}config/environments.json`, "utf8"));
}

/** "1.2.3" -> "1.2.3.0" (manifests need four parts). */
export function manifestVersion(packageVersion) {
  const parts = packageVersion.split(".").map((p) => String(Number.parseInt(p, 10) || 0));
  while (parts.length < 4) parts.push("0");
  return parts.slice(0, 4).join(".");
}

function escapeXml(value) {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export async function renderManifest(envName) {
  const environments = await loadEnvironments();
  const env = environments[envName];
  if (!env) throw new Error(`Unknown environment "${envName}". Known: ${Object.keys(environments).join(", ")}`);
  const origin = env.origin.replace(/\/+$/, "");
  if (!origin.startsWith("https://")) throw new Error(`${envName} origin must use HTTPS: ${origin}`);

  const pkg = JSON.parse(await readFile(`${root}package.json`, "utf8"));
  const template = await readFile(`${root}manifests/manifest.template.xml`, "utf8");
  const values = {
    ID: env.id,
    VERSION: manifestVersion(pkg.version),
    DISPLAY_NAME: env.displayName,
    ORIGIN: origin,
  };
  const rendered = template.replace(/\{\{([A-Z_]+)\}\}/g, (match, key) => {
    if (!(key in values)) throw new Error(`Unknown placeholder ${match}`);
    return escapeXml(values[key]);
  });
  return rendered.replace(/<!--\s*Add-in-only XML manifest[\s\S]*?-->\n/, `<!-- Generated for the "${envName}" environment by npm run manifests. Do not edit by hand. -->\n`);
}
