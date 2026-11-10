// Serve dist/ over HTTPS on https://localhost:3000 for sideloading manifest.dev.xml.
// First run `npm run dev-certs` once to install a trusted localhost certificate.
import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import { createServer } from "node:https";
import { extname, join, normalize } from "node:path";
import { getHttpsServerOptions } from "office-addin-dev-certs";
import { root } from "./lib/manifest.mjs";

const PORT = Number(process.env.PORT || 3000);
const dist = join(root, "dist");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml; charset=utf-8",
  ".map": "application/json; charset=utf-8",
};

const options = await getHttpsServerOptions();

createServer(options, async (req, res) => {
  const pathname = decodeURIComponent(new URL(req.url ?? "/", "https://localhost").pathname);
  const file = normalize(join(dist, pathname === "/" ? "index.html" : pathname));
  if (!file.startsWith(dist)) {
    res.writeHead(403).end();
    return;
  }
  try {
    if (!(await stat(file)).isFile()) throw new Error("not a file");
    res.writeHead(200, { "Content-Type": TYPES[extname(file)] ?? "application/octet-stream", "Cache-Control": "no-store" });
    createReadStream(file).pipe(res);
  } catch {
    res.writeHead(404).end("Not found");
  }
}).listen(PORT, () => {
  console.log(`Serving dist/ at https://localhost:${PORT}/`);
  console.log("Sideload manifests/manifest.dev.xml (see docs/sideload-guide.md).");
});
