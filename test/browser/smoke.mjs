// Load the BUILT add-in (dist/) in Chromium with a simulated Office.js and drive
// the real button handler and the real confirmation dialog page.
// Run after `npm run build`. Usage: node test/browser/smoke.mjs [--screenshot <dir>]
import assert from "node:assert/strict";
import { createReadStream } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import { createServer } from "node:http";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";

const root = fileURLToPath(new URL("../..", import.meta.url));
const dist = join(root, "dist");
const mock = await readFile(join(root, "test/browser/office-mock.js"), "utf8");
const shotIndex = process.argv.indexOf("--screenshot");
const shotDir = shotIndex === -1 ? undefined : process.argv[shotIndex + 1];

const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png" };
const server = createServer(async (req, res) => {
  const file = normalize(join(dist, new URL(req.url, "http://x").pathname));
  try {
    if (!file.startsWith(dist) || !(await stat(file)).isFile()) throw new Error();
    res.writeHead(200, { "Content-Type": TYPES[extname(file)] ?? "application/octet-stream" });
    createReadStream(file).pipe(res);
  } catch {
    res.writeHead(404).end();
  }
});
await new Promise((r) => server.listen(0, "127.0.0.1", r));
const base = `http://127.0.0.1:${server.address().port}`;

const b64 = (text) => Buffer.from(text, "utf8").toString("base64");
const file = (id, name, extra = {}) => ({
  id,
  name,
  size: 10,
  attachmentType: "file",
  isInline: false,
  contentType: "application/octet-stream",
  ...extra,
});

const browser = await chromium.launch();
const results = [];

async function run(name, scenario, check) {
  const context = await browser.newContext({ deviceScaleFactor: shotDir ? 2 : 1, viewport: { width: 520, height: 520 } });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.route("**/*", (route) => {
    const url = route.request().url();
    if (url.startsWith("https://appsforoffice.microsoft.com/")) {
      return route.fulfill({ contentType: "text/javascript", body: mock });
    }
    if (url.startsWith(base)) return route.continue();
    return route.abort();
  });
  await page.addInitScript((s) => {
    window.__scenario = s;
  }, scenario);
  await page.goto(`${base}/commands.html`);
  await page.waitForFunction(() => window.__officeMock && window.__officeMock.associated.replyAllWithAttachments);
  try {
    await check(page);
    assert.deepEqual(errors, [], "page errors");
    results.push(`PASS  ${name}`);
  } catch (error) {
    results.push(`FAIL  ${name}\n      ${error.message.split("\n").join("\n      ")}`);
  } finally {
    await context.close();
  }
}

const invoke = (page) => page.evaluate(() => window.__officeMock.invoke("replyAllWithAttachments"));
const state = (page) => page.evaluate(() => JSON.parse(JSON.stringify(window.__officeMock)));

await run(
  "DOCX + PDF + inline logo opens one Reply All with both files",
  {
    attachments: [
      file("a1", "Draft Agreement.docx"),
      file("a2", "image001.png", { contentType: "image/png", isInline: true }),
      file("a3", "Exhibit A.pdf", { contentType: "application/pdf" }),
    ],
    contents: { a1: { content: b64("DOCX bytes") }, a2: { content: b64("logo") }, a3: { content: b64("PDF bytes") } },
  },
  async (page) => {
    await invoke(page);
    const s = await state(page);
    assert.equal(s.completed, 1);
    assert.deepEqual(s.reads, ["a1", "a3"]);
    assert.deepEqual(s.displayCalls, [
      {
        htmlBody: "",
        attachments: [
          { type: "base64", name: "Draft Agreement.docx", base64file: b64("DOCX bytes"), inLine: false },
          { type: "base64", name: "Exhibit A.pdf", base64file: b64("PDF bytes"), inLine: false },
        ],
      },
    ]);
    assert.equal(s.notifications[0].type, "progressIndicator");
    assert.equal(s.notifications.at(-1).op, "remove");
  },
);

await run("No attachments opens a normal Reply All", { attachments: [], contents: {} }, async (page) => {
  await invoke(page);
  const s = await state(page);
  assert.equal(s.completed, 1);
  assert.deepEqual(s.displayCalls, [{ htmlBody: "" }]);
});

await run(
  "Old Outlook (no Mailbox 1.15) shows an error and opens nothing",
  { supported: false, attachments: [file("a1", "a.pdf")], contents: { a1: { content: b64("x") } } },
  async (page) => {
    await invoke(page);
    const s = await state(page);
    assert.equal(s.completed, 1);
    assert.deepEqual(s.displayCalls, []);
    assert.equal(s.notifications.at(-1).type, "errorMessage");
  },
);

const mixed = {
  attachments: [
    file("a1", "Exhibit A.pdf"),
    file("a2", "RE: Closing checklist", { attachmentType: "item" }),
    file("a3", "Board deck.pptx", { attachmentType: "cloud" }),
    file("a4", "Scan 2.pdf"),
  ],
  contents: { a1: { content: b64("PDF") }, a4: { fail: true } },
};

await run("Unsupported + failed read: dialog lists them; Continue opens with the available file", mixed, async (page) => {
  const done = invoke(page);
  const dialog = page.frameLocator("#dialog");
  await dialog.locator("#issues li").first().waitFor();
  assert.deepEqual(await dialog.locator("#issues .name").allTextContents(), [
    "RE: Closing checklist",
    "Board deck.pptx",
    "Scan 2.pdf",
  ]);
  assert.equal(await dialog.locator("#continue").textContent(), "Continue with 1 file");
  let s = await state(page);
  assert.deepEqual(s.displayCalls, [], "nothing opens before the sender chooses");
  assert.equal(s.completed, 0, "event stays open while waiting");
  assert.ok(!s.dialogUrls[0].includes("Scan"), "file names never go in the dialog URL");
  if (shotDir) await page.screenshot({ path: join(shotDir, "confirm-dialog.png") });
  await dialog.locator("#continue").click();
  await done;
  s = await state(page);
  assert.equal(s.completed, 1);
  assert.equal(s.dialogClosed, true);
  assert.deepEqual(s.displayCalls, [
    { htmlBody: "", attachments: [{ type: "base64", name: "Exhibit A.pdf", base64file: b64("PDF"), inLine: false }] },
  ]);
});

await run("Unsupported attachment: Cancel opens nothing", mixed, async (page) => {
  const done = invoke(page);
  const dialog = page.frameLocator("#dialog");
  await dialog.locator("#issues li").first().waitFor();
  await dialog.locator("#cancel").click();
  await done;
  const s = await state(page);
  assert.equal(s.completed, 1);
  assert.deepEqual(s.displayCalls, []);
});

await run(
  "Rapid double click opens only one reply",
  { attachments: [file("a1", "a.pdf")], contents: { a1: { content: b64("A") } }, readDelayMs: 200 },
  async (page) => {
    await page.evaluate(() =>
      Promise.all([
        window.__officeMock.invoke("replyAllWithAttachments"),
        window.__officeMock.invoke("replyAllWithAttachments"),
      ]),
    );
    const s = await state(page);
    assert.equal(s.completed, 2);
    assert.equal(s.displayCalls.length, 1);
  },
);

await browser.close();
server.close();

console.log(results.join("\n"));
const failed = results.filter((r) => r.startsWith("FAIL")).length;
console.log(`\n${results.length - failed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
