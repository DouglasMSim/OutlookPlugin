import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { decodeCounts, encodeCounts, parseDialogDetails, parseDialogToParent } from "../../src/core/dialog-protocol";
import { renderManifest } from "../../scripts/lib/manifest.mjs";

const root = fileURLToPath(new URL("../..", import.meta.url));

function sourceFiles(dir: string): string[] {
  return readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    return statSync(path).isDirectory() ? sourceFiles(path) : path.endsWith(".ts") ? [path] : [];
  });
}

describe("privacy", () => {
  it("only src/core/log.ts writes to the console", () => {
    const offenders = sourceFiles(join(root, "src")).filter(
      (file) => !file.endsWith(join("core", "log.ts")) && /\bconsole\./.test(readFileSync(file, "utf8")),
    );
    expect(offenders).toEqual([]);
  });

  it("never puts file names in the dialog URL", () => {
    expect(encodeCounts(2, 3)).toBe("a=2&i=3");
    expect(decodeCounts("#a=2&i=3")).toEqual({ availableCount: 2, issueCount: 3 });
    // Office may append its own query string after the fragment.
    expect(decodeCounts("#a=1&i=4?_host_Info=Outlook")).toEqual({ availableCount: 1, issueCount: 4 });
    expect(decodeCounts("")).toBeUndefined();
  });
});

describe("dialog protocol", () => {
  it("parses only known messages", () => {
    expect(parseDialogToParent('{"type":"continue"}')).toEqual({ type: "continue" });
    expect(parseDialogToParent('{"type":"delete-everything"}')).toBeUndefined();
    expect(parseDialogToParent("not json")).toBeUndefined();
    expect(parseDialogDetails('{"type":"details","availableCount":1,"issues":[{"name":"a","text":"b"}]}')).toEqual({
      type: "details",
      availableCount: 1,
      issues: [{ name: "a", text: "b" }],
    });
    expect(parseDialogDetails('{"type":"details","availableCount":1,"issues":[{"name":3}]}')).toBeUndefined();
  });
});

describe("manifests", () => {
  for (const env of ["dev", "pilot"] as const) {
    it(`${env} manifest declares the read-surface ExecuteFunction command`, async () => {
      const xml = await renderManifest(env);
      expect(xml).toContain("<Permissions>ReadItem</Permissions>");
      expect(xml).toContain('<bt:Sets DefaultMinVersion="1.15">');
      expect(xml).toContain('xsi:type="MessageReadCommandSurface"');
      expect(xml).toContain('<Action xsi:type="ExecuteFunction">');
      expect(xml).toContain("<FunctionName>replyAllWithAttachments</FunctionName>");
      expect(xml).toContain("<SupportsSharedFolders>true</SupportsSharedFolders>");
      expect(xml).not.toMatch(/\{\{[A-Z_]+\}\}/);
      for (const url of xml.match(/https?:\/\/[^"<\s]+/g) ?? []) {
        // XML namespace identifiers are not fetched.
        if (url.startsWith("http://schemas.microsoft.com/") || url.startsWith("http://www.w3.org/")) continue;
        expect(url.startsWith("https://")).toBe(true);
      }
    });
  }

  it("uses different add-in IDs for dev and pilot", async () => {
    const id = (xml: string) => /<Id>([^<]+)<\/Id>/.exec(xml)?.[1];
    expect(id(await renderManifest("dev"))).not.toBe(id(await renderManifest("pilot")));
  });

  it("references only resource IDs that are defined, and icon files that exist", async () => {
    const xml = await renderManifest("pilot");
    const defined = new Set([...xml.matchAll(/<bt:(?:Image|Url|String) id="([^"]+)"/g)].map((m) => m[1]));
    for (const [, ref] of xml.matchAll(/resid="([^"]+)"/g)) expect(defined).toContain(ref);
    for (const [, file] of xml.matchAll(/\/assets\/(icon-\d+\.png)/g)) {
      expect(statSync(join(root, "assets", file as string)).isFile()).toBe(true);
    }
  });

  it("registers the same function name the manifest calls", () => {
    const source = readFileSync(join(root, "src", "commands", "commands.ts"), "utf8");
    expect(source).toContain('associate("replyAllWithAttachments"');
  });
});
