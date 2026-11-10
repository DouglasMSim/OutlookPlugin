import { describe, expect, it } from "vitest";
import { decodedByteLength, normalizeBase64 } from "../../src/core/base64";
import { mapWithConcurrency } from "../../src/core/concurrency";
import { b64, deferred } from "./helpers";

describe("base64 helpers", () => {
  it("computes decoded length including padding", () => {
    for (const text of ["", "a", "ab", "abc", "abcd", "hello world", "契約書"]) {
      expect(decodedByteLength(b64(text))).toBe(new TextEncoder().encode(text).length);
    }
  });

  it("accepts valid Base64 and strips line breaks", () => {
    expect(normalizeBase64("aGVs\r\nbG8=")).toBe("aGVsbG8=");
    expect(normalizeBase64("")).toBe("");
  });

  it("rejects malformed Base64", () => {
    expect(normalizeBase64("abc")).toBeUndefined();
    expect(normalizeBase64("ab$=")).toBeUndefined();
    expect(normalizeBase64("a===")).toBeUndefined();
  });
});

describe("mapWithConcurrency", () => {
  it("keeps input order when tasks finish out of order", async () => {
    const gates = [deferred<void>(), deferred<void>(), deferred<void>()];
    const run = mapWithConcurrency([0, 1, 2], 3, async (n) => {
      await gates[n]?.promise;
      return n * 10;
    });
    gates[2]?.resolve();
    gates[0]?.resolve();
    gates[1]?.resolve();
    expect(await run).toEqual([0, 10, 20]);
  });

  it("never exceeds the limit", async () => {
    let active = 0;
    let peak = 0;
    await mapWithConcurrency(Array.from({ length: 10 }, (_, i) => i), 3, async () => {
      active += 1;
      peak = Math.max(peak, active);
      await new Promise((r) => setTimeout(r, 2));
      active -= 1;
    });
    expect(peak).toBe(3);
  });

  it("handles an empty list", async () => {
    expect(await mapWithConcurrency([], 3, async () => 1)).toEqual([]);
  });
});
