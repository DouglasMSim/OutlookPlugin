import { describe, expect, it } from "vitest";
import { planAttachments } from "../../src/core/classify";
import { decide, resolveConfirmation, toReplyAttachments } from "../../src/core/decide";
import { attachment, b64, testLimits } from "./helpers";

describe("decide (all-or-ask)", () => {
  it("opens a plain Reply All when there is nothing to carry over", () => {
    const plan = planAttachments([attachment("logo.png", { contentType: "image/png", isInline: true })], testLimits);
    expect(decide(plan, [], testLimits)).toEqual({ kind: "plain-reply" });
  });

  it("opens one reply with every file when all reads succeed", () => {
    const plan = planAttachments([attachment("a.docx"), attachment("b.pdf")], testLimits);
    const decision = decide(plan, [{ ok: true, base64: b64("A") }, { ok: true, base64: b64("BB") }], testLimits);
    expect(decision).toEqual({
      kind: "reply-with-files",
      files: [
        { index: 0, name: "a.docx", base64: b64("A"), bytes: 1 },
        { index: 1, name: "b.pdf", base64: b64("BB"), bytes: 2 },
      ],
    });
  });

  it("stops for confirmation when any read fails, keeping the available files", () => {
    const plan = planAttachments([attachment("a.docx"), attachment("b.pdf")], testLimits);
    const decision = decide(plan, [{ ok: true, base64: b64("A") }, { ok: false, reason: "read-failed" }], testLimits);
    expect(decision.kind).toBe("needs-confirmation");
    if (decision.kind !== "needs-confirmation") return;
    expect(decision.available.map((f) => f.name)).toEqual(["a.docx"]);
    expect(decision.issues).toEqual([{ index: 1, name: "b.pdf", reason: "read-failed" }]);
  });

  it("stops for confirmation when an unsupported attachment is present, even if every file read", () => {
    const plan = planAttachments([attachment("Fwd", { attachmentType: "item" }), attachment("a.pdf")], testLimits);
    const decision = decide(plan, [{ ok: true, base64: b64("A") }], testLimits);
    expect(decision.kind).toBe("needs-confirmation");
  });

  it("re-checks real decoded sizes after reading", () => {
    const limits = { ...testLimits, maxFileBytes: 3, maxTotalBytes: 4 };
    // Reported sizes of 0 let both through planning; real sizes are 3 and 3 bytes.
    const plan = planAttachments([attachment("a", { size: 0 }), attachment("b", { size: 0 })], limits);
    const decision = decide(plan, [{ ok: true, base64: b64("aaa") }, { ok: true, base64: b64("bbb") }], limits);
    expect(decision.kind).toBe("needs-confirmation");
    if (decision.kind !== "needs-confirmation") return;
    expect(decision.issues).toEqual([{ index: 1, name: "b", reason: "total-too-large" }]);
  });

  it("orders issues by original attachment position", () => {
    const plan = planAttachments(
      [attachment("a.pdf"), attachment("Fwd", { attachmentType: "item" }), attachment("c.pdf")],
      testLimits,
    );
    const decision = decide(plan, [{ ok: false, reason: "read-failed" }, { ok: true, base64: b64("C") }], testLimits);
    if (decision.kind !== "needs-confirmation") throw new Error("expected confirmation");
    expect(decision.issues.map((i) => i.index)).toEqual([0, 1]);
  });

  it("requires one outcome per planned file", () => {
    const plan = planAttachments([attachment("a.pdf")], testLimits);
    expect(() => decide(plan, [], testLimits)).toThrow();
  });
});

describe("resolveConfirmation", () => {
  const plan = planAttachments([attachment("a.pdf"), attachment("Fwd", { attachmentType: "item" })], testLimits);
  const decision = decide(plan, [{ ok: true, base64: b64("A") }], testLimits);
  if (decision.kind !== "needs-confirmation") throw new Error("expected confirmation");

  it("opens nothing on cancel", () => {
    expect(resolveConfirmation(decision, "cancel")).toBeUndefined();
  });

  it("continues with the available files", () => {
    expect(resolveConfirmation(decision, "continue")).toEqual({ kind: "reply-with-files", files: decision.available });
  });

  it("continues with a plain reply when no files are available", () => {
    expect(resolveConfirmation({ ...decision, available: [] }, "continue")).toEqual({ kind: "plain-reply" });
  });
});

describe("toReplyAttachments", () => {
  it("maps to Base64 ReplyFormAttachment objects, preserving order, names and bytes", () => {
    const files = [
      { index: 0, name: "Résumé – 2025.docx", base64: b64("one"), bytes: 3 },
      { index: 3, name: "契約書.pdf", base64: b64("two"), bytes: 3 },
      { index: 4, name: "契約書.pdf", base64: b64("three"), bytes: 5 },
    ];
    expect(toReplyAttachments(files)).toEqual([
      { type: "base64", name: "Résumé – 2025.docx", base64file: b64("one"), inLine: false },
      { type: "base64", name: "契約書.pdf", base64file: b64("two"), inLine: false },
      { type: "base64", name: "契約書.pdf", base64file: b64("three"), inLine: false },
    ]);
  });
});
