import { describe, expect, it } from "vitest";
import { classifyAttachment, planAttachments } from "../../src/core/classify";
import { attachment, testLimits } from "./helpers";

const MB = 1024 * 1024;

describe("classifyAttachment", () => {
  it("treats ordinary files as files", () => {
    expect(classifyAttachment(attachment("Draft Agreement.docx"))).toBe("file");
    expect(classifyAttachment(attachment("Exhibit A.pdf", { contentType: "application/pdf" }))).toBe("file");
  });

  it("treats an attached (non-inline) image as an ordinary file", () => {
    expect(classifyAttachment(attachment("photo.png", { contentType: "image/png" }))).toBe("file");
  });

  it("leaves inline images (signature logos) to Outlook's quoted body", () => {
    expect(classifyAttachment(attachment("image001.png", { contentType: "image/png", isInline: true }))).toBe(
      "inline-graphic",
    );
    // Recognised by extension when the content type is missing.
    expect(classifyAttachment(attachment("logo.JPG", { contentType: "", isInline: true }))).toBe("inline-graphic");
  });

  it("keeps inline non-image files, which some mail apps mark inline", () => {
    expect(classifyAttachment(attachment("contract.pdf", { contentType: "application/pdf", isInline: true }))).toBe(
      "file",
    );
  });

  it("flags attached emails, calendar items, cloud links and unknown types", () => {
    expect(classifyAttachment(attachment("Fwd: hello", { attachmentType: "item" }))).toBe("attached-item");
    expect(classifyAttachment(attachment("Big deck.pptx", { attachmentType: "cloud" }))).toBe("cloud");
    expect(classifyAttachment(attachment("thing", { attachmentType: "somethingNew" }))).toBe("unsupported-type");
  });

  it("flags protected wrappers", () => {
    expect(classifyAttachment(attachment("message.rpmsg"))).toBe("protected");
    expect(classifyAttachment(attachment("smime.p7m"))).toBe("protected");
    expect(classifyAttachment(attachment("x.bin", { contentType: "application/pkcs7-mime" }))).toBe("protected");
  });
});

describe("planAttachments", () => {
  it("keeps original order and skips inline graphics", () => {
    const list = [
      attachment("b.docx"),
      attachment("logo.png", { contentType: "image/png", isInline: true }),
      attachment("a.pdf"),
    ];
    const plan = planAttachments(list, testLimits);
    expect(plan.toRead.map((f) => f.info.name)).toEqual(["b.docx", "a.pdf"]);
    expect(plan.toRead.map((f) => f.index)).toEqual([0, 2]);
    expect(plan.inlineGraphics).toBe(1);
    expect(plan.issues).toEqual([]);
  });

  it("keeps distinct files that share a name", () => {
    const plan = planAttachments([attachment("scan.pdf"), attachment("scan.pdf")], testLimits);
    expect(plan.toRead).toHaveLength(2);
    expect(plan.toRead[0]?.info.id).not.toBe(plan.toRead[1]?.info.id);
  });

  it("reports unsupported attachments with their position", () => {
    const plan = planAttachments([attachment("a.pdf"), attachment("Fwd", { attachmentType: "item" })], testLimits);
    expect(plan.issues).toEqual([{ index: 1, name: "Fwd", reason: "attached-item" }]);
    expect(plan.toRead).toHaveLength(1);
  });

  it("does not read a file over the per-file limit", () => {
    const plan = planAttachments(
      [attachment("huge.zip", { size: testLimits.maxFileBytes + 1 }), attachment("ok.pdf")],
      testLimits,
    );
    expect(plan.issues).toEqual([{ index: 0, name: "huge.zip", reason: "file-too-large" }]);
    expect(plan.toRead.map((f) => f.info.name)).toEqual(["ok.pdf"]);
  });

  it("admits files in order until the total limit, then reports the rest", () => {
    const limits = { ...testLimits, maxFileBytes: 10 * MB, maxTotalBytes: 15 * MB };
    const plan = planAttachments(
      [attachment("one", { size: 8 * MB }), attachment("two", { size: 8 * MB }), attachment("three", { size: 5 * MB })],
      limits,
    );
    expect(plan.toRead.map((f) => f.info.name)).toEqual(["one", "three"]);
    expect(plan.issues).toEqual([{ index: 1, name: "two", reason: "total-too-large" }]);
  });

  it("enforces the file count and name length limits", () => {
    const limits = { ...testLimits, maxFiles: 2 };
    const longName = `${"n".repeat(256)}.pdf`;
    const plan = planAttachments([attachment("1"), attachment(longName), attachment("2"), attachment("3")], limits);
    expect(plan.toRead.map((f) => f.info.name)).toEqual(["1", "2"]);
    expect(plan.issues.map((i) => i.reason)).toEqual(["name-too-long", "too-many-files"]);
  });

  it("handles a message with no attachments", () => {
    const plan = planAttachments([], testLimits);
    expect(plan).toEqual({ toRead: [], issues: [], inlineGraphics: 0, totalAttachments: 0 });
  });
});
