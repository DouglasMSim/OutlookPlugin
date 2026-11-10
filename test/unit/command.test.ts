import { describe, expect, it } from "vitest";
import { createCommandHandler, runCommand, type HostPort } from "../../src/commands/command";
import { attachment, b64, deferred, fakeHost, fakeItem, testLimits } from "./helpers";

describe("runCommand", () => {
  it("explains and stops when Outlook lacks Mailbox 1.15", async () => {
    const item = fakeItem({ attachments: [attachment("a.pdf")] });
    const outcome = await runCommand(fakeHost({ item, supported: false }), testLimits);
    expect(outcome).toBe("unsupported");
    expect(item.opened).toEqual([]);
    expect(item.reads).toEqual([]);
    expect(item.statuses.at(-1)?.kind).toBe("error");
  });

  it("does nothing when no message is available", async () => {
    expect(await runCommand(fakeHost({}), testLimits)).toBe("no-item");
  });

  it("opens a normal Reply All when there are no attachments", async () => {
    const item = fakeItem({ attachments: [] });
    expect(await runCommand(fakeHost({ item }), testLimits)).toBe("opened-plain");
    expect(item.opened).toEqual([[]]);
  });

  it("opens a normal Reply All when the only attachments are inline graphics", async () => {
    const logo = attachment("image001.png", { contentType: "image/png", isInline: true });
    const item = fakeItem({ attachments: [logo] });
    expect(await runCommand(fakeHost({ item }), testLimits)).toBe("opened-plain");
    expect(item.reads).toEqual([]);
    expect(item.opened).toEqual([[]]);
  });

  it("DOCX + PDF + inline logo: one reply with both files, in order, bytes intact", async () => {
    const docx = attachment("Draft Agreement.docx");
    const logo = attachment("image001.png", { contentType: "image/png", isInline: true });
    const pdf = attachment("Exhibit A.pdf", { contentType: "application/pdf" });
    const item = fakeItem({
      attachments: [docx, logo, pdf],
      contents: { [docx.id]: { ok: true, base64: b64("DOCX") }, [pdf.id]: { ok: true, base64: b64("PDF") } },
    });
    const host = fakeHost({ item });

    expect(await runCommand(host, testLimits)).toBe("opened-with-files");
    expect(item.reads).toEqual([docx.id, pdf.id]);
    expect(item.opened).toEqual([
      [
        { type: "base64", name: "Draft Agreement.docx", base64file: b64("DOCX"), inLine: false },
        { type: "base64", name: "Exhibit A.pdf", base64file: b64("PDF"), inLine: false },
      ],
    ]);
    expect(host.confirmations).toEqual([]);
    expect(item.statuses[0]).toEqual({ kind: "progress", message: "Preparing 2 attachments for Reply All…" });
    expect(item.statuses.at(-1)).toEqual({ kind: "clear" });
  });

  it("keeps duplicate and Unicode file names", async () => {
    const a = attachment("契約書.pdf");
    const b = attachment("契約書.pdf");
    const item = fakeItem({
      attachments: [a, b],
      contents: { [a.id]: { ok: true, base64: b64("first") }, [b.id]: { ok: true, base64: b64("second") } },
    });
    await runCommand(fakeHost({ item }), testLimits);
    expect(item.opened[0]?.map((x) => [x.name, x.base64file])).toEqual([
      ["契約書.pdf", b64("first")],
      ["契約書.pdf", b64("second")],
    ]);
  });

  it("asks before opening when an attachment is unsupported, and opens nothing on cancel", async () => {
    const item = fakeItem({ attachments: [attachment("a.pdf"), attachment("Fwd: note", { attachmentType: "item" })] });
    const host = fakeHost({ item, confirm: "cancel" });
    expect(await runCommand(host, testLimits)).toBe("cancelled");
    expect(item.opened).toEqual([]);
    expect(host.confirmations).toEqual([
      { availableCount: 1, issues: [{ name: "Fwd: note", text: "Attached email or calendar item (not supported yet)" }] },
    ]);
  });

  it("continues with the available files only after the sender agrees", async () => {
    const pdf = attachment("a.pdf");
    const item = fakeItem({
      attachments: [pdf, attachment("deck.pptx", { attachmentType: "cloud" })],
      contents: { [pdf.id]: { ok: true, base64: b64("A") } },
    });
    expect(await runCommand(fakeHost({ item, confirm: "continue" }), testLimits)).toBe("opened-with-files");
    expect(item.events).toEqual([`read:${pdf.id}`, "confirm", "open"]);
    expect(item.opened).toEqual([[{ type: "base64", name: "a.pdf", base64file: b64("A"), inLine: false }]]);
  });

  it("opens a plain reply if the sender continues and no files are available", async () => {
    const item = fakeItem({ attachments: [attachment("Fwd", { attachmentType: "item" })] });
    expect(await runCommand(fakeHost({ item, confirm: "continue" }), testLimits)).toBe("opened-plain");
    expect(item.opened).toEqual([[]]);
  });

  it("treats a failed or rejected read as an issue, never a silent omission", async () => {
    const a = attachment("a.pdf");
    const b = attachment("b.pdf");
    const c = attachment("c.pdf");
    const item = fakeItem({
      attachments: [a, b, c],
      contents: {
        [a.id]: { ok: false, reason: "read-failed" },
        [b.id]: () => Promise.reject(new Error("boom")),
        [c.id]: { ok: false, reason: "unexpected-format" },
      },
    });
    const host = fakeHost({ item, confirm: "cancel" });
    await runCommand(host, testLimits);
    expect(host.confirmations[0]?.issues.map((i) => i.name)).toEqual(["a.pdf", "b.pdf", "c.pdf"]);
    expect(host.confirmations[0]?.availableCount).toBe(0);
    expect(item.opened).toEqual([]);
  });

  it("does not read a file over the size limit", async () => {
    const big = attachment("big.zip", { size: testLimits.maxFileBytes + 1 });
    const item = fakeItem({ attachments: [big] });
    const host = fakeHost({ item, confirm: "cancel" });
    await runCommand(host, testLimits);
    expect(item.reads).toEqual([]);
    expect(host.confirmations[0]?.issues[0]?.text).toMatch(/per-file limit/);
  });

  it("reports and opens nothing if the confirmation dialog cannot be shown", async () => {
    const item = fakeItem({ attachments: [attachment("Fwd", { attachmentType: "item" })] });
    expect(await runCommand(fakeHost({ item, confirm: "unavailable" }), testLimits)).toBe("confirm-unavailable");
    expect(item.opened).toEqual([]);
    expect(item.statuses.at(-1)?.kind).toBe("error");
  });

  it("opens nothing if the selected message changed during processing", async () => {
    const item = fakeItem({ attachments: [attachment("a.pdf")], itemId: "original" });
    const host = fakeHost({ item, currentItemId: () => "different" });
    expect(await runCommand(host, testLimits)).toBe("selection-changed");
    expect(item.opened).toEqual([]);
  });

  it("surfaces an Outlook error and does not retry", async () => {
    const item = fakeItem({ attachments: [attachment("a.pdf")], openResult: { ok: false, code: "9041" } });
    expect(await runCommand(fakeHost({ item }), testLimits)).toBe("open-failed");
    expect(item.opened).toHaveLength(1);
    expect(item.statuses.at(-1)?.message).toContain("9041");
  });
});

describe("createCommandHandler", () => {
  function counter() {
    const event = { calls: 0, completed() { event.calls += 1; } };
    return event;
  }

  it("completes the Office event exactly once, after the reply opens", async () => {
    const item = fakeItem({ attachments: [attachment("a.pdf")] });
    const handle = createCommandHandler(() => fakeHost({ item }), testLimits);
    const event = counter();
    const order: string[] = [];
    const originalOpen = item.openReplyAll.bind(item);
    item.openReplyAll = async (a) => {
      order.push("open");
      return originalOpen(a);
    };
    const originalCompleted = event.completed.bind(event);
    event.completed = () => {
      order.push("completed");
      originalCompleted();
    };
    await handle(event);
    expect(event.calls).toBe(1);
    expect(order).toEqual(["open", "completed"]);
  });

  it("completes the event when something throws", async () => {
    const host: HostPort = {
      ...fakeHost({}),
      getItem: () => {
        throw new Error("Outlook went away");
      },
    };
    const handle = createCommandHandler(() => host, testLimits);
    const event = counter();
    expect(await handle(event)).toBe("error");
    expect(event.calls).toBe(1);
  });

  it("ignores a second click while the first is still working, and opens only one reply", async () => {
    const gate = deferred<void>();
    const a = attachment("a.pdf");
    const item = fakeItem({
      attachments: [a],
      contents: {
        [a.id]: async () => {
          await gate.promise;
          return { ok: true, base64: b64("A") };
        },
      },
    });
    const handle = createCommandHandler(() => fakeHost({ item }), testLimits);
    const first = counter();
    const second = counter();

    const firstRun = handle(first);
    expect(await handle(second)).toBe("busy");
    expect(second.calls).toBe(1);
    expect(first.calls).toBe(0);

    gate.resolve();
    expect(await firstRun).toBe("opened-with-files");
    expect(first.calls).toBe(1);
    expect(item.opened).toHaveLength(1);
  });

  it("accepts a new click after the previous one finished", async () => {
    const item = fakeItem({ attachments: [] });
    const handle = createCommandHandler(() => fakeHost({ item }), testLimits);
    await handle(counter());
    await handle(counter());
    expect(item.opened).toHaveLength(2);
  });
});
