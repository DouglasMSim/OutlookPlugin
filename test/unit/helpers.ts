import type { ConfirmRequest, HostPort, MailItemPort, OpenResult } from "../../src/commands/command";
import type { ConfirmChoice } from "../../src/core/decide";
import { DEFAULT_LIMITS, type Limits } from "../../src/core/limits";
import type { AttachmentInfo, Base64ReplyAttachment, ReadOutcome } from "../../src/core/types";

let nextId = 1;

export function attachment(name: string, overrides: Partial<AttachmentInfo> = {}): AttachmentInfo {
  return {
    id: `att-${nextId++}`,
    name,
    size: 1000,
    attachmentType: "file",
    isInline: false,
    contentType: "application/octet-stream",
    ...overrides,
  };
}

/** Base64 of a UTF-8 string, standing in for real file bytes. */
export function b64(text: string): string {
  return btoa(String.fromCharCode(...new TextEncoder().encode(text)));
}

export const testLimits: Limits = { ...DEFAULT_LIMITS, readTimeoutMs: 1000 };

type ReadBehaviour = ReadOutcome | (() => Promise<ReadOutcome>);

export interface FakeItem extends MailItemPort {
  reads: string[];
  opened: Base64ReplyAttachment[][];
  statuses: { kind: "progress" | "error" | "clear"; message?: string }[];
  events: string[];
}

export function fakeItem(options: {
  attachments: AttachmentInfo[];
  contents?: Record<string, ReadBehaviour>;
  openResult?: OpenResult;
  itemId?: string;
}): FakeItem {
  const item: FakeItem = {
    itemId: options.itemId ?? "item-1",
    attachments: options.attachments,
    reads: [],
    opened: [],
    statuses: [],
    events: [],
    async readAttachment(id: string): Promise<ReadOutcome> {
      item.reads.push(id);
      item.events.push(`read:${id}`);
      const behaviour = options.contents?.[id];
      if (behaviour === undefined) return { ok: true, base64: b64(`content of ${id}`) };
      return typeof behaviour === "function" ? behaviour() : behaviour;
    },
    async openReplyAll(attachments: Base64ReplyAttachment[]): Promise<OpenResult> {
      item.opened.push(attachments);
      item.events.push("open");
      return options.openResult ?? { ok: true };
    },
    async showProgress(message: string) {
      item.statuses.push({ kind: "progress", message });
    },
    async showError(message: string) {
      item.statuses.push({ kind: "error", message });
      item.events.push("error");
    },
    async clearStatus() {
      item.statuses.push({ kind: "clear" });
    },
  };
  return item;
}

export interface FakeHost extends HostPort {
  confirmations: ConfirmRequest[];
}

export function fakeHost(options: {
  item?: MailItemPort;
  supported?: boolean;
  currentItemId?: () => string | undefined;
  confirm?: ConfirmChoice | "unavailable";
}): FakeHost {
  const host: FakeHost = {
    confirmations: [],
    supportsRequiredApi: () => options.supported ?? true,
    getItem: () => options.item,
    currentItemId: options.currentItemId ?? (() => options.item?.itemId),
    async confirm(request: ConfirmRequest) {
      host.confirmations.push(request);
      if (options.item && "events" in options.item) (options.item as FakeItem).events.push("confirm");
      return options.confirm ?? "cancel";
    },
  };
  return host;
}

/** A promise you resolve later, for controlling timing in tests. */
export function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((r) => {
    resolve = r;
  });
  return { promise, resolve };
}
