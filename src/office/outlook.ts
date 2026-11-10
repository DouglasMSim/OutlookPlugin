import type { ConfirmRequest, HostPort, MailItemPort, OpenResult } from "../commands/command";
import { normalizeBase64 } from "../core/base64";
import { clampNotification } from "../core/messages";
import type { AttachmentInfo, Base64ReplyAttachment, ReadOutcome } from "../core/types";
import { confirmWithDialog } from "./dialog";

const STATUS_KEY = "replyAllWithAttachmentsStatus";

function errorCode(error: Office.Error | undefined): string {
  if (!error) return "unknown";
  return error.code !== undefined ? String(error.code) : error.name || "unknown";
}

function snapshotAttachments(item: Office.MessageRead): AttachmentInfo[] {
  return (item.attachments ?? []).map((a) => ({
    id: a.id,
    name: a.name ?? "",
    size: typeof a.size === "number" ? a.size : 0,
    attachmentType: String(a.attachmentType ?? ""),
    isInline: Boolean(a.isInline),
    contentType: a.contentType ?? "",
  }));
}

function setStatus(item: Office.MessageRead, details: Office.NotificationMessageDetails): Promise<void> {
  return new Promise((resolve) => {
    try {
      item.notificationMessages.replaceAsync(STATUS_KEY, details, () => resolve());
    } catch {
      resolve();
    }
  });
}

export function createMailItemPort(item: Office.MessageRead): MailItemPort {
  // Snapshot everything up front. Later calls go to this exact item object.
  const itemId = typeof item.itemId === "string" && item.itemId ? item.itemId : undefined;
  const attachments = snapshotAttachments(item);

  return {
    itemId,
    attachments,

    readAttachment(id: string, timeoutMs: number): Promise<ReadOutcome> {
      return new Promise((resolve) => {
        let settled = false;
        const settle = (outcome: ReadOutcome) => {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          resolve(outcome);
        };
        const timer = setTimeout(() => settle({ ok: false, reason: "read-failed" }), timeoutMs);
        try {
          item.getAttachmentContentAsync(id, (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
              settle({ ok: false, reason: "read-failed" });
              return;
            }
            // Only a Base64 file result is acceptable. Eml/iCalendar/Url results mean
            // this is not an ordinary file, whatever the metadata said.
            if (String(result.value.format) !== "base64" || typeof result.value.content !== "string") {
              settle({ ok: false, reason: "unexpected-format" });
              return;
            }
            const base64 = normalizeBase64(result.value.content);
            settle(base64 === undefined ? { ok: false, reason: "unexpected-format" } : { ok: true, base64 });
          });
        } catch {
          settle({ ok: false, reason: "read-failed" });
        }
      });
    },

    openReplyAll(replyAttachments: Base64ReplyAttachment[]): Promise<OpenResult> {
      return new Promise((resolve) => {
        // Empty htmlBody: Outlook supplies recipients, subject, quoted history and thread.
        const formData: Office.ReplyFormData = { htmlBody: "" };
        if (replyAttachments.length > 0) {
          formData.attachments = replyAttachments as unknown as Office.ReplyFormAttachment[];
        }
        try {
          item.displayReplyAllFormAsync(formData, (result) => {
            resolve(
              result.status === Office.AsyncResultStatus.Succeeded
                ? { ok: true }
                : { ok: false, code: errorCode(result.error) },
            );
          });
        } catch {
          // The API throws synchronously when a parameter exceeds its documented limit.
          resolve({ ok: false, code: "rejected" });
        }
      });
    },

    showProgress(message: string): Promise<void> {
      return setStatus(item, {
        type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
        message: clampNotification(message),
      });
    },

    showError(message: string): Promise<void> {
      return setStatus(item, {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: clampNotification(message),
      });
    },

    clearStatus(): Promise<void> {
      return new Promise((resolve) => {
        try {
          item.notificationMessages.removeAsync(STATUS_KEY, () => resolve());
        } catch {
          resolve();
        }
      });
    },
  };
}

export function createOfficeHost(): HostPort {
  return {
    supportsRequiredApi(): boolean {
      try {
        return Office.context.requirements.isSetSupported("Mailbox", "1.15");
      } catch {
        return false;
      }
    },

    getItem(): MailItemPort | undefined {
      const item = Office.context?.mailbox?.item as Office.MessageRead | undefined;
      if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) return undefined;
      return createMailItemPort(item);
    },

    currentItemId(): string | undefined {
      const id = Office.context?.mailbox?.item?.itemId;
      return typeof id === "string" && id ? id : undefined;
    },

    confirm(request: ConfirmRequest) {
      return confirmWithDialog(request);
    },
  };
}
