import type { AttachmentInfo, Issue, IssueReason, PlannedFile } from "./types";
import type { Limits } from "./limits";

const IMAGE_EXTENSION = /\.(png|jpe?g|gif|bmp|tiff?|webp|svg|emf|wmf|ico|heic)$/i;

/** Wrappers that mean the real content is encrypted or rights-managed. */
const PROTECTED_NAMES = /^(message\.rpmsg|smime\.p7m)$/i;
const PROTECTED_CONTENT_TYPES = new Set([
  "application/x-microsoft-rpmsg-message",
  "application/pkcs7-mime",
  "application/x-pkcs7-mime",
]);

export type AttachmentClass = "file" | "inline-graphic" | Extract<
  IssueReason,
  "attached-item" | "cloud" | "protected" | "unsupported-type"
>;

export function isImage(info: AttachmentInfo): boolean {
  return info.contentType.toLowerCase().startsWith("image/") || IMAGE_EXTENSION.test(info.name);
}

/**
 * Decide what kind of attachment this is.
 *
 * Only inline *images* are left to Outlook's quoted body. Some senders (for
 * example Apple Mail) mark ordinary PDFs or documents as inline. Those cannot
 * render in the quoted HTML, so skipping them would silently drop a real file.
 */
export function classifyAttachment(info: AttachmentInfo): AttachmentClass {
  const type = info.attachmentType.toLowerCase();
  if (type === "item") return "attached-item";
  if (type === "cloud") return "cloud";
  if (type !== "file") return "unsupported-type";
  if (PROTECTED_NAMES.test(info.name) || PROTECTED_CONTENT_TYPES.has(info.contentType.toLowerCase())) {
    return "protected";
  }
  if (info.isInline && isImage(info)) return "inline-graphic";
  return "file";
}

export interface AttachmentPlan {
  /** Files to read, in original order. */
  toRead: PlannedFile[];
  /** Attachments that cannot be included, found before reading anything. */
  issues: Issue[];
  /** Inline graphics left to Outlook's quoted body. */
  inlineGraphics: number;
  /** Total attachments on the message, including inline graphics. */
  totalAttachments: number;
}

/**
 * Classify every attachment and apply the known limits using the sizes Outlook
 * reports, so nothing over a limit is read into memory. Files are admitted in
 * their original order; once the total cap is reached, later files are
 * reported rather than silently dropped.
 */
export function planAttachments(attachments: readonly AttachmentInfo[], limits: Limits): AttachmentPlan {
  const toRead: PlannedFile[] = [];
  const issues: Issue[] = [];
  let inlineGraphics = 0;
  let plannedBytes = 0;

  attachments.forEach((info, index) => {
    const kind = classifyAttachment(info);
    if (kind === "inline-graphic") {
      inlineGraphics += 1;
      return;
    }
    if (kind !== "file") {
      issues.push({ index, name: info.name, reason: kind });
      return;
    }

    const size = Number.isFinite(info.size) && info.size > 0 ? info.size : 0;
    let reason: IssueReason | undefined;
    if (info.name.length > limits.maxNameLength) reason = "name-too-long";
    else if (size > limits.maxFileBytes) reason = "file-too-large";
    else if (toRead.length >= limits.maxFiles) reason = "too-many-files";
    else if (plannedBytes + size > limits.maxTotalBytes) reason = "total-too-large";

    if (reason) {
      issues.push({ index, name: info.name, reason });
      return;
    }
    plannedBytes += size;
    toRead.push({ index, info });
  });

  return { toRead, issues, inlineGraphics, totalAttachments: attachments.length };
}
