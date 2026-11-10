/**
 * Snapshot of the metadata Outlook reports for one attachment on the selected
 * message. Copied out of Office.AttachmentDetails at invocation time so later
 * steps never re-read the (possibly changed) live item.
 */
export interface AttachmentInfo {
  /** Outlook attachment ID. Not an Exchange item ID; only valid for this item and session. */
  id: string;
  name: string;
  /** Size in bytes as reported by Outlook. */
  size: number;
  /** Office.MailboxEnums.AttachmentType value: "file", "item", "cloud", or something newer. */
  attachmentType: string;
  isInline: boolean;
  contentType: string;
}

/** Why a specific attachment cannot be carried into the reply. */
export type IssueReason =
  | "attached-item"
  | "cloud"
  | "protected"
  | "unsupported-type"
  | "file-too-large"
  | "total-too-large"
  | "too-many-files"
  | "name-too-long"
  | "read-failed"
  | "unexpected-format";

export interface Issue {
  /** Position in the message's original attachment list. */
  index: number;
  name: string;
  reason: IssueReason;
}

/** An attachment selected for reading, with its original position. */
export interface PlannedFile {
  index: number;
  info: AttachmentInfo;
}

export type ReadOutcome =
  | { ok: true; base64: string }
  | { ok: false; reason: "read-failed" | "unexpected-format" };

/** A file whose content was read successfully. */
export interface ReadFile {
  index: number;
  name: string;
  base64: string;
  bytes: number;
}

/**
 * Shape of an Office.ReplyFormAttachment of type Base64 (Mailbox 1.15).
 * Declared locally so core logic does not depend on the Office global.
 */
export interface Base64ReplyAttachment {
  type: "base64";
  name: string;
  base64file: string;
  inLine: false;
}
