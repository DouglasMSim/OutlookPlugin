import { formatMegabytes, type Limits } from "./limits";
import type { IssueReason } from "./types";

/** Outlook notification messages are limited to 150 characters. */
export const NOTIFICATION_MAX = 150;

export function clampNotification(text: string): string {
  return text.length <= NOTIFICATION_MAX ? text : `${text.slice(0, NOTIFICATION_MAX - 1)}…`;
}

function plural(count: number, word: string): string {
  return `${count} ${word}${count === 1 ? "" : "s"}`;
}

export const TEXT = {
  progress: (files: number) => `Preparing ${plural(files, "attachment")} for Reply All…`,
  unsupported:
    "This Outlook version can't add files to a reply from an add-in. Update Outlook (classic Outlook needs Version 2412 or later).",
  selectionChanged:
    "The selected message changed while files were being prepared. Nothing was opened. Select the message and try again.",
  openFailed: (code: string) =>
    `Outlook reported an error opening the reply (${code}). Check for an open draft before trying again.`,
  confirmUnavailable: (issues: number) =>
    `${plural(issues, "attachment")} can't be included and the confirmation window couldn't open. No reply was opened.`,
  unexpected: "Reply All with Attachments hit an unexpected error. Check for an open draft before trying again.",
} as const;

export function describeIssue(reason: IssueReason, limits: Limits): string {
  switch (reason) {
    case "attached-item":
      return "Attached email or calendar item (not supported yet)";
    case "cloud":
      return "Cloud file link (Outlook replies can't include these)";
    case "protected":
      return "Protected or encrypted content";
    case "unsupported-type":
      return "An attachment type Outlook can't add to a reply";
    case "file-too-large":
      return `Larger than the ${formatMegabytes(limits.maxFileBytes)} per-file limit`;
    case "total-too-large":
      return `Would take the reply over the ${formatMegabytes(limits.maxTotalBytes)} total limit`;
    case "too-many-files":
      return `More than ${limits.maxFiles} files`;
    case "name-too-long":
      return `File name is longer than ${limits.maxNameLength} characters`;
    case "read-failed":
      return "Outlook couldn't read this file";
    case "unexpected-format":
      return "Outlook returned this file in an unexpected format";
  }
}
