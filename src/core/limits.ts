export interface Limits {
  /** Largest single file carried into a reply, in bytes. */
  maxFileBytes: number;
  /** Largest total of all files carried into one reply, in bytes. */
  maxTotalBytes: number;
  /** Most files carried into one reply. */
  maxFiles: number;
  /** Documented maximum length of Office.ReplyFormAttachment.name. */
  maxNameLength: number;
  /** How many attachments are read from Outlook at the same time. */
  readConcurrency: number;
  /** Give up on a single attachment read after this long. */
  readTimeoutMs: number;
}

const MB = 1024 * 1024;

/**
 * PROVISIONAL product limits for the pilot. See docs/pilot-notes.md.
 *
 * Exchange Online's default maximum send size is 35 MB, and that is measured
 * after Base64/MIME encoding, which adds roughly a third. About 25 MB of raw
 * file data is therefore the most a default mailbox can send in one message,
 * so a larger cap would only produce drafts that fail on send. These numbers
 * must be confirmed by the large-file fixtures in each live client before
 * general release.
 */
export const DEFAULT_LIMITS: Limits = {
  maxFileBytes: 25 * MB,
  maxTotalBytes: 25 * MB,
  maxFiles: 50,
  maxNameLength: 255,
  readConcurrency: 3,
  readTimeoutMs: 120_000,
};

export function formatMegabytes(bytes: number): string {
  return `${Math.round((bytes / MB) * 10) / 10} MB`;
}
