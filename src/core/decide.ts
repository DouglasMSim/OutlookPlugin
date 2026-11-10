import type { AttachmentPlan } from "./classify";
import type { Limits } from "./limits";
import { decodedByteLength } from "./base64";
import type { Base64ReplyAttachment, Issue, ReadFile, ReadOutcome } from "./types";

export type Decision =
  /** Nothing to carry over: open an ordinary Reply All. */
  | { kind: "plain-reply" }
  /** Every eligible file was read: open one Reply All with all of them. */
  | { kind: "reply-with-files"; files: ReadFile[] }
  /** Something cannot be included: stop and ask before opening anything. */
  | { kind: "needs-confirmation"; available: ReadFile[]; issues: Issue[] };

/**
 * The all-or-ask rule. A reply with files opens only when every attachment
 * that should be carried over was read successfully and fits the limits.
 * Otherwise the caller must ask the sender before opening a partial reply.
 *
 * `outcomes[i]` is the read result for `plan.toRead[i]`.
 */
export function decide(plan: AttachmentPlan, outcomes: readonly ReadOutcome[], limits: Limits): Decision {
  if (outcomes.length !== plan.toRead.length) {
    throw new Error("Each planned file needs exactly one read outcome.");
  }

  const issues: Issue[] = [...plan.issues];
  const available: ReadFile[] = [];
  let totalBytes = 0;

  plan.toRead.forEach(({ index, info }, i) => {
    const outcome = outcomes[i] as ReadOutcome;
    if (!outcome.ok) {
      issues.push({ index, name: info.name, reason: outcome.reason });
      return;
    }
    // Re-check limits against the real decoded size; reported sizes can be approximate.
    const bytes = decodedByteLength(outcome.base64);
    if (bytes > limits.maxFileBytes) {
      issues.push({ index, name: info.name, reason: "file-too-large" });
      return;
    }
    if (totalBytes + bytes > limits.maxTotalBytes) {
      issues.push({ index, name: info.name, reason: "total-too-large" });
      return;
    }
    totalBytes += bytes;
    available.push({ index, name: info.name, base64: outcome.base64, bytes });
  });

  issues.sort((a, b) => a.index - b.index);

  if (issues.length > 0) return { kind: "needs-confirmation", available, issues };
  if (available.length === 0) return { kind: "plain-reply" };
  return { kind: "reply-with-files", files: available };
}

export type ConfirmChoice = "continue" | "cancel";

/** What to open after the sender answers the confirmation. */
export function resolveConfirmation(
  decision: Extract<Decision, { kind: "needs-confirmation" }>,
  choice: ConfirmChoice,
): Exclude<Decision, { kind: "needs-confirmation" }> | undefined {
  if (choice === "cancel") return undefined;
  return decision.available.length > 0
    ? { kind: "reply-with-files", files: decision.available }
    : { kind: "plain-reply" };
}

/** Map read files to Base64 ReplyFormAttachment objects, preserving order and names. */
export function toReplyAttachments(files: readonly ReadFile[]): Base64ReplyAttachment[] {
  return files.map((file) => ({
    type: "base64",
    name: file.name,
    base64file: file.base64,
    inLine: false,
  }));
}
