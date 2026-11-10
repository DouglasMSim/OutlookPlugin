import { planAttachments } from "../core/classify";
import { mapWithConcurrency } from "../core/concurrency";
import { decide, resolveConfirmation, toReplyAttachments, type ConfirmChoice, type Decision } from "../core/decide";
import type { Limits } from "../core/limits";
import { logEvent } from "../core/log";
import { describeIssue, TEXT } from "../core/messages";
import type { AttachmentInfo, Base64ReplyAttachment, ReadOutcome } from "../core/types";

export type OpenResult = { ok: true } | { ok: false; code: string };

/** The selected message, captured once when the button is clicked. */
export interface MailItemPort {
  readonly itemId: string | undefined;
  readonly attachments: readonly AttachmentInfo[];
  readAttachment(id: string, timeoutMs: number): Promise<ReadOutcome>;
  /** Open Outlook's native Reply All. Must be called at most once per click. */
  openReplyAll(attachments: Base64ReplyAttachment[]): Promise<OpenResult>;
  showProgress(message: string): Promise<void>;
  showError(message: string): Promise<void>;
  clearStatus(): Promise<void>;
}

export interface ConfirmRequest {
  availableCount: number;
  issues: { name: string; text: string }[];
}

export interface HostPort {
  supportsRequiredApi(): boolean;
  /** Snapshot of the message the command was invoked on. */
  getItem(): MailItemPort | undefined;
  /** ID of the message Outlook currently has selected, to detect a change mid-run. */
  currentItemId(): string | undefined;
  confirm(request: ConfirmRequest): Promise<ConfirmChoice | "unavailable">;
}

export type Outcome =
  | "unsupported"
  | "no-item"
  | "busy"
  | "opened-plain"
  | "opened-with-files"
  | "cancelled"
  | "confirm-unavailable"
  | "selection-changed"
  | "open-failed"
  | "error";

async function quietly(action: () => Promise<void>): Promise<void> {
  try {
    await action();
  } catch {
    // Status messages are best effort; they must never block completion.
  }
}

/** One invocation of the button, from snapshot to (at most) one reply form. */
export async function runCommand(host: HostPort, limits: Limits): Promise<Outcome> {
  const started = Date.now();
  logEvent("start");

  const item = host.getItem();
  if (!item) {
    logEvent("no-item");
    return "no-item";
  }
  if (!host.supportsRequiredApi()) {
    logEvent("unsupported-host");
    await quietly(() => item.showError(TEXT.unsupported));
    return "unsupported";
  }

  const plan = planAttachments(item.attachments, limits);
  logEvent("plan", {
    attachments: plan.totalAttachments,
    toRead: plan.toRead.length,
    inlineGraphics: plan.inlineGraphics,
    issues: plan.issues.length,
  });

  if (plan.toRead.length > 0) await quietly(() => item.showProgress(TEXT.progress(plan.toRead.length)));

  const outcomes = await mapWithConcurrency(plan.toRead, limits.readConcurrency, ({ info }) =>
    item.readAttachment(info.id, limits.readTimeoutMs).catch((): ReadOutcome => ({ ok: false, reason: "read-failed" })),
  );
  const decision = decide(plan, outcomes, limits);
  logEvent("read-done", {
    ok: outcomes.filter((o) => o.ok).length,
    failed: outcomes.filter((o) => !o.ok).length,
    ms: Date.now() - started,
  });

  let final: Exclude<Decision, { kind: "needs-confirmation" }>;
  if (decision.kind === "needs-confirmation") {
    await quietly(() => item.clearStatus());
    logEvent("confirm-shown", { available: decision.available.length, issues: decision.issues.length });
    const choice = await host.confirm({
      availableCount: decision.available.length,
      issues: decision.issues.map((issue) => ({ name: issue.name, text: describeIssue(issue.reason, limits) })),
    });
    if (choice === "unavailable") {
      logEvent("confirm-unavailable");
      await quietly(() => item.showError(TEXT.confirmUnavailable(decision.issues.length)));
      return "confirm-unavailable";
    }
    const resolved = resolveConfirmation(decision, choice);
    if (!resolved) {
      logEvent("cancelled");
      return "cancelled";
    }
    final = resolved;
  } else {
    final = decision;
  }

  // Never open a reply for a different message than the one the sender clicked on.
  if (item.itemId !== undefined && host.currentItemId() !== item.itemId) {
    logEvent("selection-changed");
    await quietly(() => item.showError(TEXT.selectionChanged));
    return "selection-changed";
  }

  const attachments = final.kind === "reply-with-files" ? toReplyAttachments(final.files) : [];
  const result = await item.openReplyAll(attachments);
  if (!result.ok) {
    logEvent("open-failed", { files: attachments.length });
    await quietly(() => item.showError(TEXT.openFailed(result.code)));
    return "open-failed";
  }

  await quietly(() => item.clearStatus());
  logEvent("opened", { files: attachments.length, ms: Date.now() - started });
  return attachments.length > 0 ? "opened-with-files" : "opened-plain";
}

export interface CommandEvent {
  completed(): void;
}

/**
 * Wrap runCommand for Outlook: ignore repeat clicks while a run is in
 * progress, and complete the Office event exactly once on every exit path,
 * only after all asynchronous work has finished.
 */
export function createCommandHandler(getHost: () => HostPort, limits: Limits) {
  let busy = false;

  return async function handle(event: CommandEvent): Promise<Outcome> {
    let completed = false;
    const complete = () => {
      if (completed) return;
      completed = true;
      try {
        event.completed();
      } catch {
        // Nothing useful to do if Outlook rejects completion.
      }
    };

    if (busy) {
      logEvent("busy");
      complete();
      return "busy";
    }

    busy = true;
    let host: HostPort | undefined;
    try {
      host = getHost();
      return await runCommand(host, limits);
    } catch {
      logEvent("error");
      const item = (() => {
        try {
          return host?.getItem();
        } catch {
          return undefined;
        }
      })();
      if (item) await quietly(() => item.showError(TEXT.unexpected));
      return "error";
    } finally {
      busy = false;
      complete();
    }
  };
}
