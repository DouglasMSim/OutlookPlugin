/**
 * Messages exchanged between the command runtime and the confirmation dialog.
 *
 * File names travel only through Office's in-memory messageChild channel.
 * They are never put in the dialog URL, so they never reach the web server.
 */

export type DialogToParent = { type: "ready" } | { type: "continue" } | { type: "cancel" };

export interface DialogDetails {
  type: "details";
  availableCount: number;
  issues: { name: string; text: string }[];
}

export function parseDialogToParent(raw: string): DialogToParent | undefined {
  try {
    const value: unknown = JSON.parse(raw);
    if (value && typeof value === "object" && "type" in value) {
      const type = (value as { type: unknown }).type;
      if (type === "ready" || type === "continue" || type === "cancel") return { type };
    }
  } catch {
    // Ignore malformed messages.
  }
  return undefined;
}

export function parseDialogDetails(raw: string): DialogDetails | undefined {
  try {
    const value = JSON.parse(raw) as Partial<DialogDetails>;
    if (
      value?.type === "details" &&
      typeof value.availableCount === "number" &&
      Array.isArray(value.issues) &&
      value.issues.every((i) => typeof i?.name === "string" && typeof i?.text === "string")
    ) {
      return value as DialogDetails;
    }
  } catch {
    // Ignore malformed messages.
  }
  return undefined;
}

/** Counts only (no names) for the URL fragment, used if messageChild is unavailable. */
export function encodeCounts(availableCount: number, issueCount: number): string {
  return `a=${availableCount}&i=${issueCount}`;
}

export function decodeCounts(fragment: string): { availableCount: number; issueCount: number } | undefined {
  const params = new URLSearchParams(fragment.replace(/^#/, "").split("?")[0]);
  const rawA = params.get("a");
  const rawI = params.get("i");
  if (rawA === null || rawI === null || rawA === "" || rawI === "") return undefined;
  const a = Number(rawA);
  const i = Number(rawI);
  if (!Number.isInteger(a) || !Number.isInteger(i) || a < 0 || i < 0) return undefined;
  return { availableCount: a, issueCount: i };
}
