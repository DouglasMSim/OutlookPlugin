import { decodeCounts, parseDialogDetails, type DialogDetails, type DialogToParent } from "../core/dialog-protocol";

function send(message: DialogToParent): void {
  Office.context.ui.messageParent(JSON.stringify(message));
}

function element<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing #${id}`);
  return el as T;
}

function plural(count: number, word: string): string {
  return `${count} ${word}${count === 1 ? "" : "s"}`;
}

function renderSummary(availableCount: number, issueCount: number): void {
  element("summary").textContent =
    issueCount === 1
      ? "1 attachment can't be included in this reply."
      : `${issueCount} attachments can't be included in this reply.`;
  element("continue").textContent =
    availableCount > 0
      ? `Continue with ${plural(availableCount, "file")}`
      : "Open Reply All without attachments";
  element("available").textContent =
    availableCount > 0
      ? `${plural(availableCount, "other file")} can be attached.`
      : "No other files can be attached.";
}

function renderDetails(details: DialogDetails): void {
  renderSummary(details.availableCount, details.issues.length);
  const list = element<HTMLUListElement>("issues");
  list.replaceChildren(
    ...details.issues.map((issue) => {
      const li = document.createElement("li");
      const name = document.createElement("span");
      name.className = "name";
      name.textContent = issue.name || "(unnamed attachment)";
      const reason = document.createElement("span");
      reason.className = "reason";
      reason.textContent = issue.text;
      li.append(name, reason);
      return li;
    }),
  );
  list.hidden = details.issues.length === 0;
}

Office.onReady(() => {
  const counts = decodeCounts(window.location.hash);
  if (counts) renderSummary(counts.availableCount, counts.issueCount);

  element("continue").addEventListener("click", () => send({ type: "continue" }));
  element("cancel").addEventListener("click", () => send({ type: "cancel" }));

  try {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      if (!("message" in arg)) return;
      const details = parseDialogDetails(arg.message);
      if (details) renderDetails(details);
    });
  } catch {
    // DialogApi 1.2 not available: the counts-only summary stays.
  }

  element<HTMLButtonElement>("continue").disabled = false;
  element<HTMLButtonElement>("cancel").disabled = false;
  send({ type: "ready" });
});
