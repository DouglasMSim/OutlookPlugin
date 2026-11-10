import { DEFAULT_LIMITS } from "../core/limits";
import { createOfficeHost } from "../office/outlook";
import { createCommandHandler } from "./command";

const handle = createCommandHandler(createOfficeHost, DEFAULT_LIMITS);

/** ExecuteFunction entry point named in the manifest's <FunctionName>. */
function replyAllWithAttachments(event: Office.AddinCommands.Event): void {
  void handle(event);
}

Office.onReady(() => {
  // Nothing to initialise; the handler runs per click.
});

// Newer clients use Office.actions.associate; older classic Outlook builds
// look the function up on the global object. Register both.
if (typeof Office !== "undefined" && Office.actions && typeof Office.actions.associate === "function") {
  Office.actions.associate("replyAllWithAttachments", replyAllWithAttachments);
}
(globalThis as unknown as Record<string, unknown>).replyAllWithAttachments = replyAllWithAttachments;
