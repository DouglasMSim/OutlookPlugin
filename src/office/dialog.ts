import type { ConfirmRequest } from "../commands/command";
import type { ConfirmChoice } from "../core/decide";
import { encodeCounts, parseDialogToParent, type DialogDetails } from "../core/dialog-protocol";

function dialogUrl(request: ConfirmRequest): string {
  const url = new URL("dialog.html", window.location.href);
  url.search = "";
  // Counts only. File names are sent later through messageChild, never in the URL.
  url.hash = encodeCounts(request.availableCount, request.issues.length);
  return url.href;
}

function supportsMessageChild(): boolean {
  try {
    return Office.context.requirements.isSetSupported("DialogApi", "1.2");
  } catch {
    return false;
  }
}

/**
 * Ask the sender whether to continue with the files that are available.
 * Resolves "unavailable" if Outlook cannot show the dialog at all.
 */
export function confirmWithDialog(request: ConfirmRequest): Promise<ConfirmChoice | "unavailable"> {
  return new Promise((resolve) => {
    let settled = false;
    let dialog: Office.Dialog | undefined;
    const finish = (choice: ConfirmChoice | "unavailable") => {
      if (settled) return;
      settled = true;
      try {
        dialog?.close();
      } catch {
        // Already closed.
      }
      resolve(choice);
    };

    const details: DialogDetails = { type: "details", ...request };

    try {
      Office.context.ui.displayDialogAsync(
        dialogUrl(request),
        { height: 55, width: 40, displayInIframe: true },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            finish("unavailable");
            return;
          }
          dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            if (!("message" in arg)) return;
            if (arg.origin && arg.origin !== window.location.origin) return;
            const message = parseDialogToParent(arg.message);
            if (!message) return;
            if (message.type === "ready") {
              if (supportsMessageChild()) {
                try {
                  dialog?.messageChild(JSON.stringify(details));
                } catch {
                  // The dialog falls back to counts-only text.
                }
              }
              return;
            }
            finish(message.type);
          });
          // The sender closed the dialog with its X button (or it failed): treat as cancel.
          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => finish("cancel"));
        },
      );
    } catch {
      finish("unavailable");
    }
  });
}
