/**
 * This page only appears when Outlook ignores the add-in command (for example
 * an Outlook build older than Mailbox 1.15) and falls back to the base
 * manifest. It explains why the button is missing.
 */
Office.onReady(() => {
  const status = document.getElementById("status");
  if (!status) return;
  let supported = false;
  try {
    supported = Office.context.requirements.isSetSupported("Mailbox", "1.15");
  } catch {
    supported = false;
  }
  status.textContent = supported
    ? "This Outlook supports the add-in. If the button is missing, look under Apps or the … menu on the message."
    : "This Outlook version is too old for this add-in (it needs Outlook API Mailbox 1.15). Classic Outlook for Windows needs Version 2412 (Build 18324.20172) or later.";
});
