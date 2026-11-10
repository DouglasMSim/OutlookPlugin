# Supported clients

The add-in needs Outlook API requirement set **Mailbox 1.15**, the first version where a reply form accepts Base64 file attachments.

| Client | Minimum version | Where the button appears | Live status |
|---|---|---|---|
| Classic Outlook for Windows (Microsoft 365) | Version 2412 (Build 18324.20172) | Home ribbon in the reading pane; Message tab when the message is open in its own window | Untested |
| New Outlook for Windows | Current release (supports 1.15) | **Apps** button, or the **…** menu on the message | Untested |
| Outlook on the web (Microsoft 365 work/school) | Current release (supports 1.15) | **Apps** button, or the **…** menu on the message | Untested |
| Outlook on Mac, iOS, Android | Not targeted in the pilot | | Not tested |

"Untested" means no live run has been recorded yet in [test-results.md](test-results.md). Update this table as fixtures pass.

Mailboxes: Microsoft 365 (Exchange Online) primary mailboxes and shared mailboxes accessed with the employee's own account. The manifest sets `SupportsSharedFolders`. Shared mailboxes must be tested in each client.

## Checking the classic Outlook version

**File → Office Account → About Outlook.** The first line shows `Version 2412 (Build 18324.20172 Click-to-Run)` or later. Monthly Enterprise and Semi-Annual channels reach a given build later than Current Channel.

## Why the button might be missing

| Symptom | Cause |
|---|---|
| A small "Reply All with Attachments" add-in pane or bar appears instead of a button | Outlook is older than Mailbox 1.15. It ignores the button definition and falls back to the basic add-in page, which says the version is too old. Update Outlook. |
| Nothing appears at all | Add-in not installed in this mailbox, or installation not propagated yet (wait a few minutes and restart Outlook). Or the tenant blocks personal add-ins; use the admin deployment in [pilot-deployment.md](pilot-deployment.md). |
| Missing on one message only | The add-in only activates on **received messages in read mode**: not on drafts, meeting invites or appointments. Rights-protected (IRM) messages may also not activate add-ins depending on client and policy; record this in the protected-message fixture. |
| New Outlook / web: not in the toolbar | Add-in buttons live under **Apps** or **…**. In Outlook on the web you can pin it via **Settings → Mail → Customize actions**. |
