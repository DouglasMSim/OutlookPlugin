# Live client test plan

Run every fixture in **each** client: classic Outlook for Windows, new Outlook for Windows, Outlook on the web. Run the core fixtures (F1–F3) in both a **primary** and a **shared** mailbox. Record results per client in [test-results.md](test-results.md), with the Outlook build and date. A passing form callback alone is not a pass: a fixture passes only when the received message is checked.

Use controlled test mailboxes only. Hash files before sending and after receiving (see [sideload-guide.md](sideload-guide.md#hashing-files)).

## Fixtures

| ID | Setup | Steps | Pass when |
|---|---|---|---|
| F1 | Message with a `.docx` and a `.pdf`, several To and Cc recipients, and a custom Reply-To | Click the button; compare with native Reply All on the same message; send | Same recipients (honouring Reply-To), subject and thread as native Reply All. Both files arrive with identical names and SHA-256 hashes. |
| F2 | Same as F1 | Add a recipient who never got the original; send | The new recipient receives both files, hashes match. |
| F3 | Same as F1 | Remove one file in the compose window; send | The sent message contains only the remaining file. |
| F4 | Attached `.png` plus an HTML signature with an inline logo | Click; send | The PNG arrives as a file. The logo is not added to the file list; it stays in the quoted body. |
| F5 | No attachments | Click | An ordinary Reply All opens, no dialog. |
| F6 | Only inline images (signature logo) | Click | An ordinary Reply All opens, no extra files. |
| F7 | Two different files with the same name | Click; send | Both arrive, each with its own hash. |
| F8 | Unicode names (`Résumé – 2025.docx`, `契約書.pdf`, emoji) | Click; send | Names and hashes match. |
| F9 | A PDF that a Mac Mail sender attached "inline" | Click; send | The PDF is carried over. |
| F10 | One ordinary file plus an attached email (drag an email into a new message) | Click | Record whether the button is shown. Dialog lists the attached email. **Continue** opens a reply with only the ordinary file; **Cancel** opens nothing. |
| F11 | Cloud attachment (OneDrive "share link" attachment) plus one ordinary file | Click | Same as F10, with the cloud file listed. |
| F12 | Rights-protected (IRM / "Do Not Forward" / sensitivity label with encryption) message with attachments | Open it | Record whether Outlook shows the button at all. If it does, no silent omission or partial reply without the dialog. |
| F13 | S/MIME encrypted message with attachments | Open it | As F12. |
| F14 | Files totalling about 20 MB, then about 24 MB | Click; send; time it | Reply opens; files arrive intact. Note time and any Outlook errors. |
| F15 | One file over 25 MB, or files totalling over 25 MB | Click | Dialog lists the over-limit files. Nothing is read or opened before the choice. |
| F16 | Any message with files | Double- and triple-click the button quickly | Exactly one reply opens. |
| F17 | Message with several large files | Click, then immediately select a different message | Either the reply opens for the **original** message, or nothing opens and an error says the selection changed. Never a reply to the wrong message. |
| F18 | Any message with files | Click; close the draft without sending | Normal Outlook discard/save-draft behaviour; nothing is sent. |
| F19 | Conversation view on, several messages with different files in the thread | Select a middle message; click | Only that message's files are attached. |
| F20 | Message opened in its own window (not the reading pane) | Click from that window | Correct message and files. |
| F21 | Old classic Outlook build (< Version 2412), if available | Open a message | Record what appears. Expected: no button; the fallback pane says the version is too old. |
| F22 | Shared mailbox (F1 to F3 again) | As F1 | As F1, and the received message shows the expected **From** identity. |
| F23 | Admin deployment via Integrated Apps to the pilot group | Sign in as a pilot user who never self-installed | Record whether the button is already pinned next to Reply / Reply all / Forward in new Outlook and on the web, or must be pinned by hand. |

## Failed reads

A real read failure is hard to force. Ways to try: open a message, go offline (or disconnect the network) before clicking, and see whether a cached client can still read the files. If you can't reproduce one, the failure path is covered by the automated tests (unit + browser) only. Record it as such.

## What to record per run

Client, Outlook build/version, OS or browser, mailbox type (primary/shared), fixture ID, date, result (Pass/Fail/Untested/Not applicable), and notes (error codes shown, timings for F14).
