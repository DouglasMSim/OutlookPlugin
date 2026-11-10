# Install and test the pilot build

## 1. Install the pilot add-in (no admin needed)

The pilot build is hosted at **https://douglasmsim.github.io/OutlookPlugin/**, which also has a download button for the manifest.

1. Download the manifest: https://douglasmsim.github.io/OutlookPlugin/manifest.xml (or `manifests/manifest.pilot.xml` from this repo).
2. Sign in to a **Microsoft 365 test mailbox** at https://aka.ms/olksideload. Outlook on the web opens with the add-ins window.
3. **My add-ins → Custom Addins → Add a custom add-in → Add from File**, choose the manifest, then **Install**.
4. The install belongs to the mailbox. Classic and new Outlook for Windows pick it up after a restart (allow a few minutes).

If **Add a custom add-in** is missing, tenant policy blocks personal add-ins. Ask the admin to deploy via Integrated Apps ([pilot-deployment.md](pilot-deployment.md)).

To remove: same window, **Custom Addins → … → Remove**.

Microsoft's guide: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing

## 2. Run the owner walkthrough

Use two controlled test mailboxes (**Sender** and **Tester**) plus at least two other controlled recipients (**CC** and **New**). Never use real client mail.

1. From **Sender**, send a message to **Tester**, CC **CC**, with a `.docx` and a `.pdf` attached. Before sending, record a SHA-256 hash of each file (commands below).
2. In **Tester**, open the message and click **Reply All with Attachments**.
3. Compare with Outlook's own **Reply All** on the same message: same To/Cc, same subject, same quoted history, same conversation.
4. Confirm both files are attached with their original names. Add **New** as a recipient, write a line, send.
5. In **New**'s mailbox, save both files and hash them. The hashes must match step 1.
6. Repeat in classic Outlook for Windows, new Outlook for Windows and Outlook on the web: first in Tester's primary mailbox, then in a shared mailbox Tester has access to. For the shared mailbox, check the **From** address on the received message.
7. Record each run in [test-results.md](test-results.md): client, Outlook build, mailbox type, date, pass/fail, notes. Leave anything not run as **Untested**.

### Hashing files

Windows PowerShell:

```powershell
Get-FileHash -Algorithm SHA256 "Draft Agreement.docx"
```

macOS / Linux:

```bash
shasum -a 256 "Draft Agreement.docx"
```

### Finding the build

- Classic Outlook: **File → Office Account → About Outlook**.
- New Outlook: **Settings → General → About Outlook**.
- Outlook on the web: **Settings → General → About**; note the browser and version too.

## 3. Developer setup (local HTTPS)

Requirements: Node.js 20+ and Git. On the machine running Outlook:

```bash
git clone https://github.com/DouglasMSim/OutlookPlugin.git
```

```bash
cd OutlookPlugin && npm ci
```

```bash
npm run dev-certs
```

```bash
npm start
```

`npm run dev-certs` installs a trusted localhost certificate once. `npm start` builds the dev bundle and serves it at `https://localhost:3000`. Then sideload `manifests/manifest.dev.xml` exactly as in step 1. It installs as **Reply All with Attachments (Dev)** with its own add-in ID, so it can sit next to the pilot build.

The dev build only works on the computer serving `localhost:3000`.

### Debugging

The button is a function command (no visible UI), so:

- **Outlook on the web:** click the button once so its code loads, then open the browser's developer tools (F12). The code runs in a hidden frame loaded from `commands.html`. See [Debug a function command](https://learn.microsoft.com/office/dev/add-ins/testing/debug-function-command).
- **New Outlook for Windows:** close it, then start it with `olk.exe --devtools` to get Edge DevTools. See [Develop Outlook add-ins for the new Outlook on Windows](https://learn.microsoft.com/office/dev/add-ins/outlook/one-outlook#debug-your-add-in).
- **Classic Outlook for Windows:** set the `UseDirectDebugger` value to `1` under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` (use the dev add-in ID from `manifests/manifest.dev.xml`) and attach VS Code or Edge. See [Debug function commands in Outlook add-ins](https://learn.microsoft.com/office/dev/add-ins/outlook/debug-ui-less).

Console output only contains event codes and counts, for example `[reply-all-with-attachments] opened {"files":2,"ms":412}`.
