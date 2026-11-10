# Test results

## Automated checks

Run on 2026-11-10 in a Linux cloud build environment (Node 22.23, Chromium 153 headless), against commit at the time of the initial pilot build. The same checks run in GitHub Actions on every push (`.github/workflows/deploy.yml`).

| Check | Command | Result |
|---|---|---|
| Type checking | `npm run typecheck` | Pass |
| Unit tests: classification, limits, all-or-ask decision, mapping, Base64, ordering/concurrency, command flow (repeat clicks, completion on every path, changed selection, no retry), privacy, manifest structure | `npm test` | Pass (56 tests) |
| Generated manifests up to date | `npm run manifests -- --check` | Pass |
| Manifests are well-formed XML | `npm run validate:manifest` | Pass |
| Production build | `npm run build` | Pass |
| Built bundle in Chromium with simulated Office.js, including the real confirmation dialog page | `npm run test:browser` | Pass (6 scenarios) |
| Microsoft's online manifest validation | `npm run validate:manifest -- --online` | Not run here: the build environment's network blocks the validation service. Runs in GitHub Actions after each deploy. |
| Clean-checkout setup (`npm ci` + `npm run verify` from a fresh clone) | see README | Pass |

Simulated Office.js tests prove the add-in's own logic and wiring. They do **not** prove Outlook behaviour: that requires the live fixtures below.

## Live client results

Nothing below has been run yet. Fill in as fixtures are executed (see [test-plan.md](test-plan.md)). Leave **Untested** for anything not run.

| Fixture | Classic Windows (primary) | Classic Windows (shared) | New Windows (primary) | New Windows (shared) | Web (primary) | Web (shared) |
|---|---|---|---|---|---|---|
| F1 DOCX+PDF, recipients, Reply-To | Untested | Untested | Untested | Untested | Untested | Untested |
| F2 Added recipient | Untested | Untested | Untested | Untested | Untested | Untested |
| F3 Removed file | Untested | Untested | Untested | Untested | Untested | Untested |
| F4 PNG + inline logo | Untested | | Untested | | Untested | |
| F5 No attachments | Untested | | Untested | | Untested | |
| F6 Inline only | Untested | | Untested | | Untested | |
| F7 Duplicate names | Untested | | Untested | | Untested | |
| F8 Unicode names | Untested | | Untested | | Untested | |
| F9 Inline PDF | Untested | | Untested | | Untested | |
| F10 Attached email | Untested | | Untested | | Untested | |
| F11 Cloud attachment | Untested | | Untested | | Untested | |
| F12 IRM-protected | Untested | | Untested | | Untested | |
| F13 S/MIME | Untested | | Untested | | Untested | |
| F14 Large (20/24 MB) | Untested | | Untested | | Untested | |
| F15 Over limit | Untested | | Untested | | Untested | |
| F16 Rapid clicks | Untested | | Untested | | Untested | |
| F17 Switch messages | Untested | | Untested | | Untested | |
| F18 Cancel draft | Untested | | Untested | | Untested | |
| F19 Conversation view | Untested | | Untested | | Untested | |
| F20 Separate window | Untested | | Untested | | Pass (pop-ups allowed) | |
| F21 Old build | Untested | | n/a | | n/a | |
| F22 Shared From identity | | Untested | | Untested | | Untested |

### Run log

| Date | Tester | Client + build | Mailbox | Fixtures | Result | Notes |
|---|---|---|---|---|---|---|
| 2026-11-10 | Owner | Outlook on the web, Microsoft Edge (build not recorded) | Primary | Informal check, F20 | Pass after fix | Real received message with one .docx. Reading pane: Reply All opened with the sender and the file attached. Separate window: nothing happened until pop-ups were allowed for `https://outlook.cloud.microsoft`, then it worked. Draft discarded; nothing sent, so delivery and bytes were not checked. |
| | | | | | | |
