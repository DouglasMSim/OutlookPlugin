# Reply All with Attachments

An Outlook add-in that adds a **Reply All with Attachments** button to received messages. One click opens Outlook's normal, editable Reply All with the selected message's files already attached. The sender can add or remove recipients, remove a file, edit the body, and send when ready.

- Outlook web add-in (Office.js, TypeScript), add-in-only XML manifest, `MessageReadCommandSurface` button, `ExecuteFunction` handler.
- Requires Outlook API **Mailbox 1.15**: classic Outlook for Windows Version 2412 (Build 18324.20172) or later, new Outlook for Windows, and Outlook on the web.
- Mail and file contents stay inside Outlook. Nothing is sent to any server.

**Status:** pilot build. Automated checks pass; live Outlook client tests have not been run yet ([test results](docs/test-results.md)). Licensing, the admin portal and billing are not in this build.

## Install the pilot

Open **https://douglasmsim.github.io/OutlookPlugin/**, download the manifest, and follow the steps on that page. Full walkthrough: [docs/sideload-guide.md](docs/sideload-guide.md). Admin deployment to a test group: [docs/pilot-deployment.md](docs/pilot-deployment.md).

## Documentation

| Doc | What it covers |
|---|---|
| [docs/pilot-notes.md](docs/pilot-notes.md) | Behaviour, the v1 "nothing partial without asking" boundary, limits, privacy, what the add-in can't observe |
| [docs/supported-clients.md](docs/supported-clients.md) | Supported clients and versions, where the button appears, why it may be missing |
| [docs/sideload-guide.md](docs/sideload-guide.md) | Installing the pilot, the owner walkthrough, local HTTPS development, debugging |
| [docs/test-plan.md](docs/test-plan.md) | Live fixtures to run in each client |
| [docs/test-results.md](docs/test-results.md) | What has actually been run, and the live results table |
| [docs/pilot-deployment.md](docs/pilot-deployment.md) | Integrated Apps deployment, updates and rollback |

## Development

Requires Node.js 20+.

```bash
npm ci
```

```bash
npm run verify
```

`npm run verify` runs type checking, unit tests, the manifest up-to-date check, offline manifest validation, the production build and the browser smoke tests. The browser tests need Chromium for `playwright-core`; on a new machine run `npx playwright-core install chromium` once.

| Script | Purpose |
|---|---|
| `npm start` | Build the dev bundle and serve it at `https://localhost:3000` (run `npm run dev-certs` once first) |
| `npm run build` | Production build of the pilot into `dist/` |
| `npm test` | Unit tests (Office.js mocked) |
| `npm run test:browser` | Built bundle in headless Chromium with simulated Office.js |
| `npm run manifests` | Regenerate `manifests/manifest.dev.xml` and `manifest.pilot.xml` from the template |
| `npm run validate:manifest -- --online` | Also run Microsoft's online manifest validator |

### Layout

```
src/core/        Pure logic: classification, limits, all-or-ask decision, mapping, logging
src/commands/    Button handler and the command flow (runs per click)
src/office/      Office.js adapters: reading attachments, opening the reply, notifications, dialog
src/dialog/      Confirmation dialog page logic
src/pages/       HTML, CSS, install page, fallback page
manifests/       Manifest template and generated dev/pilot manifests
config/          Origins and add-in IDs per environment
assets/          Icon source (SVG) and rendered PNGs (node scripts/icons.mjs)
test/unit/       Vitest unit tests
test/browser/    Playwright smoke tests and the simulated Office.js
```

### Hosting

`.github/workflows/deploy.yml` tests and builds every push, deploys `main` to GitHub Pages, then runs Microsoft's online manifest validation. To host elsewhere (for example Azure Static Web Apps), change `pilot.origin` in `config/environments.json`, run `npm run manifests`, and deploy `dist/`.
