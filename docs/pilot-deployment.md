# Pilot deployment and rollback (tenant administrator)

The add-in is a single XML manifest. Code and icons are served from **https://douglasmsim.github.io/OutlookPlugin/** (GitHub Pages, built by `.github/workflows/deploy.yml` from the `main` branch).

## Deploy to an IT test group

Requires Global Administrator or Exchange Administrator.

1. Download `manifest.xml` from https://douglasmsim.github.io/OutlookPlugin/manifest.xml (identical to `manifests/manifest.pilot.xml` in the repo).
2. Microsoft 365 admin center → **Settings → Integrated apps → Upload custom apps**.
3. App type: **Office Add-in**. Choose **Upload manifest file (.xml) from device** and select the file.
4. **Assign users**: choose **Specific users/groups** and pick the IT test group. Don't choose "Entire organization" for the pilot.
5. Accept the permission request. The add-in asks for **Read item** only: it can read the selected message and its attachments, and cannot send mail or read other items.
6. Deploy. It can take up to 24 hours (usually less) to reach every client. Users restart Outlook to pick it up.

Microsoft's reference: https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins

## Rollback

| Goal | Action | Effect |
|---|---|---|
| Remove the add-in from the pilot group | Integrated apps → the add-in → **Users** → remove the group or user assignment (or **Remove app**) | Button disappears after propagation and restart. Nothing is stored server-side to clean up. |
| Roll back a bad code release | Revert the commit on `main` and push, or re-run an earlier successful "Build, test and deploy" workflow | Pages serves the previous code within minutes. No manifest change needed. |
| Emergency stop for everyone using this host | Repository Settings → Pages → **Unpublish site** | The button stops working everywhere (Outlook cannot load its code). Use only if removal by assignment is too slow. |

## Updating

- **Code-only changes** (bug fixes): merge to `main`. The workflow tests, builds and redeploys. Users get the new code on their next click after Outlook reloads the add-in. No admin action needed.
- **Manifest changes** (button label, icons, permissions, requirement sets): bump `version` in `package.json`, run `npm run manifests`, merge, then upload the new manifest in Integrated apps (**Update**). A permission increase requires admin consent again.

The pilot add-in ID (`7db00ae2-968a-4085-b891-51456ac34d44`) must stay the same for the life of the product. Changing it creates a different add-in.

## Production hosting

The brief names Azure Static Web Apps as the production host. Moving there means changing `pilot.origin` in `config/environments.json` (or adding a `production` environment), regenerating the manifests, and redeploying the new manifest through Integrated apps.
