# CLAUDE.md

## Latest handoff notes - 2026-06-10

This file is the handoff note for Claude/Codex in later chats. Read this section first.

### Current Git / workspace caution

The working tree may contain local changes from both Codex and Claude. Do not reset or revert them unless the user explicitly asks.

At the time this note was written, local GAS files had pending edits:

- `Code.js`
- `Index.html`
- `docs/index.html`
- `docs/config.js`
- `docs/README.md`
- `docs/push/onesignal/OneSignalSDKWorker.js`

There may also be local `docs/` edits from Claude or follow-up testing. Check `git status --short` before staging anything, and stage only the files related to the current task.

### Latest GitHub Pages launcher commit

Codex pushed this commit:

```text
3d416b0 Handle iOS in-app Safari install flow
```

What it changed:

- iPad/iPhone opened from LINE/Facebook/in-app browsers no longer shows a direct GAS launch button.
- The launcher now detects in-app browsers and shows Thai instructions to open the GitHub Pages link in Safari first.
- Direct GAS launch button is shown only when the page is actually opened in Safari.
- Home Screen standalone mode still embeds the GAS app full-screen.
- Android Chrome intent flow remains in place.
- Manifest start URL cache marker changed to `ios-line-safari-1`.
- Service worker cache changed to `ipad-jv-wrapper-v19`.

Test URL:

```text
https://wisanu15.github.io/IPAD_JV/index.html?v=ios-line-safari-1
```

Important diagnosis from the latest iPad screenshot:

- The top bar still showed `LINE` and an `X`, so it was not real Safari.
- Opening the GAS URL from that webview caused the Google Drive "unable to open file" error.
- Correct iPad flow is: open GitHub Pages in real Safari, Share, Add to Home Screen, then open the `iPad JV` icon.

### Admin remembered-login change

The user asked: "make admin remember automatically; no checkbox."

Codex already edited local `Index.html` so admin remember tokens are saved automatically after admin login:

- Removed the `admin-remember-device` checkbox from the admin login UI.
- Changed `saveAdminRememberIfRequested()` so it always calls `createRememberForCurrentSession()` when `_authTok` exists.
- Logout behavior should still clear the remembered token.

This GAS change may still need to be pushed/deployed depending on what Claude has done after this handoff. Verify before deploying:

```powershell
rg -n "admin-remember-device|saveAdminRememberIfRequested|createRememberForCurrentSession" Index.html Code.js
```

Expected current `saveAdminRememberIfRequested()` behavior:

```js
function saveAdminRememberIfRequested() {
  if (!_authTok) return;
  gsr
    .withSuccessHandler(function(r) {
      if (r && r.success && r.rememberToken) saveRememberedAdminToken(r.rememberToken);
    })
    .withFailureHandler(function(){})
    .createRememberForCurrentSession();
}
```

### Admin mobile push notification change

The user asked to add mobile notifications to the installed GitHub Pages PWA when students submit an iPad issue.

Codex added a OneSignal Web Push integration:

- `docs/config.js` now has `ONESIGNAL_APP_ID`.
- `docs/index.html` loads the OneSignal Web SDK, initializes a dedicated OneSignal service worker, and shows a bell button only inside standalone installed mode after OneSignal is configured.
- `docs/push/onesignal/OneSignalSDKWorker.js` imports the OneSignal v16 service worker.
- `Code.js` has `notifyAdminPush_()` and `submitIssue()` calls it after the existing email notification.
- `docs/README.md` contains setup steps.

Required OneSignal setup:

1. Create a OneSignal Web Push app for `https://wisanu15.github.io`.
2. In OneSignal Web Settings, enable custom service worker settings:
   - Path to service worker files: `/IPAD_JV/push/onesignal/`
   - Service worker filename: `OneSignalSDKWorker.js`
   - Service worker registration scope: `/IPAD_JV/push/onesignal/`
3. Put the Web Push App ID in `docs/config.js` as `ONESIGNAL_APP_ID`.
4. Add Apps Script Project Settings > Script Properties:
   - `ONESIGNAL_APP_ID`
   - `ONESIGNAL_REST_API_KEY`
   - `IPAD_JV_PWA_URL` = `https://wisanu15.github.io/IPAD_JV/`
5. Publish GitHub Pages and deploy a new Apps Script web app version.
6. Open the installed app as an admin, tap the bell button once, and allow notifications.

Do not put the OneSignal REST API key in `docs/config.js`; it must stay in Apps Script Script Properties.

## Project overview

This project is a Google Apps Script web app for the iPad JV management system. The main Apps Script files are:

- `Code.js` - server-side Google Apps Script logic
- `Index.html` - main web app UI served by Apps Script
- `appsscript.json` - Apps Script manifest/scopes/web app config

There is also a GitHub Pages launcher/PWA wrapper in `docs/`:

- `docs/index.html` - mobile launcher page
- `docs/config.js` - points to the deployed Apps Script Web App URL
- `docs/manifest.webmanifest` - PWA manifest
- `docs/sw.js` - service worker
- `docs/push/onesignal/OneSignalSDKWorker.js` - OneSignal Web Push service worker
- `docs/icons/` - launcher icons

## Important separation

Do not push `docs/` files into Apps Script. They are only for GitHub Pages.

`.claspignore` intentionally restricts clasp uploads to:

- `appsscript.json`
- `Code.js`
- `Index.html`

This prevents errors like:

```text
ReferenceError: window is not defined (file "docs/config")
```

That happened once because `docs/config.js` was accidentally uploaded to Apps Script and GAS tried to execute browser-only code server-side.

## Current deployment notes

The GitHub Pages site is:

```text
https://wisanu15.github.io/IPAD_JV/
```

The launcher currently points to the Apps Script URL using the account-routing form:

```text
https://script.google.com/a/*/macros/s/AKfycby079h5QsFuJtoOWk9E2-jJK1uhxnjlpw1Jrg-HthTxa5CeN15CRPyZCVxvHRYYdstmUQ/exec
```

The `/a/*/` form is intentional. It helps Android Chrome when multiple Google accounts are present. Without it, some Android devices showed a Google Drive error page.

## Mobile launcher behavior

iPhone/iPad:

- Uses Apple Home Screen metadata/icons.
- When opened from the Home Screen in standalone mode, it loads the GAS app in an iframe.
- iPad has been confirmed working.

Android:

- Direct GAS access works after fixing Chrome account/cache issues.
- Incognito working confirmed that the deployment itself was not broken.
- Launcher uses Chrome intent routing when opened from LINE/in-app browsers.
- Current flow is:
  1. Open GitHub Pages launcher from LINE.
  2. Tap `เปิดระบบใน Chrome`.
  3. It opens the launcher in Chrome first.
  4. User can tap `ติดตั้ง`.
  5. Then use `เปิดระบบใน Chrome`.

The latest launcher cache/version marker used in the wrapper is:

```text
ios-line-safari-1
```

The latest service worker cache name is:

```text
ipad-jv-wrapper-v19
```

## Recent feature work

Admin remembered login was added and then changed to automatic remember:

- Admin login page no longer has a remember-device checkbox.
- Admin remember tokens are saved automatically after successful admin password login.
- The client stores a random remember token in `localStorage`.
- Passwords are not stored locally.
- Server stores remember tokens in `PropertiesService`.
- Logout clears the remembered token locally and server-side.

Relevant server functions in `Code.js`:

- `createRememberToken_`
- `consumeRememberToken`
- `forgetRememberToken`
- `createRememberForCurrentSession`

Relevant client functions in `Index.html`:

- `getRememberedAdminToken`
- `saveRememberedAdminToken`
- `clearRememberedAdminToken`
- `saveAdminRememberIfRequested`
- `tryRememberedAdminLogin`

## Completed requested GAS change

The user wanted admin remembered login to happen automatically, without requiring the admin to tick a checkbox.

Implemented behavior:

1. Keep the server-side remember-token system as-is.
2. In `Index.html`, make `saveAdminRememberIfRequested()` always save a remember token whenever an admin login succeeds.
3. Remove or hide the `admin-remember-device` checkbox from the admin login UI.
4. Keep logout behavior unchanged: `confirmLogout()` should still call `clearRememberedAdminToken()` so logout removes remembered login.

Current expected function:

```js
function saveAdminRememberIfRequested() {
  if (!_authTok) return;
  gsr
    .withSuccessHandler(function(r) {
      if (r && r.success && r.rememberToken) saveRememberedAdminToken(r.rememberToken);
    })
    .withFailureHandler(function(){})
    .createRememberForCurrentSession();
}
```

## Apps Script deployment caution

This environment can `clasp push`, but updating an existing web app deployment may fail unless the account is in the same domain as the script owner.

If deployment update fails with:

```text
Only users in the same domain as the script owner may deploy this script.
```

then the script owner must manually update the Web App deployment:

1. Apps Script editor
2. Deploy
3. Manage deployments
4. Select the existing Web App
5. Edit
6. Choose the newest version
7. Deploy

## Useful commands

Use full Git path if `git` is not in PATH:

```powershell
& 'C:\Program Files\Git\cmd\git.exe' status --short --branch
```

Use `clasp.cmd` instead of `clasp.ps1` if PowerShell execution policy blocks scripts:

```powershell
& 'C:\Users\Kru_E\AppData\Local\Programs\node-portable\clasp.cmd' push
```

Check GitHub Pages live files with cache-busting URLs, for example:

```text
https://wisanu15.github.io/IPAD_JV/index.html?v=latest
```

## Important user preference/context

The user wants the Android install flow to be simple:

- From LINE, tap a button to open Chrome.
- Show install first in Chrome.
- Keep the launcher simple.

Avoid reintroducing extra install-flow buttons unless asked. The OneSignal bell button was explicitly requested for admin push notifications and should remain hidden unless `ONESIGNAL_APP_ID` is configured and the app is opened in standalone installed mode.
