# CLAUDE.md

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
line-install-first-1
```

The latest service worker cache name is:

```text
ipad-jv-wrapper-v12
```

## Recent feature work

Admin remembered login was added:

- Admin login page now has a remember-device checkbox.
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

## Pending requested GAS change

The user now wants admin remembered login to happen automatically, without requiring the admin to tick a checkbox.

Recommended implementation:

1. Keep the server-side remember-token system as-is.
2. In `Index.html`, make `saveAdminRememberIfRequested()` always save a remember token whenever an admin login succeeds.
3. Remove or hide the `admin-remember-device` checkbox from the admin login UI.
4. Keep logout behavior unchanged: `confirmLogout()` should still call `clearRememberedAdminToken()` so logout removes remembered login.

Current function to change:

```js
function saveAdminRememberIfRequested() {
  var cb = document.getElementById('admin-remember-device');
  if (!cb || !cb.checked || !_authTok) return;
  gsr
    .withSuccessHandler(function(r) {
      if (r && r.success && r.rememberToken) saveRememberedAdminToken(r.rememberToken);
    })
    .withFailureHandler(function(){})
    .createRememberForCurrentSession();
}
```

Suggested replacement:

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

Then remove or hide this admin-login checkbox block:

```html
<label style="display:flex;align-items:center;gap:8px;margin:4px 0 12px;font-size:13px;color:var(--muted);cursor:pointer">
  <input id="admin-remember-device" type="checkbox" checked>
  <span>จำเครื่องนี้ไว้ ไม่ต้องกรอกรหัสครั้งหน้า</span>
</label>
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

Avoid reintroducing extra buttons unless asked.
