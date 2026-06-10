# iPad JV GitHub Pages wrapper

This folder is a small mobile launcher for the Google Apps Script web app. It is hosted on GitHub Pages and gives Android, iPhone, and iPad users a stable Home Screen entry point.

## Setup

1. Deploy the Apps Script web app as a new version.
2. Copy the `/exec` URL.
3. Paste it into `docs/config.js`:

```js
window.IPAD_JV_CONFIG = {
  GAS_WEB_APP_URL: 'https://script.google.com/macros/s/DEPLOYMENT_ID/exec'
};
```

4. Push this repository to GitHub.
5. In GitHub repository settings, enable Pages from the `docs` folder.
6. Open the GitHub Pages URL on the phone or tablet.

## Install

Android:

1. Open the GitHub Pages URL in Chrome.
2. Tap Test Android fullscreen.
3. If the app loads correctly, go back to the launcher page.
4. Choose Add to Home screen or Install app.
5. Open iPad JV from the Home Screen. It should open fullscreen.

iPhone/iPad:

1. Remove the old iPad JV Home Screen icon if one already exists.
2. Open the GitHub Pages URL in Safari.
3. Tap Share.
4. Choose Add to Home Screen.
5. Open iPad JV from the Home Screen. It should open fullscreen and load the app inside the launcher.

If Apps Script shows a Google Drive error inside an in-app browser, copy the Apps Script `/exec` URL and paste it into Safari or Chrome directly.

## Admin push notifications

This wrapper supports OneSignal Web Push for admin issue alerts.

1. Create a OneSignal Web Push app for `https://wisanu15.github.io`.
2. In OneSignal Web Settings, enable custom service worker settings:
   - Path to service worker files: `/IPAD_JV/push/onesignal/`
   - Service worker filename: `OneSignalSDKWorker.js`
   - Service worker registration scope: `/IPAD_JV/push/onesignal/`
3. Put the OneSignal Web Push App ID in `docs/config.js` as `ONESIGNAL_APP_ID`.
4. In Apps Script Project Settings > Script Properties, add:
   - `ONESIGNAL_APP_ID`
   - `ONESIGNAL_REST_API_KEY`
   - `IPAD_JV_PWA_URL` = `https://wisanu15.github.io/IPAD_JV/`
5. Deploy/publish both GitHub Pages and the Apps Script web app.
6. Open the installed app as an admin, tap the bell button once, and allow notifications.
