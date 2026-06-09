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
2. Choose Add to Home screen or Install app.
3. Open iPad JV and tap Open with Chrome.

iPhone/iPad:

1. Open the GitHub Pages URL in Safari.
2. Tap Share.
3. Choose Add to Home Screen.
4. Open iPad JV and tap Open in Safari.

If Apps Script shows a Google Drive error inside an in-app browser, copy the Apps Script `/exec` URL and paste it into Safari or Chrome directly.
