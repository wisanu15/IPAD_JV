# iPad JV GitHub Pages wrapper

This folder is a small installable PWA wrapper for Android Chrome. It hosts the real manifest and service worker on GitHub Pages, then opens the Google Apps Script web app inside an iframe.

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
6. Open the GitHub Pages URL on Android Chrome and use Install app/Add to Home screen.

The Apps Script web app must keep `setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)` in `doGet`, otherwise the iframe will be blocked.
