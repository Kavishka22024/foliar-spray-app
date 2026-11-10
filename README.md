# Installable phone app (works without signal)

This folder is the same app as the Google web page, packaged so a phone can install it and open it with NO signal.

## One-time steps
1. In Apps Script: **Deploy > Manage deployments > edit > Version: New version > Deploy** (the new version contains the endpoint this app talks to).
2. Open `config.js` and paste your web app URL (the one ending in `/exec`).
3. Upload this whole folder to an HTTPS static host. Free options: GitHub Pages, Netlify (drag and drop the folder), Cloudflare Pages.
4. On each phone: open the hosted address in Chrome (Android) or Safari (iPhone), log in once WITH signal, then **Add to Home screen**. From then on it opens like an app, even without signal.

## How it behaves in the field
- Everything typed is saved on the phone straight away and survives closing the app.
- Pressing Submit with no signal puts the entries in a queue on the phone; they upload by themselves when signal returns.
- The first login each day (or after 6 hours) needs signal. Without signal the phone keeps working from what it remembers.
- After you change the app, run `node tools/build-pwa.js` again and upload the folder again. Phones update themselves the next time they open the app with signal.
