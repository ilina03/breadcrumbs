# breadcrumbs

Breadcrumbs is pomodoro timer + habit tracker inspired by e-ink displays and the "breadcrumb method" used for ADHD.


When a focus session ends (or you stop early), you leave a breadcrumb: a quick note on where you stopped and what to do next. It shows up under the timer next time so it's easier to get going again.

Use it here: https://ilina03.github.io/breadcrumbs/ (add it to your home screen)

## Syncing with Google Sheets

Optional. Without it, your data stays on one device.

1. [Make a copy of the sheet](https://docs.google.com/spreadsheets/d/1a2IrBvzYHSFSxDh5Ri9xxLgXiRWR1ugxSEDvOIuqoos/copy)
2. In your copy: Extensions → Apps Script, paste in [Code.gs](Code.gs)
3. Change `SECRET` to any phrase you make up. It's basically a password between the app and your sheet.
4. Run `setup` and allow permissions (Advanced → Go to… is normal, it's your own script)
5. Deploy → New deployment → Web app. Execute as: Me, access: Anyone.
6. Copy the URL ending in `/exec`, then paste it and your secret into the app under sync \& backup.

Edit your habits in the `habits` tab. Every focus session and breadcrumb is logged in `sessions`.

## Hosting your own

Fork this repo and turn on GitHub Pages (Settings → Pages → main, root). Everything is in `index.html`; no build step.

