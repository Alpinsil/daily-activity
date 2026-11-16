# Daily Activity Tracker

The GitHub Pages workflow publishes this app at the repository Pages root. The app uses the Google Apps Script deployment URL in `config.js`; that URL is public and does not contain the bot token.

## Apps Script

1. Copy `Code.gs` and `daily_activity_bridge.html` into the Apps Script project. The HTML file must be named `daily_activity_bridge`.
2. Set the Script Property `DAILY_ACTIVITY_ALLOWED_ORIGIN` to `https://alpinsil.github.io`.
3. Keep the existing `DAILY_ACTIVITY_BOT_TOKEN` Script Property for server-side bot requests.
4. Deploy a new Web App version that runs as the script owner and is accessible to web users.
5. In Apps Script and the bot configuration, use the same deployment URL. The bot token stays in server configuration and must never be added to `config.js`.

The bridge permits public read/write operations from the configured Pages origin without login. Anyone who can access the published app can manage the selected tracker spreadsheet. Use a dedicated spreadsheet with no unrelated data.

## GitHub Pages

The workflow in `.github/workflows/deploy-daily-activity.yml` publishes the static files. Enable GitHub Pages with **GitHub Actions** as the source, then push the repository to its `main` branch.

The Apps Script bridge is needed for hosted read/write operations. Without it, only the app shell loads; tracker data cannot synchronize.
