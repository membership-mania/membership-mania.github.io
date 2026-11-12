# Member Mile

BBYO's involvement leaderboard: https://membership-mile.github.io

- `index.html` is the whole site.
- `data-YYYY-MM-DD.json` holds each day's numbers; `latest.json` points at today's and the previous day's file.
- `scripts/refresh.cjs` builds those files from the international Google Sheet and the US & Canada Game Card workbook.
- `.github/workflows/refresh.yml` runs the refresh every morning in GitHub's cloud. To run it now: Actions tab → Daily refresh → Run workflow.

## Where the numbers come from

- **International:** the "Interntional Reporting Summary FY27" Google Sheet, Summary tab. The sheet must stay viewable by anyone with the link.
- **US & Canada:** the SharePoint "Game Card For Home Page.xlsx" workbook. Put a direct download link to it in the `GAME_CARD_URL` secret (Settings → Secrets and variables → Actions). If the link stops working, the site keeps the last US & Canada numbers and shows "North America pending".
