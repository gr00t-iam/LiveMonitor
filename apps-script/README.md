# Tech Monitor — Google Sheets / Apps Script Deployment

This version replaces Supabase with the **Live Dashboard** Google Sheet and a bound Google Apps Script web app.

## Files

- `Code.gs` — locked, row-specific database operations.
- `Index.html` — shared dashboard interface.
- `appsscript.json` — Central Time and spreadsheet authorization settings.

## One-time deployment

1. Open the [Live Dashboard spreadsheet](https://docs.google.com/spreadsheets/d/1Snz0D4WBgKX9P2PT8M6lyCzPw9F3LDk6ao1l30PX6xQ/edit).
2. Select **Extensions → Apps Script**.
3. In the Apps Script editor, open `Code.gs` and replace its contents with this repository's `apps-script/Code.gs`.
4. Select **Add a file (+) → HTML**, name it exactly `Index`, and paste this repository's `apps-script/Index.html` into it.
5. Open **Project Settings** and enable **Show "appsscript.json" manifest file in editor**.
6. Replace the manifest contents with this repository's `apps-script/appsscript.json`.
7. Click **Save project**.
8. From the function selector, choose `getInitialState`, click **Run**, and approve the requested Google Sheets access. A successful run returns without an exception.
9. Select **Deploy → New deployment**.
10. Choose **Web app** as the deployment type.
11. Set **Execute as** to **Me**.
12. For **Who has access**, select the narrowest option that every team member can use. Prefer your organization/domain; otherwise use **Anyone** if company policy permits it.
13. Click **Deploy**, then copy the `/exec` web-app URL. This is the team's live dashboard URL.

Do not use the `/dev` test URL for the team. It only works for project editors and always runs the most recently saved code.

## Updating the live app

1. Paste updated source files into the bound Apps Script project.
2. Select **Deploy → Manage deployments**.
3. Edit the existing web-app deployment.
4. Choose **New version**, add a short description, and deploy.
5. Keep using the same `/exec` URL.

## Shared timer behavior

- The Sheet stores UTC shift starts and update deadlines.
- Each open browser calculates its countdown locally from the shared deadline.
- The dashboard checks the Sheet every 10 seconds while visible and every 30 seconds while hidden.
- Writes are protected with `LockService.getScriptLock()`.
- Each technician has a version number. A stale screen receives a conflict response instead of overwriting a newer action.
- Only the affected technician row is changed. The entire dashboard is never overwritten.

## Database tabs

- **Technicians** — current shared technician and timer state.
- **Activity Log** — append-only event history.
- **Settings** — shared duration, alert, timezone, and polling configuration.
- **Command Staff** — names available in the Current Operator selector.

## Local visual preview

Serve `apps-script/Index.html` from a local web server and add `?preview=1` to the URL. Preview mode uses sample data and never writes to Google Sheets.

## Important operating rules

- Use the dashboard buttons for normal timer changes instead of manually editing timestamp columns.
- Staff may edit names and settings in the Sheet when needed, but do not rename the four tabs or their header columns.
- Use `America/Chicago`, not a fixed `CST` offset, so daylight-saving changes remain correct.
