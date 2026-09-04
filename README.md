# LiveMonitor

Tech Monitor is migrating from the legacy Supabase-backed GitHub Pages build to a Google Sheets and Google Apps Script web app.

## Google Sheets version

The current migration source is in [`apps-script/`](apps-script/):

- [`Code.gs`](apps-script/Code.gs) — shared Google Sheets backend
- [`Index.html`](apps-script/Index.html) — dashboard frontend
- [`appsscript.json`](apps-script/appsscript.json) — Apps Script manifest
- [`README.md`](apps-script/README.md) — step-by-step deployment and update instructions

The Google Sheets build uses row-specific writes, server-generated UTC timestamps, a script lock, and technician version checks so simultaneous team actions do not overwrite each other.

The root `index.html` remains the legacy deployment until the Apps Script `/exec` URL is created and verified. After verification, the old Supabase build can be retired without interrupting the team.
