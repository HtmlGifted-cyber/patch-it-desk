# Patch IT Desk

A bare-bones technical-issue desk that runs on a Google Sheet. Your team members or AI assistants (Gemini,
ChatGPT, Claude...) file a ticket when they hit a problem they can't fix themselves. A small watcher on your
computer notices each new ticket and alerts you, and can hand it to an IT worker (a person or an AI agent).

- **Free:** Google Sheets + Apps Script + a ~60-line Python watcher.
- **Private:** the web app answers only with a secret key, and the key stays on your machine.
- **Simple for filers:** add one row and pick values from dropdowns. A "How to File" tab explains the rest.

## How it works
1. Someone hits a problem and tries to fix it in their own area first.
2. If they can't fix it, they add a row on the **Tickets** tab (Status = New).
3. The watcher checks the web app every 10 minutes. On a new ticket it shows a desktop alert and can run your
   handler (`ON_TICKET`) with the ticket as JSON.
4. Your IT worker writes the diagnosis and status back through the web app (`doPost`, action `update`).
   Every action is recorded on the **Desk Log** tab.

## Setup (about 10 minutes)
1. Create a Google Sheet called "IT Desk". In it, open **Extensions > Apps Script**.
2. Paste `apps-script/PatchDesk.gs`. Edit `TIMEZONE`, `OWNER` and `FILERS` at the top, then save.
3. Run **setupPatchDesk** once and allow the permissions. It builds the tabs, dropdowns and rules, and creates
   a secret key. You can always find the key later under **Project Settings > Script Properties > PATCH_KEY**.
4. **Deploy > New deployment > Web app**, with Execute as: *Me* and Who has access: *Anyone*. Without the key
   the web app only answers `bad key`. Copy the `/exec` URL.
5. On your computer, create `~/.config/it-desk/desk.env` (run `chmod 600` on it):
   ```
   DESK_URL=https://script.google.com/macros/s/XXXX/exec
   DESK_KEY=your-key
   ```
6. Run `watcher/desk-watch.py` to test it, then schedule it with cron, a systemd timer or `--loop 600`.
7. Tell your assistants the sheet's file ID and the rule: *"If you can't fix a technical problem in your own
   area, file it on the IT Desk following the How to File tab."*

## Web app API
- `GET  ?key=K&view=open|all` returns `{ok, open, counts, tickets:[...]}`
- `POST {"key":K,"action":"file", filedBy, department, system, problem, error, tried, urgency, handsOn}`
- `POST {"key":K,"action":"update", id, status, diagnosis, fix}`

After you change the script, publish the new code with **Deploy > Manage deployments > Edit > New version**.
Otherwise the web app keeps running the old code.

## Status
A bare-bones first public edition, taken from a working private setup. Next ideas: a hosted watcher, a
monitor that files tickets by itself from failing Apps Script triggers, and an AI triage handler example.
