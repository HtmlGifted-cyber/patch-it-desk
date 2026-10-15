# Patch IT Desk

A free, bare-bones technical-issue desk that runs on a Google Sheet. Your team members or AI assistants
(Gemini, ChatGPT, Claude...) file a ticket when they hit a problem they can't fix themselves. Your IT
worker (a person or an AI agent) diagnoses it, hands the work to whoever can do it, and closes it out —
and everyone gets told.

- **Free:** Google Sheets + Apps Script + small Python helpers (standard library only).
- **Private:** the web app answers only with a secret key, and the key stays on your computer.
- **Simple for filers:** add one row and pick values from dropdowns. A "How to File" tab explains the rest.

## What it does

- **Tickets** with status, urgency, diagnosis and fix, and a **Desk Log** of every change.
- **Self-monitoring:** every hour `patchMonitor` reads Google's *"Summary of failures for Google Apps Script"*
  emails and files ONE ticket when the same script function fails 3+ times in 24 h (after that it only
  updates the count, never duplicates). Tip: set your important triggers to **Failure notification → Notify
  me immediately** so those emails arrive quickly.
- **Closed History** tab: every fixed/closed ticket is recorded. A new ticket about the same system (or from
  the same filer with a similar problem) gets *"POSSIBLE REPEAT of PT-0007 … earlier fix: …"* in its
  diagnosis. Reopens are counted.
- **Optional Notice Board** (a second Google Sheet, set `BOARD_ID`): fixing/closing a ticket posts a
  *RESOLVED* notice with the diagnosis and fix and asks the filer to confirm; *Assigned to Worker* posts a
  *WORK ORDER* with the steps; and you can post *WORK COMPLETED* notices for work done outside a ticket.
  The board's tabs (`Notices`, `Run Log`) are created for you.
- **Worker sweep write-back:** an AI worker that runs regular checks can log each sweep to the board's
  `Run Log` (one row per sweep ID, duplicates ignored) by opening a URL, or — if it can only send email —
  by emailing the desk account.
- **Replies:** the dashboard shows board rows that mention a ticket; *"PT-0002 done"* turns the card green.
- **Dashboard** (`dashboard/`): a local page to work the desk — filters, Start working / Send to worker /
  Waiting on owner / Mark fixed / Close (click twice to confirm), Reopen, a results bar, replies per ticket,
  and a *Report completed work* form.
- **Watcher** (`watcher/desk-watch.py`): desktop alert on new tickets; can hand each ticket to your own handler.

## Setup (about 15 minutes)

1. Create a Google Sheet called **IT Desk** and open **Extensions > Apps Script**.
2. Paste `apps-script/PatchDesk.gs` into `Code.gs`. At the top, set `TIMEZONE`, `WORKER_NAME`, `FILERS`,
   and (optional) `BOARD_ID` and `SWEEP_SENDERS`.
3. **Project Settings → Show "appsscript.json"**, then replace it with `apps-script/appsscript.json`
   (permissions: Sheets, Drive, Gmail, your email address, triggers). Edit its `timeZone` too.
4. Run **setupPatchDesk** once and allow the permissions. It builds the tabs and dropdowns, adds the hourly
   monitor trigger and creates a secret key (Project Settings → Script Properties → `PATCH_KEY`).
5. Optional board: run **authorizeBoard** (opens the board and confirms access). Optional sweeps: run
   **setupSweepKey** (logs a separate `SWEEP_KEY` that can only log sweeps).
6. **Deploy → New deployment → Web app**, Execute as *Me*, Who has access *Anyone*. Copy the `/exec` URL.
   After any later code or permission change: **Deploy → Manage deployments → Edit → New version**,
   or the web app keeps running the old code.
7. On your computer create `~/.config/it-desk/desk.env` and `chmod 600` it:
   ```
   DESK_URL=https://script.google.com/macros/s/XXXX/exec
   DESK_KEY=your-PATCH_KEY
   ```
8. Run `python3 dashboard/desk_helper.py` and open **http://127.0.0.1:8795/**. The helper serves the page,
   keeps the key, and listens on this computer only. Or run `watcher/desk-watch.py --loop 600` for alerts only.
9. Tell your assistants the rule: *"If you can't fix a technical problem in your own area, file it on the IT
   Desk following the How to File tab."*

### Worker sweeps

- **By URL** (no login needed; test it in an Incognito window, because a browser signed in to several Google
  accounts can show "Page Not Found"):
  `<web app url>?action=sweep&key=SWEEP_KEY&sweepId=SWEEP-20261015-1400&rowsChecked=120&changes=3&anomalies=0&actions=what%20you%20did&status=COMPLETED`
- **By email** (for agents that can't open URLs or write Sheets): send to the desk account (run
  `sweepInboxAddress` to see it) from an address listed in `SWEEP_SENDERS`, subject
  `[DESK SWEEP] SWEEP-20261015-1400`, body lines `Status:`, `Rows checked:`, `Changes detected:`,
  `Anomalies found:`, `Actions taken:`, `Details:`. The hourly monitor logs it and labels the email
  *IT Desk/Sweep logged*.

## Web app API

- `GET ?key=K&view=open|all` → `{ok, open, counts, tickets:[...]}` · `GET ?key=K&view=replies` → `{ok, replies:[...]}`
- `POST {"key":K,"action":"file", filedBy, department, system, problem, error, tried, urgency, handsOn}`
- `POST {"key":K,"action":"update", id, status, diagnosis, fix}`
- `POST {"key":K,"action":"notice", title, area, details, needsConfirm}`
- `POST {"key":K,"action":"backup", ...}` — saves a copy into the *IT Backups* folder before a change
- `GET|POST action=sweep` with `key=SWEEP_KEY` (see above)

## Notes

- Never commit `desk.env` or your keys. The sheet IDs you put in the script are your own; keep your copy private.
- Google only emails failure summaries to the account that owns each trigger.
