#!/usr/bin/env python3
"""desk-watch.py - polls your IT Desk web app and alerts on new tickets.

Config file (chmod 600):  ~/.config/it-desk/desk.env
    DESK_URL=https://script.google.com/macros/s/.../exec
    DESK_KEY=<the key printed by setupPatchDesk>
Optional:
    ON_TICKET=/path/to/handler   # run for each new ticket, ticket JSON on stdin

Usage:  desk-watch.py            # check once (good for cron / systemd timer)
        desk-watch.py --loop 600 # check every 600 seconds
The key is only ever sent to your own web app URL; it is never printed."""
import json, os, subprocess, sys, time, urllib.parse, urllib.request

CFG = os.path.expanduser("~/.config/it-desk/desk.env")
SEEN = os.path.expanduser("~/.config/it-desk/seen.json")


def load_cfg():
    return dict(l.strip().split("=", 1) for l in open(CFG) if "=" in l and not l.startswith("#"))


def check(cfg):
    q = urllib.parse.urlencode({"view": "open", "key": cfg["DESK_KEY"]})
    with urllib.request.urlopen(cfg["DESK_URL"] + "?" + q, timeout=40) as r:
        d = json.load(r)
    if not d.get("ok"):
        raise SystemExit("desk refused: " + str(d.get("error")))
    try:
        seen = set(json.load(open(SEEN)))
    except Exception:
        seen = set()
    new = [t for t in d["tickets"] if t["Ticket ID"] not in seen]
    for t in new:
        title = f'IT Desk: {t["Ticket ID"]} ({t.get("Urgency", "Normal")})'
        body = f'{t.get("Filed By", "")}: {t.get("System / Script", "")} - {str(t.get("Problem", ""))[:140]}'
        print(title, "|", body)
        try:
            subprocess.run(["notify-send", title, body], timeout=10)
        except Exception:
            pass
        if cfg.get("ON_TICKET"):
            subprocess.Popen([cfg["ON_TICKET"]], stdin=subprocess.PIPE, text=True,
                             start_new_session=True).communicate(json.dumps(t))
    if new:
        os.makedirs(os.path.dirname(SEEN), exist_ok=True)
        json.dump(sorted(seen | {t["Ticket ID"] for t in d["tickets"]}), open(SEEN, "w"))
    print(f'{d["open"]} open ticket(s), {len(new)} new')


if __name__ == "__main__":
    cfg = load_cfg()
    if len(sys.argv) > 2 and sys.argv[1] == "--loop":
        while True:
            try:
                check(cfg)
            except Exception as e:
                print("check failed:", e)
            time.sleep(int(sys.argv[2]))
    else:
        check(cfg)
