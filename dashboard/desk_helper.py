#!/usr/bin/env python3
"""desk_helper.py - standalone Python 3.8 stdlib server for IT Desk dashboard.
Listening on 127.0.0.1:8795.
GET /patch-desk
POST /desk/update
POST /desk/notice
"""
import calendar, json, os, re, subprocess, threading, time, urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HOME = os.path.expanduser("~")
ALLOWED_ORIGINS = {"http://127.0.0.1", "http://localhost", "http://127.0.0.1:8795", "http://localhost:8795"}

DESK_ENV = HOME + "/.config/it-desk/desk.env"
DESK_SEEN = HOME + "/.config/it-desk/seen-tickets.json"
_desk = {"ok": False, "message": "not checked yet", "checked": 0, "tickets": [], "open": 0}


def parse_dt(s):
    if not s:
        return 0
    try:
        if s.endswith("Z"):
            s = s[:-1] + "+00:00"
        return calendar.timegm(time.strptime(s.split(".")[0].split("+")[0], "%Y-%m-%dT%H:%M:%S"))
    except Exception:
        try:
            return float(s)
        except Exception:
            return 0


def desk_check():
    global _desk
    try:
        cfg = dict(l.strip().split("=", 1) for l in open(DESK_ENV) if "=" in l and not l.startswith("#"))
        url = cfg["DESK_URL"] + "?view=all&key=" + cfg["DESK_KEY"]
        with urllib.request.urlopen(url, timeout=40) as r:
            d = json.load(r)
        if not d.get("ok"):
            raise ValueError(d.get("error", "desk refused"))
        keep = ("Ticket ID", "Filed", "Filed By", "Department", "System / Script", "Problem", "Exact Error",
                "Urgency", "Needs Owner Hands-On", "Status", "IT Diagnosis", "Fix / Next Step", "Updated", "Closed")
        raw_tickets = d.get("tickets", [])
        now_ts = time.time()
        one_day_ago = now_ts - 86400

        processed = []
        open_count = 0
        for t in raw_tickets:
            tk = {k: t.get(k, "") for k in keep}
            status = str(tk.get("Status", "New")).strip()
            is_done = status in ("Fixed", "Closed")
            updated_str = tk.get("Closed") or tk.get("Updated") or tk.get("Filed") or ""
            ts = parse_dt(updated_str)

            if not is_done:
                open_count += 1
                tk["recent_done"] = False
                processed.append(tk)
            else:
                if ts >= one_day_ago:
                    tk["recent_done"] = True
                    processed.append(tk)

        tickets = processed

        try:
            seen = set(json.load(open(DESK_SEEN)))
        except Exception:
            seen = set()

        not_done_tickets = [t for t in tickets if not t.get("recent_done") and str(t.get("Status","")) not in ("Fixed", "Closed")]
        new = [t for t in not_done_tickets if t["Ticket ID"] not in seen]
        for t in new:
            try:
                subprocess.run(["notify-send", "-u", "critical" if t["Urgency"] in ("High", "Down") else "normal",
                                "IT Desk: " + t["Ticket ID"],
                                f'{t["Filed By"]}: {t["System / Script"]} - {str(t["Problem"])[:140]}'],
                               timeout=10, env=dict(os.environ, DISPLAY=":0"))
            except Exception:
                pass
        if new:
            os.makedirs(os.path.dirname(DESK_SEEN), exist_ok=True)
            json.dump(sorted(seen | {t["Ticket ID"] for t in not_done_tickets}), open(DESK_SEEN, "w"))

        replies = []
        try:
            with urllib.request.urlopen(cfg["DESK_URL"] + "?view=replies&key=" + cfg["DESK_KEY"], timeout=60) as r:
                rd = json.load(r)
            if rd.get("ok"):
                replies = rd.get("replies", [])
        except Exception:
            replies = _desk.get("replies", [])

        _desk = {"ok": True, "checked": int(time.time()), "open": open_count, "tickets": tickets, "replies": replies}
    except Exception as e:
        _desk = dict(_desk, ok=False, message=str(e)[:160], checked=int(time.time()))


def desk_loop():
    while True:
        desk_check()
        time.sleep(600)


class H(BaseHTTPRequestHandler):
    def _origin_ok(self):
        o = self.headers.get("Origin")
        return o is None or o in ALLOWED_ORIGINS

    def _send(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        o = self.headers.get("Origin")
        if o in ALLOWED_ORIGINS:
            self.send_header("Access-Control-Allow-Origin", o)
            self.send_header("Access-Control-Allow-Headers", "X-Eco-Control, Content-Type")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self._send(204 if self._origin_ok() else 403, {})

    def do_GET(self):
        if self.path in ("/", "/itdesk.html"):   # serve the desk page itself: open http://127.0.0.1:8795/
            body = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "itdesk.html"), "rb").read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            return self.wfile.write(body)
        if self.path == "/patch-desk" and self._origin_ok():
            return self._send(200, _desk)
        self._send(404, {"error": "not found"})

    def do_POST(self):
        if self.path == "/desk/notice":
            if not (self._origin_ok() and self.headers.get("X-Eco-Control") == "1"):
                return self._send(403, {"error": "refused"})
            try:
                data = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0)) or 0) or b"{}")
            except Exception:
                return self._send(400, {"ok": False, "message": "invalid json"})
            title = str(data.get("title", "")).strip()
            if not title or len(title) > 200:
                return self._send(400, {"ok": False, "message": "title required (max 200 characters)"})
            payload = {"action": "notice", "title": title,
                       "details": str(data.get("details", ""))[:1500],
                       "area": str(data.get("area", ""))[:100],
                       "needsConfirm": str(data.get("needsConfirm", ""))[:100]}
            try:
                cfg = dict(l.strip().split("=", 1) for l in open(DESK_ENV) if "=" in l and not l.startswith("#"))
                payload["key"] = cfg["DESK_KEY"]
                req = urllib.request.Request(cfg["DESK_URL"], data=json.dumps(payload).encode(),
                                             headers={"Content-Type": "application/json"}, method="POST")
                with urllib.request.urlopen(req, timeout=40) as r:
                    resp = json.load(r)
                if not resp.get("ok"):
                    resp["message"] = resp.get("error") or "desk refused"
                return self._send(200, resp)
            except Exception as e:
                return self._send(500, {"ok": False, "message": str(e)[:200]})

        if self.path == "/desk/update":
            if not (self._origin_ok() and self.headers.get("X-Eco-Control") == "1"):
                return self._send(403, {"error": "refused"})
            try:
                length = int(self.headers.get("Content-Length", 0))
                body_data = self.rfile.read(length) if length else b"{}"
                data = json.loads(body_data.decode("utf-8"))
            except Exception:
                return self._send(400, {"ok": False, "message": "invalid json"})

            tid = str(data.get("id", ""))
            status = str(data.get("status", ""))
            diagnosis = data.get("diagnosis")
            fix = data.get("fix")

            if not re.match(r"^PT-\d{4,}$", tid):
                return self._send(400, {"ok": False, "message": "invalid id format"})
            if status not in ("New", "Diagnosing", "Assigned to Worker", "Waiting on Owner", "Fixed", "Closed"):
                return self._send(400, {"ok": False, "message": "invalid status"})

            if diagnosis is not None and len(str(diagnosis)) > 2000:
                return self._send(400, {"ok": False, "message": "diagnosis too long"})
            if fix is not None and len(str(fix)) > 2000:
                return self._send(400, {"ok": False, "message": "fix too long"})

            try:
                cfg = dict(l.strip().split("=", 1) for l in open(DESK_ENV) if "=" in l and not l.startswith("#"))
                payload = {"key": cfg["DESK_KEY"], "action": "update", "id": tid, "status": status}
                if diagnosis is not None and str(diagnosis).strip() != "":
                    payload["diagnosis"] = str(diagnosis)
                if fix is not None and str(fix).strip() != "":
                    payload["fix"] = str(fix)

                req_data = json.dumps(payload).encode("utf-8")
                req = urllib.request.Request(cfg["DESK_URL"], data=req_data, headers={"Content-Type": "application/json"}, method="POST")
                with urllib.request.urlopen(req, timeout=40) as r:
                    resp_json = json.load(r)

                threading.Thread(target=desk_check, daemon=True).start()

                if isinstance(resp_json, dict) and resp_json.get("ok") and "message" not in resp_json:
                    resp_json["message"] = "Saved" + (" · " + resp_json["board"] if resp_json.get("board") not in (None, "none") else "")
                return self._send(200, resp_json if isinstance(resp_json, dict) else {"ok": True, "message": str(resp_json)})
            except Exception as e:
                return self._send(500, {"ok": False, "message": str(e)[:200]})

        self._send(404, {"error": "not found"})

    def log_message(self, *a):
        pass


if __name__ == "__main__":
    threading.Thread(target=desk_loop, daemon=True).start()
    ThreadingHTTPServer(("127.0.0.1", 8795), H).serve_forever()
