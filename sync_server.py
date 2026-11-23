#!/usr/bin/env python3
"""Mini HTTP trigger cho QLCN: nut tren web chay run_daily.sh bang tay.

Chi lang nghe tren 127.0.0.1:8765 — duoc goi tu trinh duyet mo TREN MAY NAY.
Co xu ly preflight OPTIONS + header Access-Control-Allow-Private-Network
de trang https (congno.click) goi duoc vao localhost (Chrome Private Network
Access). Chay nen qua cron @reboot.

Endpoints:
  POST /run-sync?who=<user>   -> chay run_daily.sh (bo qua neu dang chay)
  GET  /sync-status           -> trang thai lan chay + running + log tail
  GET  /healthz               -> kiem tra song
"""

import json
import os
import subprocess
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUN_SCRIPT = os.path.join(BASE_DIR, "run_daily.sh")
STATUS_FILE = os.path.join(BASE_DIR, "sync_status.json")
LOG_FILE = os.path.join(BASE_DIR, "cron.log")
PORT = 8765
RUNNING_PATTERN = "fetch_report.py|auto_sync.py"


def is_running():
    return (
        subprocess.run(
            ["pgrep", "-f", RUNNING_PATTERN], capture_output=True
        ).returncode
        == 0
    )


def load_status():
    try:
        with open(STATUS_FILE, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def log_tail(lines=8):
    try:
        with open(LOG_FILE, encoding="utf-8", errors="replace") as fh:
            return fh.read().splitlines()[-lines:]
    except Exception:
        return []


def start_run(who):
    env = dict(os.environ)
    env["QLCN_TRIGGER"] = f"web:{who}" if who else "web"
    log = open(LOG_FILE, "ab", buffering=0)
    subprocess.Popen(
        ["bash", RUN_SCRIPT],
        cwd=BASE_DIR,
        env=env,
        stdout=log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    )


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Private-Network", "true")

    def _send(self, code, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._cors()
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Access-Control-Max-Age", "600")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/sync-status":
            status = load_status()
            running = is_running()
            state = status.get("state")
            if state == "running" and not running:
                state = "stopped"  # treo giua chung (may tat/mat mang...)
            self._send(200, {
                "ok": True,
                "running": running,
                "state": state,
                "trigger": status.get("trigger"),
                "started_at": status.get("started_at"),
                "finished_at": status.get("finished_at"),
                "message": status.get("message"),
                "log_tail": log_tail(),
            })
            return
        if path == "/healthz":
            self._send(200, {"ok": True})
            return
        self._send(404, {"ok": False, "error": "not found"})

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path != "/run-sync":
            self._send(404, {"ok": False, "error": "not found"})
            return
        who = urllib.parse.parse_qs(parsed.query).get("who", [""])[0][:40]
        if is_running():
            self._send(200, {"ok": True, "status": "already_running"})
            return
        try:
            start_run(who)
        except Exception as exc:
            self._send(500, {"ok": False, "error": str(exc)})
            return
        self._send(200, {"ok": True, "status": "started"})

    def log_message(self, fmt, *args):
        return


if __name__ == "__main__":
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"sync_server dang lang nghe http://127.0.0.1:{PORT}", flush=True)
    server.serve_forever()
