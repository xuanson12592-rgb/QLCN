#!/usr/bin/env python3
"""Ghi trang thai chay nap du lieu vao sync_status.json + tab 'SyncStatus'
tren Google Sheet (de trang web doc qua GViz va hien thi cho nguoi dung).

Usage: sync_state.py <running|done|error> [trigger] [message]

Duoc goi tu run_daily.sh. Khong bao gio raise — loi ghi sheet chi in ra,
khong lam chet tien trinh chinh.
"""

import json
import os
import socket
import sys
from datetime import datetime, timezone

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATUS_FILE = os.path.join(BASE_DIR, "sync_status.json")
STATUS_SHEET = "SyncStatus"
socket.setdefaulttimeout(20)


def now_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def main():
    state = sys.argv[1] if len(sys.argv) > 1 else "running"
    trigger = sys.argv[2] if len(sys.argv) > 2 else "cron"
    message = sys.argv[3] if len(sys.argv) > 3 else ""

    status = {}
    if state != "running":
        try:
            with open(STATUS_FILE, encoding="utf-8") as fh:
                status = json.load(fh)
        except Exception:
            status = {}
        trigger = status.get("trigger") or trigger

    if state == "running":
        status = {
            "state": "running",
            "trigger": trigger,
            "started_at": now_iso(),
            "finished_at": "",
            "message": "",
        }
    else:
        status["state"] = state
        status["trigger"] = trigger
        status["finished_at"] = now_iso()
        status["message"] = message

    with open(STATUS_FILE, "w", encoding="utf-8") as fh:
        json.dump(status, fh, ensure_ascii=False)

    cell = json.dumps({
        "state": status["state"],
        "time": status.get("finished_at") or status.get("started_at"),
        "trigger": status.get("trigger"),
        "msg": message,
    }, ensure_ascii=False)

    try:
        import gspread
        from dotenv import load_dotenv

        load_dotenv(os.path.join(BASE_DIR, ".env"))
        sheet_id = os.getenv("GOOGLE_SHEET_ID", "").strip()
        client = gspread.service_account(
            filename=os.path.join(BASE_DIR, "credentials.json")
        )
        spreadsheet = client.open_by_key(sheet_id)
        try:
            ws = spreadsheet.worksheet(STATUS_SHEET)
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(title=STATUS_SHEET, rows=5, cols=2)
        ws.update(values=[[cell]], range_name="A1")
    except Exception as exc:
        print(f"sync_state: khong ghi duoc SyncStatus len sheet: {exc}")


if __name__ == "__main__":
    main()
