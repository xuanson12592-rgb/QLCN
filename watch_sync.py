#!/usr/bin/env python3
"""Cron chay moi phut: doc tab 'SyncTrigger' tren Google Sheet.

Nut "Nap du lieu ngay" tren thiet bi khac ghi mot timestamp (ISO UTC) vao
SyncTrigger!A1 thong qua Google Apps Script. Khi thay timestamp moi hon lan
da xu ly -> chay run_daily.sh (bo qua neu co tien trinh dang chay hoac lan
chay gan nhat da ket thuc SAU moc trigger).
"""

import json
import os
import socket
import subprocess
from datetime import datetime, timezone

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRIGGER_SHEET = "SyncTrigger"
LAST_FILE = os.path.join(BASE_DIR, "sync_trigger_last.txt")
STATUS_FILE = os.path.join(BASE_DIR, "sync_status.json")
LOG_FILE = os.path.join(BASE_DIR, "cron.log")
RUN_SCRIPT = os.path.join(BASE_DIR, "run_daily.sh")
socket.setdefaulttimeout(20)


def parse_iso(value):
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(str(value).strip().replace("Z", "+00:00"))
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        return None


def is_running():
    return (
        subprocess.run(
            ["pgrep", "-f", "fetch_report.py|auto_sync.py"], capture_output=True
        ).returncode
        == 0
    )


def last_processed():
    try:
        with open(LAST_FILE, encoding="utf-8") as fh:
            return fh.read().strip()
    except Exception:
        return ""


def mark_processed(trigger):
    with open(LAST_FILE, "w", encoding="utf-8") as fh:
        fh.write(trigger)


def read_trigger():
    import gspread
    from dotenv import load_dotenv

    load_dotenv(os.path.join(BASE_DIR, ".env"))
    sheet_id = os.getenv("GOOGLE_SHEET_ID", "").strip()
    client = gspread.service_account(
        filename=os.path.join(BASE_DIR, "credentials.json")
    )
    spreadsheet = client.open_by_key(sheet_id)
    try:
        ws = spreadsheet.worksheet(TRIGGER_SHEET)
    except gspread.WorksheetNotFound:
        return None
    return (ws.acell("A1").value or "").strip()


def main():
    trigger = read_trigger()
    if not trigger or trigger == last_processed():
        return

    # Co tien trinh dang chay -> lan chay do da dap ung yeu cau
    if is_running():
        mark_processed(trigger)
        return

    # Lan chay gan nhat ket thuc sau moc trigger -> da duoc xu ly (vd: nut
    # bam tren may nay chay qua sync_server, khong can chay lai)
    try:
        with open(STATUS_FILE, encoding="utf-8") as fh:
            status = json.load(fh)
    except Exception:
        status = {}
    t_dt = parse_iso(trigger)
    f_dt = parse_iso(status.get("finished_at"))
    if t_dt and f_dt and f_dt >= t_dt:
        mark_processed(trigger)
        return

    mark_processed(trigger)
    env = dict(os.environ)
    env["QLCN_TRIGGER"] = "remote-web"
    log = open(LOG_FILE, "ab", buffering=0)
    subprocess.Popen(
        ["bash", RUN_SCRIPT],
        cwd=BASE_DIR,
        env=env,
        stdout=log,
        stderr=subprocess.STDOUT,
        start_new_session=True,
    )
    print(f"watch_sync: nhan trigger {trigger} -> chay run_daily.sh")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"watch_sync loi: {exc}")
