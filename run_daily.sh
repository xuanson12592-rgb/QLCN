#!/bin/bash
# Nap du lieu cong no hang ngay luc 6h30:
#   1) fetch_report.py : dang nhap inside.fpt.net (SSO) -> tai file DanhSachHoaDonChiTiet
#   2) auto_sync.py    : doi soat Loc AU -> day len Google Sheet "BC" -> Vercel
BASE_DIR="/home/sonx1/QLCN/QLCN"
LOG="$BASE_DIR/cron.log"
PY_SYNC="/home/sonx1/venv/bin/python"                     # pandas, gspread...
PY_FETCH="/home/sonx1/Giai Trinh TT/venv/bin/python"      # playwright

echo "==================== $(date '+%Y-%m-%d %H:%M:%S') ====================" >> "$LOG"

# Neu lan chay truoc con treo -> bo qua
if pgrep -f "fetch_report.py|auto_sync.py" > /dev/null; then
    echo "[SKIP] Tien trinh van dang chay tu lan truoc." >> "$LOG"
    exit 0
fi

cd "$BASE_DIR" || exit 1

# Ghi nhan nguon kich hoat (cron / web:<user> / remote-web) de web hien trang thai
TRIGGER="${QLCN_TRIGGER:-cron}"
timeout 30 "$PY_SYNC" sync_state.py running "$TRIGGER" >> "$LOG" 2>&1 || true

echo "----- BUOC 1: Tai bao cao tu inside.fpt.net -----" >> "$LOG"
timeout 600 "$PY_FETCH" fetch_report.py >> "$LOG" 2>&1
fetch_code=$?
if [ $fetch_code -ne 0 ]; then
    echo "[WARN] fetch_report.py exit=$fetch_code — chay auto_sync voi file moi nhat hien co." >> "$LOG"
fi

echo "----- BUOC 2: Doi soat + day len Google Sheets -----" >> "$LOG"
timeout 600 "$PY_SYNC" auto_sync.py >> "$LOG" 2>&1
code=$?
if [ $code -eq 0 ]; then
    timeout 30 "$PY_SYNC" sync_state.py done "$TRIGGER" "fetch=$fetch_code" >> "$LOG" 2>&1 || true
else
    timeout 30 "$PY_SYNC" sync_state.py error "$TRIGGER" "fetch=$fetch_code sync=$code" >> "$LOG" 2>&1 || true
fi
echo "[DONE] fetch=$fetch_code sync=$code" >> "$LOG"
exit $code
