# -*- coding: utf-8 -*-
"""Theo doi canh bao tre hen tren FoxOne (fchat.fpt.com) va gui Zalo cho KTV nhan ca.

Quy trinh:
  1. Mo Chrome bang profile rieng (zalo_profile/) -> 2 tab: fchat.fpt.com (hoi thoai FoxOne) va chat.zalo.me.
     Lan dau chay: dang nhap thu cong ca 2 trang, profile se luu session cho cac lan sau.
  2. Moi POLL_SECONDS giay doc noi dung hoi thoai FoxOne, tim cac tin "CANH BAO ... TRE HEN".
  3. Tach tung dong "HD: <ma> — KTV: <ten>" theo nhom (Trien khai / Bao tri),
     tra SDT trong ktv.csv, gom HD theo KTV va gui 1 tin Zalo cho moi KTV.
  4. Tin da xu ly duoc luu vao sent_state.json de khong gui lai.

Chay:  python foxone_zalo.py            (theo doi lien tuc)
       python foxone_zalo.py --dry-run  (chi in ra, khong gui Zalo)
       python foxone_zalo.py --test-send 0905051098 "noi dung"   (thu gui 1 tin)
"""
import argparse
import csv
import hashlib
import json
import os
import random
import re
import sys
import time
import unicodedata
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "zalo_profile")
KTV_FILE = os.path.join(BASE_DIR, "ktv.csv")
STATE_FILE = os.path.join(BASE_DIR, "sent_state.json")
LOG_FILE = os.path.join(BASE_DIR, "foxone_zalo.log")

FCHAT_URL = os.environ.get("FCHAT_URL", "https://chat.fpt.com/")
FOXONE_NAME = os.environ.get("FOXONE_NAME", "FoxOne")
FOXONE_URL = os.environ.get("FOXONE_URL", "https://chat.fpt.com/chat/group/69264692836fcc1ddfbf62b8")
RELOAD_MINUTES = int(os.environ.get("RELOAD_MINUTES", "30"))
ZALO_URL = "https://chat.zalo.me/"
POLL_SECONDS = int(os.environ.get("POLL_SECONDS", "15"))
MAX_RETRY = int(os.environ.get("MAX_RETRY", "5"))

HEADER_RE = re.compile(r"C[ẢA]NH\s*B[ÁA]O[^\n]*TR[ỄE]\s*H[ẸE]N[^\n]*", re.I)
SECTION_RE = re.compile(r"^\s*(Tri[ểe]n\s*khai|B[ảa]o\s*tr[ìi])\b", re.I)
ITEM_RE = re.compile(r"H[ĐD]\s*:\s*([A-Z0-9]+)\s*[—–-]+\s*KTV\s*:\s*(.+?)\s*(?:\d{1,2}:\d{2})?\s*$", re.I)
DONG_VIEN = [
    "Anh em kỹ thuật luôn là điểm tựa vững chắc nhất của công ty. Cố gắng hết mình ca này nhé em! 💯",
    "Khách hàng đang chờ em, xử lý nhanh gọn ca này là mình về đích rồi. Cố lên em! 💪",
    "Mỗi ca đúng hẹn là một khách hàng hài lòng. Em làm được mà, cố gắng nhé! 🔥",
    "Chủ động liên hệ khách sớm giúp anh nhé, cả team luôn đứng sau em! 🤝",
    "Chặng nước rút rồi, dồn sức xử lý dứt điểm ca này nhé em. Cảm ơn em nhiều! 🙏",
    "Em là người giữ chữ tín cho FPT với khách hàng. Cố lên, đúng hẹn là chiến thắng! 🏆",
    "Vất vả cho em rồi, ưu tiên ca này giúp anh nhé. Có khó khăn gì báo anh hỗ trợ ngay! 💯",
    "Tinh thần thép, xử lý đúng hẹn! Cả team tin tưởng em! 🚀",
]
TIME_RE = re.compile(r"^\s*\d{1,2}:\d{2}\s*$")


def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def norm_name(s):
    s = unicodedata.normalize("NFC", s or "").strip().lower()
    return re.sub(r"\s+", " ", s)


def load_ktv():
    with open(KTV_FILE, encoding="utf-8-sig") as f:
        return {norm_name(r["ho_ten"]): r for r in csv.DictReader(f)}


def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"seen": {}, "initialized": False}


def save_state(state):
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1)
    os.replace(tmp, STATE_FILE)


def parse_alerts(text):
    """Tach text hoi thoai thanh danh sach canh bao.

    Moi canh bao: {"title", "items": [{"hd", "ktv", "loai"}], "raw"}.
    """
    text = unicodedata.normalize("NFC", text)
    lines = text.splitlines()
    alerts, cur, section = [], None, ""
    for line in lines:
        s = line.strip()
        if HEADER_RE.search(s) and s.upper().startswith("C"):
            if cur:
                alerts.append(cur)
            cur, section = {"title": HEADER_RE.search(s).group(0).strip(), "items": [], "raw": [s]}, ""
            continue
        if cur is None:
            continue
        if not s:
            continue
        m_sec = SECTION_RE.match(s)
        m_item = ITEM_RE.search(s.lstrip("•·-* ").strip())
        if m_sec and not m_item:
            section = m_sec.group(1).strip()
            cur["raw"].append(s)
        elif m_item:
            cur["items"].append({"hd": m_item.group(1).upper(), "ktv": m_item.group(2).strip(), "loai": section})
            cur["raw"].append(s)
            if TIME_RE.search(s.split()[-1]) or re.search(r"\d{1,2}:\d{2}\s*$", s):
                alerts.append(cur)
                cur = None
        elif s.lower().startswith("tổng") or s.lower().startswith("tong"):
            cur["raw"].append(s)
        elif TIME_RE.match(s):
            alerts.append(cur)
            cur = None
        else:
            # gap tin khac -> ket thuc canh bao hien tai
            alerts.append(cur)
            cur = None
    if cur:
        alerts.append(cur)
    return [a for a in alerts if a["items"]]


def alert_keys(alerts):
    """Khoa duy nhat cho moi canh bao; canh bao trung noi dung duoc phan biet bang thu tu xuat hien."""
    counts, out = {}, []
    for a in alerts:
        h = hashlib.sha1("\n".join(a["raw"]).encode("utf-8")).hexdigest()
        counts[h] = counts.get(h, 0) + 1
        out.append(f"{h}#{counts[h]}")
    return out


def build_messages(alert, ktv_map):
    """Gom HD theo KTV -> {sdt: (ktv_row, noi_dung)}; tra ve them danh sach KTV khong co trong ktv.csv."""
    by_ktv, missing = {}, []
    for it in alert["items"]:
        row = ktv_map.get(norm_name(it["ktv"]))
        if not row:
            missing.append(it)
            continue
        by_ktv.setdefault(row["dien_thoai"], (row, []))[1].append(it)
    out = {}
    for phone, (row, items) in by_ktv.items():
        groups = {}
        for it in items:
            groups.setdefault(it["loai"] or "Khác", []).append(it["hd"])
        lines = [f"⚠️ {alert['title'].upper()}", f"Em có {len(items)} hợp đồng cần xử lý ngay:"]
        for loai, hds in groups.items():
            lines += ["", f"{loai} ({len(hds)} HĐ)"] + [f"• HĐ: {hd}" for hd in hds]
        lines += ["", random.choice(DONG_VIEN)]
        out[phone] = (row, "\n".join(lines))
    return out, missing


# ---------------------------------------------------------------- Playwright

CDP_PORT = int(os.environ.get("CDP_PORT", "9333"))
CHROME_PATHS = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
                "/usr/bin/google-chrome"]


def open_browser(p):
    """Ket noi Chrome rieng (profile zalo_profile, cong CDP_PORT); tu mo Chrome neu chua chay."""
    import subprocess
    url = f"http://127.0.0.1:{CDP_PORT}"
    try:
        return p.chromium.connect_over_cdp(url).contexts[0]
    except Exception:
        pass
    exe = next((c for c in CHROME_PATHS if os.path.exists(c)), None)
    if not exe:
        raise RuntimeError("Khong tim thay chrome.exe")
    subprocess.Popen([exe, f"--remote-debugging-port={CDP_PORT}", f"--user-data-dir={PROFILE_DIR}",
                      "--no-first-run", "--no-default-browser-check", "--disable-background-timer-throttling",
                      "--disable-renderer-backgrounding", "--disable-backgrounding-occluded-windows",
                      FOXONE_URL, ZALO_URL])
    for _ in range(30):
        time.sleep(1)
        try:
            return p.chromium.connect_over_cdp(url).contexts[0]
        except Exception:
            pass
    raise RuntimeError("Khong ket noi duoc Chrome qua CDP")


def get_page(ctx, url_part, url):
    for pg in ctx.pages:
        if url_part in (pg.url or ""):
            return pg
    pg = ctx.pages[0] if (ctx.pages and ctx.pages[0].url in ("about:blank", "")) else ctx.new_page()
    pg.goto(url, wait_until="domcontentloaded")
    return pg


def ensure_foxone(page):
    """Dam bao khung chat dang mo la FoxOne; neu khong thi mo lai."""
    def current():
        try:
            return (page.locator(".block-name").first.inner_text(timeout=3000) or "").strip()
        except Exception:
            return ""
    if FOXONE_NAME in current():
        return
    try:
        page.get_by_text(FOXONE_NAME, exact=True).first.click(timeout=5000)
    except Exception:
        page.goto(FOXONE_URL, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    if FOXONE_NAME not in current():
        raise RuntimeError("Khong mo duoc hoi thoai FoxOne (kiem tra dang nhap chat.fpt.com)")


def read_foxone_text(page):
    """Lay text khung chat dang mo. Uu tien phan tu cuon chua nhieu text nhat ben phai."""
    return page.evaluate(r"""() => {
        const sc = document.querySelector('#messages-scroller');
        if (sc) { sc.scrollTop = sc.scrollHeight; return sc.innerText; }
        const els = Array.from(document.querySelectorAll('div,section,main'));
        let best = null, bestLen = 0;
        for (const el of els) {
            const st = getComputedStyle(el);
            if (!/(auto|scroll)/.test(st.overflowY)) continue;
            const r = el.getBoundingClientRect();
            if (r.width < 300 || r.height < 300) continue;
            const t = el.innerText || '';
            if (/TR[ỄE]\s*H[ẸE]N/i.test(t) && t.length > bestLen) { best = el; bestLen = t.length; }
        }
        return best ? best.innerText : document.body.innerText;
    }""")


def zalo_send(page, phone, message):
    page.bring_to_front()
    search = page.locator("#contact-search-input")
    search.click()
    search.fill("")
    search.type(phone, delay=40)
    # Muc "Tim ban qua so dien thoai" hien SDT -> bam vao de mo khung chat
    result = page.get_by_text(phone, exact=True).last
    try:
        result.wait_for(state="visible", timeout=8000)
    except Exception:
        raise RuntimeError(f"Khong tim thay tai khoan Zalo cua SDT {phone}")
    result.click()
    box = page.locator("#richInput")
    try:
        box.wait_for(state="visible", timeout=4000)
    except Exception:
        # Nguoi la: Zalo hien popup ho so -> bam "Nhan tin"
        btn = page.get_by_text("Nhắn tin", exact=True)
        if btn.count():
            btn.first.click()
        try:
            box.wait_for(state="visible", timeout=6000)
        except Exception:
            page.keyboard.press("Escape")
            raise RuntimeError(f"Khong mo duoc khung chat Zalo voi {phone} (SDT cua chinh tai khoan gui?)")
    page.wait_for_timeout(800)
    box.click()
    # Dan ca tin nhieu dong bang su kien paste (Shift+Enter tren Zalo Web lam mat cac dong sau)
    page.evaluate("""(t) => { const el = document.querySelector('#richInput'); el.focus();
        const dt = new DataTransfer(); dt.setData('text/plain', t);
        el.dispatchEvent(new ClipboardEvent('paste', {clipboardData: dt, bubbles: true, cancelable: true})); }""", message)
    page.wait_for_timeout(300)
    if not (box.inner_text() or "").strip():
        page.keyboard.insert_text(message.replace("\n", " "))
    page.wait_for_timeout(400)
    page.keyboard.press("Enter")
    page.wait_for_timeout(1500)
    if (box.inner_text() or "").strip():
        raise RuntimeError("Tin nhan chua duoc gui (o nhap van con noi dung)")


def run(dry_run=False):
    from playwright.sync_api import sync_playwright
    ktv_map, ktv_mtime = load_ktv(), os.path.getmtime(KTV_FILE)
    state = load_state()
    with sync_playwright() as p:
        ctx = open_browser(p)
        fchat = get_page(ctx, "chat.fpt.com", FCHAT_URL)
        zalo = None if dry_run else get_page(ctx, "zalo.me", ZALO_URL)
        log(f"Da mo fchat + Zalo. Neu chua dang nhap, hay dang nhap trong cua so Chrome. Kiem tra moi {POLL_SECONDS}s.")
        fchat.bring_to_front()
        last_reload = time.time()
        dry_seen = set()
        while True:
            try:
                if time.time() - last_reload > RELOAD_MINUTES * 60:
                    fchat.reload(wait_until="domcontentloaded")
                    fchat.wait_for_timeout(5000)
                    last_reload = time.time()
                if os.path.getmtime(KTV_FILE) != ktv_mtime:
                    ktv_map, ktv_mtime = load_ktv(), os.path.getmtime(KTV_FILE)
                    log(f"Da nap lai ktv.csv ({len(ktv_map)} KTV)")
                ensure_foxone(fchat)
                alerts = parse_alerts(read_foxone_text(fchat))
                if not state.get("initialized"):
                    for k in alert_keys(alerts):
                        state["seen"][k] = datetime.now().isoformat(timespec="seconds")
                    state["initialized"] = True
                    save_state(state)
                    log(f"Lan dau: danh dau {len(alerts)} canh bao cu la da xem (khong gui lai).")
                for a, k in zip(alerts, alert_keys(alerts)):
                    if k in state["seen"] or k in dry_seen:
                        continue
                    log(f"Canh bao moi: {a['title']} — {len(a['items'])} HD")
                    msgs, missing = build_messages(a, ktv_map)
                    for it in missing:
                        log(f"  [BO QUA] KTV '{it['ktv']}' (HD {it['hd']}) khong co trong ktv.csv")
                    ok = True
                    sent = state.setdefault("sent", {})
                    fails = state.setdefault("fails", {})
                    for phone, (row, text) in msgs.items():
                        if dry_run:
                            log(f"  [DRY] -> {row['ho_ten']} ({phone}):\n{text}")
                            continue
                        rk = f"{k}|{phone}"
                        if rk in sent:
                            continue
                        try:
                            zalo_send(zalo, phone, text)
                            sent[rk] = datetime.now().isoformat(timespec="seconds")
                            log(f"  [OK] Da gui {row['ho_ten']} ({phone})")
                        except Exception as e:
                            fails[rk] = fails.get(rk, 0) + 1
                            log(f"  [LOI] Gui {row['ho_ten']} ({phone}) that bai lan {fails[rk]}: {e}")
                            if fails[rk] < MAX_RETRY:
                                ok = False
                            else:
                                log(f"  [BO QUA] Da thu {MAX_RETRY} lan, khong gui lai cho {row['ho_ten']}")
                        save_state(state)
                    if dry_run:
                        dry_seen.add(k)
                    elif ok:
                        state["seen"][k] = datetime.now().isoformat(timespec="seconds")
                        save_state(state)
                    fchat.bring_to_front()
            except Exception as e:
                log(f"[LOI] {e}")
            time.sleep(POLL_SECONDS)


def test_send(phone, message):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        ctx = open_browser(p)
        zalo = get_page(ctx, "zalo.me", ZALO_URL)
        input("Dang nhap Zalo Web (neu can) roi bam Enter de gui thu...")
        zalo_send(zalo, phone, message)
        log(f"Da gui thu toi {phone}")
        input("Bam Enter de dong...")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--test-send", nargs=2, metavar=("SDT", "NOI_DUNG"))
    ap.add_argument("--parse", metavar="FILE", help="Thu tach canh bao tu file text")
    a = ap.parse_args()
    if a.parse:
        km = load_ktv()
        for al in parse_alerts(open(a.parse, encoding="utf-8-sig").read()):
            msgs, missing = build_messages(al, km)
            print("==", al["title"], [i["hd"] for i in al["items"]])
            for ph, (r, t) in msgs.items():
                print(f"-> {ph}\n{t}\n")
            for it in missing:
                print("  BO QUA:", it)
    elif a.test_send:
        test_send(*a.test_send)
    else:
        run(dry_run=a.dry_run)
