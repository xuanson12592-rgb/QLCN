"""Tao dashboard 16 KTV (lich truc, ton ca, KPI) thanh 1 file HTML.

Nguon du lieu:
- ho.mypt.vn (qua Chrome bot, cong CDP 9333): danh sach nhan vien + diem danh thang hien tai.
- Thu muc Downloads: bao_cao_chi_tiet_ton__YYYY-MM-DD.xlsx (ton ca), lich truc*.xls(x),
  Chi so do*.xlsx / bao cao SOC*.xlsx (checklist >= 24h, checklist lap).
Chi tieu KPI doc tu kpi.json (tu tao neu chua co).
"""
import csv, datetime, glob, gzip, json, math, os, re, sys, unicodedata, urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
DL_DIR = os.environ.get("DL_DIR", os.path.join(os.path.expanduser("~"), "Downloads"))
OUT_FILE = os.environ.get("OUT_FILE", os.path.join(BASE, "dashboard_ktv.html"))
KTV_FILE = os.path.join(BASE, "ktv.csv")
KPI_FILE = os.path.join(BASE, "kpi.json")
MYPT_CACHE = os.path.join(BASE, "mypt_cache.json")
CDP_URL = "http://127.0.0.1:9333"
DEPART = os.environ.get("MYPT_DEPART", "PKHAUSR1")
API = "https://apis.fpt.vn"

DEFAULT_KPI = {
    "ngay_cong_chuan_thang": 26,
    "gio_vao_ca": "08:00",
    "dung_gio_min_pct": 95,
    "ton_ca_max": 6,
    "qua_hen_max": 0,
    "checklist24_max": 0,
    "checklist_lap_max": 0,
    "doi_truong": "Lê Đình Lâm",
}

# Bang KPI "BÁO CÁO CHỈ SỐ SỨC KHỎE | CHI TIẾT NHÂN VIÊN" (Power BI Sức khỏe chi nhánh tren ho.mypt.vn)
PBI_URL = "https://wabi-south-east-asia-api.analysis.windows.net/public/reports/querydata?synchronous=true"
PBI_QUERY = os.path.join(BASE, "pbi_suckhoe_query.json")
PBI_CACHE = os.path.join(BASE, "pbi_cache.json")
# (ten, measure, so sanh, nguong, dinh dang)
KPI_DEFS = [
    ("% Đúng hẹn", "measure.% ontime", ">=", 0.95, "pct"),
    ("TGXL TK", "measure.TGXLTB TK", "<=", 18, "h"),
    ("TGXL BT", "measure.TGXLTB BT", "<=", 8, "h"),
    ("% CLL", "measure.% CLL2", "<=", 0.085, "pct"),
    ("CSAT 1,2 CLDV", "measure.%csat12_lan2", "<=", 0.03, "pct"),
    ("NSLĐ", "measure.1 NSLĐ", ">=", 7, "num"),
    ("% Rời mạng", "measure.%RM_nv", "<=", 0.006, "pct"),
]


def log(msg):
    print(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}", flush=True)


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", s).strip()


def load_ktv():
    with open(KTV_FILE, encoding="utf-8-sig") as f:
        return [{"code": r["ma_nv"].strip().zfill(8), "name": r["ho_ten"].strip(), "phone": r["dien_thoai"].strip()}
                for r in csv.DictReader(f) if r.get("ma_nv")]


def load_kpi():
    kpi = dict(DEFAULT_KPI)
    if os.path.exists(KPI_FILE):
        with open(KPI_FILE, encoding="utf-8") as f:
            kpi.update(json.load(f))
    else:
        with open(KPI_FILE, "w", encoding="utf-8") as f:
            json.dump(kpi, f, ensure_ascii=False, indent=2)
    return kpi


# ---------- ho.mypt.vn ----------

def sanitize_mypt(raw):
    """Chi giu truong can thiet, bo thong tin ca nhan nhay cam (CMND, luong, ngan hang...)."""
    emp = {}
    for e in raw["emp"]["data"]["data"]:
        local = (e.get("email") or "").split("@")[0]
        acc = "KHA01." + local.split(".", 1)[1].upper() if "." in local else ""
        emp[e["code"]] = {"name": e.get("name"), "job": e.get("jobTitle"), "acc": acc}
    checkin = []
    for c in raw["checkin"]["data"]["data"]:
        checkin.append({"code": c["code"], "date": c["checkinDate"], "time": c.get("checkinTime"),
                        "sheet": c.get("sheetTime"), "factor": c.get("workdayFactor"),
                        "ok": c.get("checkinSuccess") == "OK", "block": c.get("blockName"), "acc": c.get("accountNameMBN")})
        if c.get("accountNameMBN") and c["code"] in emp:
            emp[c["code"]]["acc"] = c["accountNameMBN"].upper()
    return {"emp": emp, "checkin": checkin, "fetched_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")}


def fetch_mypt_raw():
    from playwright.sync_api import sync_playwright
    today = datetime.date.today()
    with sync_playwright() as p:
        ctx = p.chromium.connect_over_cdp(CDP_URL).contexts[0]
        pg = next((x for x in ctx.pages if "ho.mypt.vn" in x.url), None)
        own = pg is None
        if own:
            pg = ctx.new_page()
        hdr = {}

        def on_req(r):
            if "emp-info" in r.url and not hdr:
                hdr.update({k: v for k, v in r.headers.items() if k.lower() in ("authorization", "content-type")})
        pg.on("request", on_req)
        pg.goto("https://ho.mypt.vn/api/auth/nhansu/emp", wait_until="networkidle", timeout=60000)
        if "login" in pg.url:
            raise RuntimeError("ho.mypt.vn chua dang nhap trong Chrome bot")
        pg.wait_for_timeout(2000)
        pg.locator("button", has_text="Xem").first.click(timeout=15000, force=True)
        for _ in range(20):
            if hdr:
                break
            pg.wait_for_timeout(500)
        if not hdr:
            raise RuntimeError("Khong lay duoc token ho.mypt.vn")

        def post(path, body):
            r = pg.request.post(API + path, data=json.dumps(body), headers=hdr, timeout=60000)
            return r.json()
        raw = {
            "emp": post("/mypt-ho-profile-api/emp-info?page=1&perPage=1000",
                        {"childDepart": [DEPART], "statusWork": [1], "show": "EN"}),
            "checkin": post("/mypt-ho-checkin-api/get-info-checkin?page=1&perPage=20000",
                            {"dateStart": str(today.replace(day=1)), "dateEnd": str(today),
                             "childDepart": [DEPART], "result": [], "show": "EN"}),
        }
        if own:
            pg.close()
        return raw


LICHTRUC_CACHE = os.path.join(BASE, "lichtruc_cache.json")
LICHTRUC_URL = "http://partner.fpt.net/optimal-assign/employeemanagement/divisionthejob"

def fetch_lichtruc_raw():
    """Lich truc thang hien tai tu partner.fpt.net (can phien inside.fpt.net trong Chrome bot)."""
    from playwright.sync_api import sync_playwright
    today = datetime.date.today()
    with sync_playwright() as p:
        ctx = p.chromium.connect_over_cdp(CDP_URL).contexts[0]
        # Mo lai "Lich truc" tu menu inside.fpt.net de lam moi phien partner.fpt.net (het han sau vai phut)
        left = next((f for pg in ctx.pages for f in pg.frames if "left.asp" in f.url.lower()), None)
        if left:
            try:
                left.locator("a", has_text="Lịch trực").first.click(timeout=15000)
                left.page.wait_for_timeout(6000)
            except Exception as e:
                log(f"[LOI] mo menu Lich truc: {e}")
        fr = next((f for pg in ctx.pages for f in pg.frames if "divisionthejob" in f.url.lower()), None)
        own = None
        if fr is None:
            own = ctx.new_page()
            own.goto(LICHTRUC_URL, wait_until="networkidle", timeout=60000)
            fr = own.main_frame
            if "divisionthejob" not in fr.url.lower():
                own.close()
                raise RuntimeError("chua dang nhap inside.fpt.net / partner.fpt.net trong Chrome bot")
        # Goi API truc tiep bi Chrome chan, nen bam nut "Xem bao cao" cua trang va doc response
        with fr.page.expect_response(lambda r: "handler=ViewDivisionthejobs" in r.url, timeout=60000) as info:
            fr.get_by_text("Xem báo cáo", exact=True).first.click(timeout=15000)
        resp = info.value
        st, body = resp.status, resp.text()
        if own:
            own.close()
        if st != 200:
            raise RuntimeError(f"lich truc HTTP {st}")
        rows = json.loads(body)
        return [{k: v for k, v in r.items() if k.startswith("Day") or k in ("CodeStaff", "NameStaff", "SubTeam", "TotalDay", "Months", "Years")}
                for r in rows]


def get_lichtruc():
    try:
        rows = fetch_lichtruc_raw()
        data = {"rows": rows, "fetched_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                "month": datetime.date.today().month}
        with open(LICHTRUC_CACHE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        return data, None
    except Exception as e:
        log(f"[LOI] lich truc: {e}")
        if os.path.exists(LICHTRUC_CACHE):
            with open(LICHTRUC_CACHE, encoding="utf-8") as f:
                data = json.load(f)
            if data.get("month") == datetime.date.today().month:
                return data, f"Không lấy được lịch trực mới ({e}). Đang dùng lịch trực lúc {data.get('fetched_at')}."
        return None, f"Không lấy được lịch trực từ partner.fpt.net ({e}); dùng file lịch trực trong Downloads nếu có."


def dsr_rows(data):
    """Giai nen ket qua querydata (DSR) cua Power BI thanh list dict {ten cot: gia tri}."""
    names = {x["Value"]: x["Name"] for x in data["descriptor"]["Select"]}
    ds = data["dsr"]["DS"][0]
    vd = ds.get("ValueDicts", {})
    schema, prev, out = {}, {}, []

    def walk(key, rows, ctx):
        for row in rows:
            if "S" in row:
                schema[key] = row["S"]
                prev[key] = [None] * len(row["S"])
            cols, vals = schema[key], []
            it = iter(row.get("C", []))
            rep, nul = row.get("R", 0), row.get("Ø", 0)
            for i, col in enumerate(cols):
                if rep >> i & 1:
                    v = prev[key][i]
                elif nul >> i & 1:
                    v = None
                else:
                    v = next(it) if "C" in row else row.get(col["N"])
                    if col.get("DN") and isinstance(v, int):
                        v = vd[col["DN"]][v]
                vals.append(v)
            prev[key] = vals
            nctx = dict(ctx, **{names.get(c["N"], c["N"]): v for c, v in zip(cols, vals)})
            subs = [(k, v) for m in row.get("M", []) for k, v in m.items() if k.startswith("DM")]
            if subs:
                for k, v in subs:
                    walk(k, v, nctx)
            else:
                out.append(nctx)

    for ph in ds["PH"]:
        for k, v in ph.items():
            walk(k, v, {})
    return out


def fetch_pbi_raw(captain, year, month):
    with open(PBI_QUERY, encoding="utf-8") as f:
        q = json.load(f)
    where = q["queries"][0]["Query"]["Commands"][0]["SemanticQueryDataShapeCommand"]["Query"]["Where"]
    for w in where:
        cond = w["Condition"].get("In")
        if not cond:
            continue
        props = [e["Column"]["Property"] for e in cond["Expressions"]]
        if props == ["captain_full_name"]:
            cond["Values"] = [[{"Literal": {"Value": "'%s'" % captain.replace("'", "''")}}]]
        elif props == ["Nam_quy", "month"]:
            cond["Values"] = [[{"Literal": {"Value": f"'{year}-Q{(month - 1) // 3 + 1}'"}}, {"Literal": {"Value": f"{month}L"}}]]
    key = os.environ.get("PBI_KEY") or load_kpi().get("pbi_key")
    if not key:
        raise RuntimeError("chưa khai báo pbi_key trong kpi.json")
    req = urllib.request.Request(PBI_URL, data=json.dumps(q, ensure_ascii=False).encode(),
                                 headers={"X-PowerBI-ResourceKey": key, "Content-Type": "application/json;charset=UTF-8"})
    raw = urllib.request.urlopen(req, timeout=90).read()
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    data = json.loads(raw)["results"][0]["result"]["data"]
    res = {}
    for r in dsr_rows(data):
        code = str(r.get("Dim_captain.emp_code") or "").zfill(8)
        res[code] = {k: (float(v) if isinstance(v, str) and re.fullmatch(r"-?[\d.]+(E-?\d+)?", v) else v) for k, v in r.items()}
    return res


def get_pbi(captain):
    today = datetime.date.today()
    try:
        rows = fetch_pbi_raw(captain, today.year, today.month)
        if not rows:
            raise RuntimeError(f"không có dòng nào cho đội trưởng {captain}")
        data = {"rows": rows, "fetched_at": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                "month": today.month, "year": today.year, "captain": captain}
        with open(PBI_CACHE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        return data, None
    except Exception as e:
        log(f"[LOI] power bi: {e}")
        if os.path.exists(PBI_CACHE):
            with open(PBI_CACHE, encoding="utf-8") as f:
                data = json.load(f)
            if (data.get("month"), data.get("year"), data.get("captain")) == (today.month, today.year, captain):
                return data, f"Không lấy được KPI Power BI mới ({e}). Đang dùng số lúc {data.get('fetched_at')}."
        return None, f"Không lấy được KPI từ Power BI Sức khỏe chi nhánh ({e})."


def fmt_kpi(v, kind):
    if v is None:
        return "-"
    return f"{v * 100:.2f}%" if kind == "pct" else f"{v:.2f}"


def get_mypt():
    src = os.environ.get("MYPT_JSON")
    try:
        if src:
            with open(src, encoding="utf-8") as f:
                data = sanitize_mypt(json.load(f))
        else:
            data = sanitize_mypt(fetch_mypt_raw())
        with open(MYPT_CACHE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        return data, None
    except Exception as e:
        log(f"[LOI] mypt: {e}")
        if os.path.exists(MYPT_CACHE):
            with open(MYPT_CACHE, encoding="utf-8") as f:
                data = json.load(f)
            return data, f"Không lấy được dữ liệu ho.mypt.vn ({e}). Đang dùng dữ liệu lúc {data.get('fetched_at')}."
        return {"emp": {}, "checkin": [], "fetched_at": None}, f"Không lấy được dữ liệu ho.mypt.vn: {e}"


# ---------- Excel ----------

def read_sheets(path):
    """Tra ve {ten_sheet: [dict theo dong tieu de dau tien]}."""
    out = {}
    if path.lower().endswith(".xls"):
        import xlrd
        wb = xlrd.open_workbook(path)
        for sh in wb.sheets():
            rows = [sh.row_values(i) for i in range(sh.nrows)]
            out[sh.name.strip()] = rows
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            out[ws.title.strip()] = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    res = {}
    for name, rows in out.items():
        rows = [r for r in rows if any(v not in (None, "") for v in r)]
        if not rows:
            res[name] = []
            continue
        hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
        res[name] = [{hdr[i]: r[i] for i in range(min(len(hdr), len(r))) if hdr[i]} for r in rows[1:]]
    return res


def to_dt(v):
    if isinstance(v, datetime.datetime):
        return v
    s = str(v or "").strip()
    for fmt in ("%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def fmt_dt(d):
    return d.strftime("%d/%m %H:%M") if d else ""


def ton_files():
    """{ngay: duong_dan} - moi ngay lay file lon nhat (ban day du)."""
    best = {}
    for f in glob.glob(os.path.join(DL_DIR, "bao_cao_chi_tiet_ton__*.xlsx")):
        m = re.search(r"(\d{4}-\d{2}-\d{2})", os.path.basename(f))
        if not m:
            continue
        d = m.group(1)
        if d not in best or os.path.getsize(f) > os.path.getsize(best[d]):
            best[d] = f
    return dict(sorted(best.items()))


def latest(patterns):
    files = [f for p in patterns for f in glob.glob(os.path.join(DL_DIR, p))]
    return max(files, key=os.path.getmtime) if files else None


SHIFT_DAY = {"CA1": 1, "CA2": 1, "CA3": 1, "S": 0.5, "C": 0.5, "T": 0.5}

DONE_STATUS = {"da xu ly dang theo doi"}


def build_ton(acc_to_code):
    files = ton_files()
    trend, current, meta = {}, {}, None
    for d, f in files.items():
        rows = next(iter(read_sheets(f).values()), [])
        cnt = {}
        for r in rows:
            code = acc_to_code.get(str(r.get("Nhân sự") or "").strip().upper())
            if code:
                cnt[code] = cnt.get(code, 0) + 1
        trend[d] = cnt
    if files:
        d, f = list(files.items())[-1]
        ref = datetime.datetime.fromtimestamp(os.path.getmtime(f))
        meta = {"date": d, "file": os.path.basename(f), "exported": ref.strftime("%d/%m/%Y %H:%M")}
        for r in next(iter(read_sheets(f).values()), []):
            code = acc_to_code.get(str(r.get("Nhân sự") or "").strip().upper())
            if not code:
                continue
            xanh, do = to_dt(r.get("TG Hẹn xanh")), to_dt(r.get("TG Hẹn đỏ"))
            st = str(r.get("Tình trạng") or "")
            late = bool(xanh and xanh < ref and norm(st) not in DONE_STATUS)
            current.setdefault(code, []).append({
                "hd": str(r.get("Số HĐ") or ""), "kh": str(r.get("Tên KH") or ""), "status": st,
                "type": str(r.get("Loại triển khai") or ""), "xanh": fmt_dt(xanh), "do": fmt_dt(do),
                "xanh_ts": xanh.isoformat() if xanh else "", "late": late, "red": bool(do),
                "note": str(r.get("Ghi chú triển khai TIN/PNC") or "")[:200],
            })
    for lst in current.values():
        lst.sort(key=lambda x: (not x["late"], x["xanh_ts"] or "9"))
    return current, trend, meta


def build_checklist(acc_to_code):
    f = latest(["Chi so do*.xlsx", "Chi số đo*.xlsx", "báo cáo SOC*.xlsx", "bao cao SOC*.xlsx"])
    res = {}
    if not f:
        return res, None
    accs = sorted(acc_to_code, key=len, reverse=True)
    for sheet, rows in read_sheets(f).items():
        key = "c24" if "24h" in sheet else "lap" if "lặp" in sheet.lower() or "lap" in norm(sheet) else None
        if not key:
            continue
        for r in rows:
            text = " ".join(str(v) for v in r.values() if v is not None).upper()
            for a in accs:
                if re.search(r"(?<![A-Z0-9.])" + re.escape(a) + r"(?![A-Z0-9])", text):
                    res.setdefault(acc_to_code[a], {}).setdefault(key, 0)
                    res[acc_to_code[a]][key] += 1
                    break
    mt = datetime.datetime.fromtimestamp(os.path.getmtime(f))
    return res, {"file": os.path.basename(f), "date": mt.strftime("%d/%m/%Y")}


def build_plan(codes, year, month, lt=None):
    if lt and lt.get("rows"):
        plan = {}
        for r in lt["rows"]:
            code = str(r.get("CodeStaff") or "").zfill(8)
            if code in codes and int(r.get("Months") or 0) == month and int(r.get("Years") or 0) == year:
                plan[code] = {d: str(r.get(f"Day{d}") or "").strip() for d in range(1, 32) if str(r.get(f"Day{d}") or "").strip()}
        if plan:
            return plan, {"file": f"partner.fpt.net (lấy lúc {lt.get('fetched_at')})"}
    files = [f for pat in ["lịch trực*.xls*", "lich truc*.xls*", "Lich truc*.xls*"] for f in glob.glob(os.path.join(DL_DIR, pat))]
    for f in sorted(set(files), key=os.path.getmtime, reverse=True):
        plan = plan_from_file(f, codes, year, month)
        if plan:
            return plan, {"file": os.path.basename(f)}
    return {}, None


def plan_from_file(f, codes, year, month):
    plan = {}
    for rows in read_sheets(f).values():
        for r in rows:
            raw = r.get("CodeStaff")
            if raw in (None, ""):
                continue
            code = str(raw).split(".")[0].strip().zfill(8)
            if code not in codes:
                continue
            try:
                if int(float(r.get("Months") or 0)) != month or int(float(r.get("Years") or 0)) != year:
                    continue
            except ValueError:
                continue
            plan[code] = {d: str(r.get(f"Day{d}") or "").strip() for d in range(1, 32) if str(r.get(f"Day{d}") or "").strip()}
    return plan


# ---------- tong hop ----------

def build_data():
    ktv = load_ktv()
    kpi = load_kpi()
    mypt, mypt_warn = get_mypt()
    today = datetime.date.today()
    codes = {k["code"] for k in ktv}
    acc_to_code = {}
    for code, e in mypt["emp"].items():
        if e.get("acc"):
            acc_to_code[e["acc"].upper()] = code
    days_in_month = ((today.replace(day=28) + datetime.timedelta(days=4)).replace(day=1) - datetime.timedelta(days=1)).day

    ton, trend, ton_meta = build_ton(acc_to_code)
    chk, chk_meta = build_checklist(acc_to_code)
    lt, lt_warn = get_lichtruc()
    pbi, pbi_warn = get_pbi(kpi["doi_truong"])
    plan, plan_meta = build_plan(codes, today.year, today.month, lt)

    att = {}
    for c in mypt["checkin"]:
        if c["code"] not in codes:
            continue
        try:
            d = datetime.datetime.strptime(c["date"], "%d/%m/%Y").date()
        except ValueError:
            continue
        if d.month == today.month and d.year == today.year:
            att.setdefault(c["code"], {})[d.day] = c

    target_days = math.floor(kpi["ngay_cong_chuan_thang"] * today.day / days_in_month)
    people = []
    for k in ktv:
        code = k["code"]
        e = mypt["emp"].get(code, {})
        a = att.get(code, {})
        work = [c for c in a.values() if c["sheet"] != "O" and c["ok"]]
        cong = round(sum(c["factor"] or 0 for c in work), 2)
        ontime = sum(1 for c in work if c["time"] and c["time"][:5] <= kpi["gio_vao_ca"])
        ontime_pct = round(100 * ontime / len(work), 1) if work else None
        lst = ton.get(code, [])
        late = sum(1 for x in lst if x["late"])
        c24 = chk.get(code, {}).get("c24", 0)
        lap = chk.get(code, {}).get("lap", 0)
        t = a.get(today.day)
        pr = (pbi or {}).get("rows", {}).get(code)
        checks = []
        for name, meas, op, thr, kind in KPI_DEFS:
            v = (pr or {}).get(meas)
            v = v if isinstance(v, (int, float)) else None
            ok = None if v is None else (v >= thr - 1e-12 if op == ">=" else v <= thr + 1e-12)
            checks.append({"name": name, "val": fmt_kpi(v, kind), "target": f"{'≥' if op == '>=' else '≤'} {fmt_kpi(thr, kind).replace('.00', '')}{'h' if kind == 'h' else ''}", "ok": ok})
        ops = [
            {"name": "Ngày công", "val": cong, "target": f"≥ {target_days}", "ok": cong >= target_days - 1e-9},
            {"name": "Đúng giờ", "val": f"{ontime_pct}%" if ontime_pct is not None else "-",
             "target": f"≥ {kpi['dung_gio_min_pct']}%", "ok": ontime_pct is not None and ontime_pct >= kpi["dung_gio_min_pct"]},
            {"name": "Tồn ca", "val": len(lst), "target": f"≤ {kpi['ton_ca_max']}", "ok": len(lst) <= kpi["ton_ca_max"]},
            {"name": "Quá hẹn", "val": late, "target": f"≤ {kpi['qua_hen_max']}", "ok": late <= kpi["qua_hen_max"]},
            {"name": "Checklist ≥24h", "val": c24, "target": f"≤ {kpi['checklist24_max']}", "ok": c24 <= kpi["checklist24_max"]},
            {"name": "Checklist lặp", "val": lap, "target": f"≤ {kpi['checklist_lap_max']}", "ok": lap <= kpi["checklist_lap_max"]},
        ]
        mapped = bool(e.get("acc"))
        for c in ops[2:4]:
            if not (ton_meta and mapped):
                c["val"], c["ok"] = "-", None
        for c in ops[4:6]:
            if not (chk_meta and mapped):
                c["val"], c["ok"] = "-", None
        pl = plan.get(code)
        if pl:
            tgt = sum(SHIFT_DAY.get(pl.get(d, ""), 0) for d in range(1, today.day + 1))
            ops[0]["target"], ops[0]["ok"] = f"≥ {tgt:g} (theo lịch)", cong >= tgt - 1e-9
        cal = {}
        for d in range(1, days_in_month + 1):
            c = a.get(d)
            p = plan.get(code, {}).get(d, "")
            if c and c["sheet"] == "O":
                st = "off"
            elif c and c["ok"]:
                st = "full" if (c["factor"] or 0) >= 0.95 else "part"
            elif c:
                st = "miss"
            elif p == "O":
                st = "off" if d <= today.day else "future"
            elif d < today.day and p:
                st = "miss"
            else:
                st = "future" if d > today.day else "none"
            cal[d] = {"st": st, "plan": p, "time": (c or {}).get("time") or "", "f": (c or {}).get("factor")}
        people.append({
            "code": code, "name": k["name"], "phone": k["phone"], "acc": e.get("acc", ""), "job": e.get("job", ""),
            "block": next((c["block"] for c in a.values() if c.get("block")), ""),
            "today": {"st": cal[today.day]["st"], "time": (t or {}).get("time") or "", "plan": cal[today.day]["plan"]},
            "cong": cong, "ontime_pct": ontime_pct, "ton": lst, "late": late, "c24": c24, "lap": lap,
            "checks": checks, "score": sum(1 for c in checks if c["ok"] is True), "cal": cal, "ops": ops,
            "pbi": {"total": (pr or {}).get("measure.total_score_csatlan2"), "rank": (pr or {}).get("measure.xep_loai_csatlan2"),
                    "level": (pr or {}).get("Sum(Dim_captain.level)"), "au": (pr or {}).get("measure.AU"),
                    "tk": (pr or {}).get("measure.Slg_Triển_khai"), "bt": (pr or {}).get("measure.Slg BT")} if pr else None,
            "trend": [trend[d].get(code, 0) for d in trend],
        })
    warns = [w for w in [mypt_warn, lt_warn, pbi_warn] if w]
    if pbi:
        miss = [k["name"] for k in ktv if k["code"] not in pbi["rows"]]
        if miss:
            warns.append("Power BI chưa có KPI của: " + ", ".join(miss) + ".")
    if ton_meta and ton_meta["date"] != str(today):
        warns.append(f"Báo cáo tồn mới nhất là ngày {ton_meta['date']} — tải file bao_cao_chi_tiet_ton hôm nay vào Downloads để cập nhật.")
    if not ton_meta:
        warns.append("Không tìm thấy file bao_cao_chi_tiet_ton__*.xlsx trong Downloads.")
    return {
        "generated": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), "today": today.day, "month": today.month,
        "year": today.year, "days": days_in_month, "kpi": kpi, "people": people, "trend_dates": list(trend),
        "ton_meta": ton_meta, "chk_meta": chk_meta, "plan_meta": plan_meta, "mypt_at": mypt.get("fetched_at"), "warns": warns,
        "pbi_at": (pbi or {}).get("fetched_at"), "kpi_defs": [[n, op, t, k] for n, _, op, t, k in KPI_DEFS],
    }


def main():
    data = build_data()
    with open(os.path.join(BASE, "dashboard_template.html"), encoding="utf-8") as f:
        tpl = f.read()
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    tmp = OUT_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(tpl.replace("/*__DATA__*/null", payload))
    os.replace(tmp, OUT_FILE)
    log(f"Da tao {OUT_FILE} ({len(data['people'])} KTV)")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
