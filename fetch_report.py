# -*- coding: utf-8 -*-
"""Tai bao cao 'Danh sach hoa don chi tiet' tu inside.fpt.net (FOX).

Quy trinh:
  1. Mo inside.fpt.net bang profile da luu session (inside_profile).
  2. Neu bi day ve trang login -> /Home/LoginIAMRequest -> bam 'Azure AD'
     (SSO Microsoft tu xac thuc bang cookie ESTSAUTHPERSISTENT).
  3. Mo trang InternetSubs/Accounting/BillList, goi ajax ViewDetailExport
     voi filter: Vung '06.Khanh Hoa' (303), Chi nhanh 'KHA_01', thang hien tai.
  4. Tai file xlsx ve thu muc downloads/ va in duong dan ra stdout.
"""
import os
import sys
import time
from datetime import datetime, timedelta

from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "inside_profile")
OUT_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(OUT_DIR, exist_ok=True)

SUB_COMPANY_ID = "303"          # 06.Khanh Hoa
SUB_COMPANY_NAME = "06.Khanh Hoa"
BRANCH_ID = "KHA_01"            # KHA - Khanh Hoa 01

INSIDE_URL = "http://inside.fpt.net/"
IAM_REQUEST_URL = "http://login.fpt.net/Home/LoginIAMRequest"
BILLLIST_URL = "http://mvc.inside.fpt.net/InternetSubs/Accounting/BillList?vUrl="


def log(msg):
    print(f"[fetch {datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def ensure_login(page, timeout=240):
    """Dieu huong toi inside.fpt.net; xu ly chuoi login -> IAM -> Azure AD."""
    deadline = time.time() + timeout
    clicked_iam = clicked_azure = False
    while time.time() < deadline:
        url = page.url.lower()
        host_path = url.split("?")[0]

        if "inside.fpt.net" in url and "login" not in url and "login.asp" not in url:
            return True

        if "login.fpt.net" in url and not clicked_iam:
            clicked_iam = True
            log("Dang o login.fpt.net -> chuyen /Home/LoginIAMRequest")
            page.goto(IAM_REQUEST_URL, timeout=60000)
            page.wait_for_timeout(4000)
            continue

        if "iam.fpt.net" in url and not clicked_azure:
            try:
                az = page.locator("a#social-azure")
                if az.count() > 0:
                    az.first.click()
                    clicked_azure = True
                    log("Da bam 'Azure AD' — cho Microsoft SSO xac thuc")
                    page.wait_for_timeout(5000)
                    continue
            except Exception as e:
                log(f"Khong bam duoc Azure AD: {e}")
            page.wait_for_timeout(3000)
            continue

        # dang o trang microsoft -> cho redirect, hoac trang khac -> cho tiep
        page.wait_for_timeout(3000)
    return False


def export_and_download(page):
    """Goi ajax ViewDetailExport tren trang BillList, roi tai file ve."""
    page.goto(BILLLIST_URL, timeout=90000, wait_until="domcontentloaded")
    page.wait_for_selector("#btnViewReport", timeout=60000)
    page.wait_for_timeout(3000)

    now = datetime.now()
    # Bao cao cong no luon chay cho thang TRUOC (T-1): hoa don phat hanh cuoi thang truoc
    prev_month_last = now.replace(day=1) - timedelta(days=1)
    from_date = f"{prev_month_last.month:02d}/01/{prev_month_last.year}"
    log(f"Mo trang BillList OK: {page.url[:90]}")
    log(f"Filter: vung={SUB_COMPANY_NAME}({SUB_COMPANY_ID}) nhanh={BRANCH_ID} fromDate={from_date}")

    result = page.evaluate(
        r"""(args) => new Promise((resolve) => {
            try {
                // Gia lap trang thai nhu user da click dong '06.Khanh Hoa' level 1
                tempVariable.CodeName = 'A' + args.subId;
                tempVariable.Level = 1;
                lstParmeter.branchName = args.subName;
                fromDate = args.fromDate;
                month = args.fromDate.split('/')[0];
                year = args.fromDate.split('/')[2];

                // Chon chi nhanh tren multiselect de getUCControlValue doc dung
                try {
                    var ms = $("#cbbBranch").data("kendoMultiSelect");
                    if (ms) ms.value([args.branchId]);
                } catch (e) {}

                // Dam bao cac filter mac dinh giong luc user xuat
                $("#cboStatus").val("1");
                $("#cboSwitching").val("-1");
                $("#cboTypeReport").val("0");

                var d = GetViewDetailColumn();
                var fileName = "DanhSachHoaDonChiTiet_" + $.now() + "_" + args.subName + ".xlsx";
                var branchID = getUCControlValue("UC_Branch");
                var filter = {
                    subCompanyID: args.subId,
                    locationID: "",
                    branchID: branchID,
                    fromDate: fromDate,
                    paymentTypeID: ($("#cboPayment").val() + "").replace(/null/g, ""),
                    statusID: ($("#cboStatus").val() + "").replace(/null/g, ""),
                    onlineTypeID: ($("#cboOnlineType").val() + "").replace(/null/g, ""),
                    serviceTypeID: ($("#cboServiceType").val() + "").replace(/null/g, ""),
                    issuer: $("#cboIssuer").val(),
                    staffname: $("#txtStaffName").val(),
                    switching: $("#cboSwitching").val(),
                    collector: $("#txtCollector").val(),
                    statusVP: ($("#cboStatusVP").val() + "").replace(/null/g, ""),
                    statusTT: ($("#cboStatusTT").val() + "").replace(/null/g, ""),
                    paidaction: ($("#cboPaidAction").val() + "").replace(/null/g, ""),
                    typereport: ($("#cboTypeReport").val() + "").replace(/null/g, ""),
                    fileName: fileName,
                    colTitleArr: d.colTitleArr.slice(),
                    colNameArr: d.colNameArr.slice(),
                    DVKH: ($("#cboDVKH").val() + "").replace(/null/g, "")
                };
                filter.colTitleArr.splice(0, 0, "STT");
                filter.colNameArr.splice(0, 0, "STT");

                $.ajax({
                    type: "post",
                    url: "/InternetSubs/Accounting/BillList/ViewDetailExport",
                    data: filter,
                    traditional: true,
                    timeout: 600000,
                    cache: false,
                    success: function (r) {
                        resolve({ok: true, fileName: fileName, branchID: branchID, resp: (""+r).slice(0,200)});
                    },
                    error: function (x, e) {
                        resolve({ok: false, err: (x && x.status) + " " + e, fileName: fileName});
                    }
                });
            } catch (ex) {
                resolve({ok: false, err: "JS: " + ex});
            }
        })""",
        {
            "subId": SUB_COMPANY_ID,
            "subName": SUB_COMPANY_NAME,
            "branchId": BRANCH_ID,
            "fromDate": from_date,
        },
    )
    log(f"Ket qua export: {result}")
    if not result.get("ok"):
        return None

    file_name = result["fileName"]
    log(f"Server da tao file, dang tai: {file_name}")
    resp = page.request.get(
        f"http://mvc.inside.fpt.net/Upload/BillList/{file_name}", timeout=300000
    )
    if not resp.ok:
        log(f"Tai file loi HTTP {resp.status}")
        return None
    body = resp.body()
    if len(body) < 10000 or not body[:2] == b"PK":
        log(f"File tai ve khong hop le ({len(body)} bytes)")
        return None

    out_path = os.path.join(BASE_DIR, file_name)
    with open(out_path, "wb") as f:
        f.write(body)
    log(f"DA TAI: {out_path} ({len(body)} bytes)")
    return out_path


def main():
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(
            PROFILE_DIR,
            headless=True,
            accept_downloads=True,
            args=["--no-sandbox"],
        )
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto(INSIDE_URL, timeout=90000, wait_until="domcontentloaded")
            page.wait_for_timeout(5000)
            if not ensure_login(page):
                log("LOI: khong dang nhap duoc inside.fpt.net")
                try:
                    page.screenshot(path=os.path.join(OUT_DIR, "login_fail.png"))
                except Exception:
                    pass
                sys.exit(2)
            log("Da dang nhap inside.fpt.net")
            path = export_and_download(page)
            sys.exit(0 if path else 3)
        finally:
            ctx.close()


if __name__ == "__main__":
    main()
