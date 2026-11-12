"""Đối soát phân công Loc AU với báo cáo công nợ FPT và cập nhật Google Sheet."""

import os
import re
from datetime import datetime
from urllib.request import Request, urlopen

import gspread
import pandas as pd
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))

GOOGLE_SHEET_ID = (
    os.getenv("GOOGLE_SHEET_ID") or "1W2JbsnZr3fJTihZMro70HMhUWm1cYYSt5s26MVdwSHc"
).strip()
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
EXCEL_FILE = os.path.join(BASE_DIR, "Data.xlsx")
SHEET_NAME = "BC"
T1_SHEET_NAMES = ("Tháng T-1", "Tháng T-")
UPLOAD_CHUNK_ROWS = 2000
VERCEL_DEPLOY_HOOK = os.getenv("VERCEL_DEPLOY_HOOK", "").strip()
OUTPUT_COLUMNS = (
    "khu vực",
    "HD",
    "Tên KH",
    "Nhân sự DVKH",
    "Nhân viên thu",
    "Tình Trạng",
    "Kết quả đối soát",
    "Mức độ xử lý",
    "Số Ngày TT muộn so với T-1",
    "Tuổi thọ HD",
    "Trạng thái chăm sóc",
)


def text_value(value):
  """Đổi ô Excel, kể cả NaN và ngày, thành chuỗi. Ô trống thành chuỗi rỗng."""
  if value is None or pd.isna(value):
    return ""
  if isinstance(value, pd.Timestamp):
    return value.strftime("%Y-%m-%d")
  if isinstance(value, datetime):
    return value.strftime("%Y-%m-%d")
  if isinstance(value, float) and value.is_integer():
    return str(int(value))
  return str(value).strip()


def normalize_contract(value):
  return text_value(value).strip().upper()


def payment_day(value):
  """Lấy ngày trong tháng từ ngày Excel hoặc chuỗi tháng/ngày như 10/5."""
  if value is None or pd.isna(value):
    return None
  if isinstance(value, (pd.Timestamp, datetime)):
    return value.day
  if isinstance(value, (int, float)):
    if 1 <= value <= 31:
      return int(value)
    if value > 31:
      parsed = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
      return None if pd.isna(parsed) else parsed.day
  text = str(value).strip()
  if not text:
    return None
  month_day = re.fullmatch(r"\d{1,2}\s*/\s*(\d{1,2})", text)
  if month_day:
    day = int(month_day.group(1))
    return day if 1 <= day <= 31 else None
  parsed = parse_dates(pd.Series([text])).iloc[0]
  return None if pd.isna(parsed) else parsed.day


def coalesce(frame, candidates):
  """Lấy giá trị đầu tiên không rỗng. Thiếu cột thì trả về chuỗi rỗng."""
  if frame is None:
    return pd.Series(dtype="object")
  if frame.empty:
    return pd.Series([""] * len(frame), index=frame.index, dtype="object")
  lookup = {str(column).strip(): column for column in frame.columns}
  found = [lookup[name] for name in candidates if name in lookup]
  if not found:
    return pd.Series([""] * len(frame), index=frame.index, dtype="object")

  chosen = frame[found[0]].map(text_value)
  for column in found[1:]:
    fallback = frame[column].map(text_value)
    chosen = chosen.where(chosen != "", fallback)
  return chosen


def parse_dates(series):
  """Đọc ngày dạng datetime hoặc chuỗi dd/mm/yyyy."""
  if series is None or series.empty:
    return pd.Series(dtype="datetime64[ns]")
  if pd.api.types.is_datetime64_any_dtype(series):
    return pd.to_datetime(series, errors="coerce")
  text = series.map(lambda value: "" if pd.isna(value) else str(value).strip())
  parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
  iso_dates = text.str.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", na=False)
  if iso_dates.any():
    parsed.loc[iso_dates] = pd.to_datetime(
        text.loc[iso_dates], yearfirst=True, dayfirst=False, errors="coerce"
    )
  return parsed


def raw_column(frame, candidates):
  lookup = {str(column).strip(): column for column in frame.columns}
  for name in candidates:
    if name in lookup:
      return frame[lookup[name]]
  return None


def severity(days):
  """Cùng ngưỡng đang dùng trên congno.click."""
  if days == "" or pd.isna(days):
    return "Cần theo dõi"
  if int(days) > 10:
    return "Cực gấp cần OS XL ngay"
  if int(days) > 5:
    return "gấp"
  return "Cần theo dõi"


def contract_age(frame, today):
  """Số tháng từ ngày khởi tạo hợp đồng, hoặc cột Số tháng sử dụng."""
  start = raw_column(
      frame,
      (
          "Ngày khởi tạo hợp đồng",
          "Ngày tạo hợp đồng",
          "Ngày lắp đặt",
          "Ngày ký hợp đồng",
      ),
  )
  ages = pd.Series([""] * len(frame), index=frame.index, dtype="object")
  if start is not None:
    parsed = parse_dates(start)
    months = (today.year - parsed.dt.year) * 12 + today.month - parsed.dt.month
    months -= today.day < parsed.dt.day
    ages = months.astype("Int64").astype("object").where(parsed.notna(), "")

  usage = coalesce(frame, ("Số tháng sử dụng", "Tuổi thọ HD"))
  return ages.where(ages != "", usage)


def care_status(frame):
  """Xác định trạng thái chăm sóc từ ghi chú ở cột AM của báo cáo FPT."""
  notes = raw_column(
      frame,
      (
          "Ghi chú/Nội dung chăm sóc",
          "Ghi chú",
          "Nội dung chăm sóc",
      ),
  )
  if notes is None and len(frame.columns) > 38:
    notes = frame.iloc[:, 38]
  if notes is None:
    return pd.Series(["Chưa chăm sóc"] * len(frame), index=frame.index)
  return notes.map(
      lambda value: "Đã chăm sóc" if text_value(value) else "Chưa chăm sóc"
  )


def fpt_collector_column(frame):
  """Đọc cột Nhân viên thu từ FPT, ưu tiên tên cột cụ thể và fallback cột AL (index 37)."""
  collector = raw_column(
      frame,
      (
          "Nhân viên thu",
          "NV thu",
          "NV THU",
          "Nhân viên thu cước",
          "Nhân viên thu cước cũ",
          "Nguồn thu",
          "AL",
      ),
  )
  if collector is not None:
    return collector
  if len(frame.columns) > 37:
    return frame.iloc[:, 37]
  return pd.Series([""] * len(frame), index=frame.index)


def fpt_overdue_days(frame):
  """Lấy số ngày thanh toán muộn từ FPT nếu không có trong Loc AU; mặc định 0 nếu không có dữ liệu."""
  candidates = (
      "Số ngày TT muộn",
      "Số ngày muộn",
      "Số ngày chậm",
      "Số ngày trễ",
      "Số ngày TT chậm",
      "Số ngày thanh toán muộn",
      "Ngày muộn",
      "Số ngày chậm TT",
  )
  for candidate in candidates:
    column = raw_column(frame, (candidate,))
    if column is not None:
      return pd.to_numeric(column, errors="coerce").fillna(0).astype("int64")
  if len(frame.columns) > 39:
    return pd.to_numeric(frame.iloc[:, 39], errors="coerce").fillna(0).astype("int64")
  if len(frame.columns) > 37:
    return pd.to_numeric(frame.iloc[:, 37], errors="coerce").fillna(0).astype("int64")
  return pd.Series([0] * len(frame), index=frame.index, dtype="int64")


def load_loc_au():
  if not os.path.exists(EXCEL_FILE):
    raise FileNotFoundError(f"Không tìm thấy file phân công: {EXCEL_FILE}")
  excel = pd.ExcelFile(EXCEL_FILE)
  sheet_lookup = {str(name).strip().casefold(): name for name in excel.sheet_names}
  sheet_name = sheet_lookup.get("loc au")
  if sheet_name is None:
    raise ValueError("Data.xlsx không có sheet 'Loc AU'.")
  assignments = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name).dropna(how="all")
  assignments.columns = [str(column).strip() for column in assignments.columns]
  if not assignments.empty and raw_column(assignments, ("Hợp đồng", "HD")) is None:
    if len(assignments.columns) <= 1:
      raise ValueError("Sheet 'Loc AU' không có cột Hợp đồng ở cột B.")
    assignments["Hợp đồng"] = assignments.iloc[:, 1]
  return assignments


def load_previous_month(source_file, sheet_names=T1_SHEET_NAMES):
  """Đọc sheet tháng đối chiếu từ file FPT hoặc Data.xlsx."""
  paths = []
  for path in (source_file, EXCEL_FILE):
    if path and os.path.exists(path) and path not in paths:
      paths.append(path)

  for path in paths:
    excel = pd.ExcelFile(path)
    names = {str(name).strip(): name for name in excel.sheet_names}
    for wanted in sheet_names:
      if wanted not in names:
        continue
      frame = pd.read_excel(path, sheet_name=names[wanted])
      frame = frame.dropna(how="all")
      if frame.empty:
        continue
      frame.columns = [str(column).strip() for column in frame.columns]
      return frame

  return pd.DataFrame()


def build_report(assignments, current, comparison=None):
  """Đối soát toàn bộ hợp đồng FPT, bổ sung phân công từ Loc AU nếu có."""
  assignments = assignments.copy()
  current = current.copy()
  comparison = comparison.copy() if comparison is not None else pd.DataFrame()
  for frame in (assignments, current, comparison):
    frame.columns = [str(column).strip() for column in frame.columns]

  assignment_contract = raw_column(assignments, ("Hợp đồng", "HD"))
  if assignment_contract is None and len(assignments.columns) > 1:
    assignment_contract = assignments.iloc[:, 1]
  if assignment_contract is None:
    raise ValueError("Sheet 'Loc AU' không có cột Hợp đồng.")
  current_contract = raw_column(current, ("Hợp đồng", "HD"))
  if current_contract is None:
    raise ValueError("File FPT không có cột Hợp đồng.")

  assignment_column = raw_column(
      assignments,
      ("Phân công T11", "Phân T11", "Phân công", "Nhân viên phân công", "Phân T10"),
  )
  if assignment_column is None and len(assignments.columns) > 3:
    assignment_column = assignments.iloc[:, 3]

  previous_date_column = raw_column(
      assignments,
      ("Ngày TT Tháng 10", "Ngày thanh toán tháng 10", "Ngày TT tháng trước"),
  )
  if previous_date_column is None and len(assignments.columns) > 9:
    previous_date_column = assignments.iloc[:, 9]

  area_column = raw_column(assignments, ("khu vực", "Vùng", "Chi nhánh"))
  name_column = raw_column(assignments, ("Họ tên", "Tên KH"))
  assignment_rows = pd.DataFrame({
      "__hd": assignment_contract.map(normalize_contract),
      "__assigned_area": area_column.map(text_value) if area_column is not None else "",
      "__assigned_name": name_column.map(text_value) if name_column is not None else "",
      "__collector": assignment_column.map(text_value) if assignment_column is not None else "",
      "__previous_date": previous_date_column if previous_date_column is not None else "",
      "__loc_match": True,
  }, index=assignments.index)
  assignment_rows = assignment_rows.loc[assignment_rows["__hd"] != ""].drop_duplicates(
      "__hd", keep="last"
  )

  current["__hd"] = current_contract.map(normalize_contract)
  current["__care"] = care_status(current)
  today = datetime.now()
  current["__age"] = contract_age(current, pd.Timestamp(today.date()))
  fpt_collector = fpt_collector_column(current)
  fpt_overdue = fpt_overdue_days(current)
  current_rows = pd.DataFrame({
      "__hd": current["__hd"],
      "__fpt_area": coalesce(current, ("Quận (Lắp đặt)", "Chi nhánh", "Vùng")),
      "__fpt_name": coalesce(current, ("Họ tên", "Tên KH")),
      "__staff": coalesce(current, ("NVVP AU", "Acc được PC chăm sóc RM", "Nhân sự DVKH")),
      "__fpt_status": coalesce(current, ("Tình trạng", "Tình Trạng")),
      "__fpt_reconciliation": coalesce(current, ("Kết quả đối soát",)),
      "__collector_fpt": fpt_collector.map(text_value),
      "__fpt_overdue": fpt_overdue,
      "__age": current["__age"],
      "__care": current["__care"],
  }, index=current.index)
  current_rows = current_rows.loc[current_rows["__hd"] != ""]

  previous_day_by_contract = {}
  if not comparison.empty:
    comparison_contract = raw_column(comparison, ("Hợp đồng", "HD"))
    if comparison_contract is None and len(comparison.columns) > 1:
      comparison_contract = comparison.iloc[:, 1]
    comparison_day_column = raw_column(
        comparison,
        ("Ngày TT Tháng 10", "Ngày thanh toán Tháng 10", "Ngày tt", "Ngày thanh toán"),
    )
    if comparison_day_column is None and len(comparison.columns) > 4:
      comparison_day_column = comparison.iloc[:, 4]
    if comparison_contract is not None and comparison_day_column is not None:
      for contract, date_value in zip(comparison_contract, comparison_day_column):
        key = normalize_contract(contract)
        day = payment_day(date_value)
        if key and day is not None:
          previous_day_by_contract[key] = day

  report = current_rows.merge(assignment_rows, on="__hd", how="left", sort=False)
  has_loc_assignment = report["__loc_match"].fillna(False).astype(bool)
  previous_day = report["__previous_date"].map(payment_day)
  fallback_day = report["__hd"].map(previous_day_by_contract)
  previous_day = previous_day.where(previous_day.notna(), fallback_day).fillna(0)
  loc_delay = pd.to_numeric(previous_day, errors="coerce").fillna(0).astype("int64")
  fpt_delay = pd.to_numeric(report["__fpt_overdue"], errors="coerce").fillna(0).astype("int64")

  cycle_due_day = loc_delay.where(has_loc_assignment, 0).astype("int64")
  overdue_days = ((today.day - cycle_due_day).clip(lower=0)
                  .where((today.day > cycle_due_day) & (cycle_due_day > 0), 0)
                  .astype("int64"))
  overdue_days = overdue_days.where(has_loc_assignment | (fpt_delay > 0), fpt_delay)
  overdue_days = overdue_days.fillna(0).astype("int64")

  assigned_area = report["__assigned_area"].fillna("")
  report["khu vực"] = assigned_area.where(
      assigned_area != "", report["__fpt_area"]
  )
  report["HD"] = report["__hd"]
  assigned_name = report["__assigned_name"].fillna("")
  report["Tên KH"] = assigned_name.where(
      assigned_name != "", report["__fpt_name"]
  )
  report["Nhân sự DVKH"] = report["__staff"].fillna("")
  report["Nhân viên thu"] = report["__collector_fpt"].fillna(report["__collector"]).fillna("")
  report["Tình Trạng"] = report["__fpt_status"].fillna("").replace("", "Chưa thanh toán")
  fpt_reconciliation = report["__fpt_reconciliation"].fillna("").astype(str).str.strip()
  default_reconciliation = has_loc_assignment.map({
      True: "Chưa thanh toán",
      False: "Chưa đối soát",
  })
  report["Kết quả đối soát"] = fpt_reconciliation.where(
      fpt_reconciliation != "", default_reconciliation
  )
  processing_level = overdue_days.map(severity)
  report["Mức độ xử lý"] = processing_level.where(
      has_loc_assignment & overdue_days.gt(0), "BÌNH THƯỜNG"
  )
  report["Số Ngày TT muộn so với T-1"] = overdue_days
  report["Tuổi thọ HD"] = report["__age"].fillna("")
  report["Trạng thái chăm sóc"] = report["__care"].fillna("Chưa chăm sóc")
  report = report.loc[:, OUTPUT_COLUMNS].fillna("")
  processing_level = report["Mức độ xử lý"].astype(str).str.strip()
  normalized_level = processing_level.str.lower()
  report["Mức độ xử lý"] = normalized_level.map({
      "gấp": "GẤP",
      "cần theo dõi": "CẦN THEO DÕI",
      "cực gấp": "CỰC KỲ GẤP",
      "cực kỳ gấp": "CỰC KỲ GẤP",
      "cực gấp cần os xl ngay": "CỰC KỲ GẤP",
  }).fillna(processing_level.str.upper())
  for column in ("khu vực", "Nhân sự DVKH", "Nhân viên thu"):
    report[column] = report[column].map(normalize_contract)
  return report


def normalize_fpt_contract_columns(frame):
  """Làm sạch và chuẩn hóa tên cột Hợp đồng trong file FPT."""
  if frame is None or frame.empty:
    return frame

  frame = frame.copy()
  frame.columns = [str(column).strip() for column in frame.columns]
  renamed = {}
  aliases = {
      "hợp đồng": "Hợp đồng",
      "số hđ": "Hợp đồng",
      "số hợp đồng": "Hợp đồng",
      "so hd": "Hợp đồng",
      "mã hđ": "Hợp đồng",
      "ma_hd": "Hợp đồng",
      "contract": "Hợp đồng",
      "hd": "Hợp đồng",
  }
  for column in frame.columns:
    key = str(column).strip().lower()
    if key in aliases:
      renamed[column] = aliases[key]
  if renamed:
    frame.rename(columns=renamed, inplace=True)
  return frame


def find_contract_header_row(frame):
  """Tìm dòng header chứa mã HĐ nếu file Excel có ghi chú/thêm dòng trước header."""
  if frame is None or frame.empty:
    return 0

  for row_index in range(min(len(frame), 20)):
    row_values = frame.iloc[row_index].fillna("").tolist()
    joined = " ".join(str(value).strip().lower() for value in row_values)
    if any(token in joined for token in (
        "hđ",
        "hợp đồng",
        "so hd",
        "số hợp đồng",
        "số hđ",
        "mã hđ",
        "contract",
        "hd",
    )):
      return row_index
  return 0


def plain_cell(value):
  if hasattr(value, "item"):
    return value.item()
  return "" if pd.isna(value) else value


def prepare_report(source_file):
  """Đọc phân công Loc AU và đối soát với file FPT mới nhất."""
  if not source_file or not os.path.exists(source_file):
    raise FileNotFoundError(f"Không tìm thấy file báo cáo: {source_file}")

  raw_excel = pd.read_excel(source_file, header=None)
  header_row = find_contract_header_row(raw_excel)
  if header_row and header_row > 0:
    data = raw_excel.iloc[header_row:].reset_index(drop=True)
    data.columns = [str(column).strip() for column in data.iloc[0].tolist()]
    data = data.iloc[1:].reset_index(drop=True)
  else:
    data = pd.read_excel(source_file)

  current = normalize_fpt_contract_columns(data)
  assignments = load_loc_au()
  comparison = load_previous_month(EXCEL_FILE, ("Tháng T-",))
  report = build_report(assignments, current, comparison)
  rows = [
      list(OUTPUT_COLUMNS),
      *([plain_cell(value) for value in record] for record in report.itertuples(index=False)),
  ]
  return rows


def get_latest_fpt_file():
  directories = (BASE_DIR, os.path.expanduser("~/Downloads"))
  reports = []
  for directory in directories:
    if not os.path.isdir(directory):
      continue
    reports.extend(
        entry.path
        for entry in os.scandir(directory)
        if entry.is_file()
        and os.path.splitext(entry.name)[1].lower() in (".xlsx", ".xls")
        and entry.name.lower() != "data.xlsx"
        and not entry.name.startswith("~$")
    )
  if not reports:
    raise FileNotFoundError(
        "Không tìm thấy file báo cáo FPT (.xlsx/.xls) trong thư mục dự án "
        "hoặc ~/Downloads."
    )

  latest = max(reports, key=os.path.getmtime)
  print(f"📂 Đã tìm thấy file báo cáo FPT mới nhất tại: {latest}")
  return latest


# Đồng bộ báo cáo 11 cột lên tab BC.
def sync_data_to_google_sheet(rows):
  if not os.path.exists(CREDENTIALS_FILE):
    raise FileNotFoundError("Không thấy credentials.json trong thư mục QLCN.")
  if not GOOGLE_SHEET_ID:
    raise ValueError("Chưa có biến GOOGLE_SHEET_ID trong tệp .env.")
  if not rows:
    raise ValueError("Không có dữ liệu để ghi lên Google Sheet.")

  client = gspread.service_account(filename=CREDENTIALS_FILE)
  spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)
  try:
    sheet = spreadsheet.worksheet(SHEET_NAME)
  except gspread.WorksheetNotFound:
    sheet = spreadsheet.add_worksheet(
        title=SHEET_NAME,
        rows=max(len(rows), 1),
        cols=max(len(rows[0]), 1),
    )

  sheet.clear()
  sheet.resize(rows=max(len(rows), 1), cols=max(len(rows[0]), 1))
  sheet.update(
      "A1",
      rows,
      value_input_option="RAW",
  )

  if VERCEL_DEPLOY_HOOK:
    request = Request(VERCEL_DEPLOY_HOOK, data=b"", method="POST")
    try:
      with urlopen(request, timeout=30) as response:
        if response.status >= 400:
          raise RuntimeError(f"Deploy Hook trả HTTP {response.status}.")
      print("🚀 Đã phát tín hiệu làm mới giao diện Web trên Vercel!")
    except Exception as error:
      print(f"Cảnh báo: không gọi được Vercel Deploy Hook: {error}")


def validate_config():
  missing = []
  if not GOOGLE_SHEET_ID:
    missing.append("GOOGLE_SHEET_ID")
  if not os.path.exists(CREDENTIALS_FILE):
    missing.append("credentials.json")
  if missing:
    raise ValueError("Thiếu cấu hình: " + ", ".join(missing))


def main():
  validate_config()
  report_file = get_latest_fpt_file()
  rows = prepare_report(report_file)
  sync_data_to_google_sheet(rows)
  print(
      f"✅ Đã giữ lại toàn bộ {len(rows) - 1} HĐ từ file FPT gốc "
      "và đẩy lên Google Sheets tab 'BC'."
  )
  print(
      "🎉 Đã đối soát, đẩy dữ liệu lên Google Sheets và cập nhật trang Web "
      "thành công!"
  )


if __name__ == "__main__":
  main()
