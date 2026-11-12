import json
import os
from github import Github
import openpyxl

EXCEL_FILE = os.path.abspath("Data.xlsx")
JS_FILE = os.path.abspath("data.js")

# --- CẤU HÌNH GITHUB API ---
GITHUB_TOKEN = "ghp_dian_token_github_cua_ban_vao_day"  # Thay bằng Token vừa tạo
REPO_NAME = (
    "username_cua_ban/ten_repository"  # Ví dụ: "tranvana/dashboard-congno"
)
FILE_PATH_IN_REPO = "data.js"  # Đường dẫn file data.js trên repo


def update_web_data():
  if not os.path.exists(EXCEL_FILE):
    print(f"❌ Không tìm thấy tệp {EXCEL_FILE}")
    return

  print(
      f"📖 1. Đang đọc dữ liệu từ sheet 'BC' trong"
      f" {os.path.basename(EXCEL_FILE)}..."
  )
  wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)

  if "BC" not in wb.sheetnames:
    print("❌ Không tìm thấy sheet 'BC' trong tệp Excel!")
    return

  ws = wb["BC"]
  headers = [cell for cell in next(ws.iter_rows(values_only=True))]

  data_list = []
  for row in ws.iter_rows(min_row=2, values_only=True):
    if any(row):
      row_dict = {
          str(headers[i]): (str(val) if val is not None else "")
          for i, val in enumerate(row)
          if i < len(headers)
      }
      data_list.append(row_dict)

  print(f"📊 Đã xử lý {len(data_list)} dòng dữ liệu.")

  # 2. Tạo nội dung tệp data.js
  print("📝 2. Đang cập nhật dữ liệu vào tệp data.js...")
  js_content = f"const rawData = {json.dumps(data_list, ensure_ascii=False)};"

  with open(JS_FILE, "w", encoding="utf-8") as f:
    f.write(js_content)

  print("✅ Đã cập nhật tệp data.js local!")

  # 3. Đẩy file lên GitHub qua API
  print(f"🚀 3. Đang đẩy dữ liệu lên GitHub repository ({REPO_NAME})...")
  try:
    g = Github(GITHUB_TOKEN)
    repo = g.get_repo(REPO_NAME)

    try:
      # Nếu file đã tồn tại trên repo -> Update
      contents = repo.get_contents(FILE_PATH_IN_REPO)
      repo.update_file(
          contents.path,
          "Auto update BC data via Python API",
          js_content,
          contents.sha,
      )
      print("✅ Đã update tệp data.js lên GitHub!")
    except Exception:
      # Nếu file chưa có -> Create
      repo.create_file(
          FILE_PATH_IN_REPO,
          "Auto upload initial BC data via Python API",
          js_content,
      )
      print("✅ Đã tạo mới tệp data.js trên GitHub!")

    print(
        "🎉 HOÀN TẤT! Vercel sẽ tự động deploy dữ liệu mới lên congno.click"
        " trong vài giây."
    )

  except Exception as e:
    print(f"❌ Lỗi kết nối GitHub API: {e}")


if __name__ == "__main__":
  update_web_data()