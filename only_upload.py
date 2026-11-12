import os
import openpyxl
import requests

# URL API của trang congno.click
API_URL = "https://congno.click/api/upload"

excel_file = os.path.abspath("Data.xlsx")
if not os.path.exists(excel_file):
  print(f"❌ Không tìm thấy tệp {excel_file}")
  exit()

print(f"📖 Đang đọc dữ liệu từ tệp {os.path.basename(excel_file)}...")
wb = openpyxl.load_workbook(excel_file, data_only=True)
ws = wb["BC"]

# Đọc toàn bộ các dòng dữ liệu
rows_data = [
    list(row) for row in ws.iter_rows(min_row=2, values_only=True) if any(row)
]

print(f"🚀 Đang gửi {len(rows_data)} dòng dữ liệu lên API {API_URL}...")
response = requests.post(
    API_URL, json={"data": rows_data}, headers={"Authorization": "Bearer admin123"}
)

if response.status_code == 200:
  print("🎉 ĐÃ ĐẨY DỮ LIỆU LÊN DASHBOARD CONGNO.CLICK THÀNH CÔNG!")
else:
  print(f"❌ Lỗi gửi dữ liệu: {response.status_code} - {response.text}")