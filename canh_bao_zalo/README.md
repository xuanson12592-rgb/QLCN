# Cảnh báo trễ hẹn FoxOne → Zalo KTV

Theo dõi hội thoại **FoxOne** trên FPT Chat (chat.fpt.com). Khi có tin `CẢNH BÁO ... TRỄ HẸN`,
tách từng dòng `HĐ: <mã> — KTV: <tên>`, tra SĐT trong `ktv.csv` và gửi Zalo (chat.zalo.me) **chỉ cho KTV nhận ca đó**.
KTV không có trong `ktv.csv` sẽ bị bỏ qua và ghi vào `foxone_zalo.log`.

## Cài đặt (Windows)
```powershell
winget install -e --id Python.Python.3.12 --scope user
cd C:\Users\<user>\canh_bao_zalo
& "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe" -m venv venv
.\venv\Scripts\python.exe -m pip install playwright
```
Dùng Chrome có sẵn trên máy (không cần `playwright install`).

## Chạy
- `chay_canh_bao.bat` — chạy liên tục, tự khởi động lại nếu lỗi.
- Lần đầu: script mở cửa sổ Chrome riêng (profile `zalo_profile/`, cổng 9333) → đăng nhập chat.fpt.com và quét QR Zalo. Không đóng cửa sổ này.
- `python foxone_zalo.py --dry-run` — chỉ ghi log, không gửi.
- `python foxone_zalo.py --test-send <SĐT> "nội dung"` — gửi thử 1 tin.
- `python foxone_zalo.py --parse <file.txt>` — thử tách cảnh báo từ text.

Lần chạy đầu tiên đánh dấu các cảnh báo đang có là đã xem; chỉ cảnh báo mới sau đó mới được gửi.
Trạng thái lưu ở `sent_state.json`.

## Cập nhật danh sách KTV
Sửa `ktv.csv` (cột `ho_ten` phải trùng tên KTV trong tin FoxOne, `dien_thoai` là SĐT Zalo).
