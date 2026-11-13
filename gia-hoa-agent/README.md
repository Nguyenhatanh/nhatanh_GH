# Agent Đánh Giá Chương Trình Bán Hàng - Công Ty Gia Hòa

AI Agent sử dụng Google Gemini để tự động phân tích hiệu quả 4 chương trình bán hàng.

## Cài đặt

```bash
cd gia-hoa-agent
pip install -r requirements.txt
```

## Chạy giao diện web (quy trình 5 chặng)

```bash
python server.py
```

Mở http://localhost:5000:

1. **Nhập dữ liệu**: kiểm tra file đúng mẫu (đủ 4 sheet, đủ cột, cột số là số). Nếu sai, hệ thống báo lỗi và bắt chọn lại file.
2. **Làm sạch**: chuẩn hoá, xoá trùng, xoá dòng thiếu hoặc sai giá trị, SP lạ, gộp và lọc đổi trả. Có thể bật/tắt quy tắc rồi "Làm sạch lại".
3. **Xử lý**: phân loại mã KM, phân bổ doanh thu, tính kết quả, đối soát.
4. **Xác nhận**: duyệt số liệu. Nếu chưa đúng, quay lại chặng trước.
5. **Báo cáo**: xuất file Excel `output/KetQua_GiaHoa.xlsx` và dashboard.

## Chạy agent dòng lệnh (Gemini)

```bash
export GEMINI_API_KEY='your-api-key-here'
python agent.py
```

## Kiến trúc

```
agent.py                  ← Orchestrator (Gemini function calling)
tools/
  data_loader.py          ← Bước 1: Kiểm tra dữ liệu
  revenue_allocator.py    ← Bước 2: Phân bổ doanh thu
  program_calculator.py   ← Bước 3: Tính kết quả
  reconciler.py           ← Bước 4: Đối soát & xuất file
data/
  Du_Lieu_Gia_Hoa.xlsx    ← Dữ liệu đầu vào
output/
  KetQua_GiaHoa.xlsx      ← Kết quả (tự động tạo)
```

## Quy trình xử lý

1. **Tiếp nhận & kiểm tra**: Đọc 4 sheet, loại 137 đơn mã trống/NONE
2. **Phân bổ doanh thu**: Đơn 1 mã → gán trọn; đơn nhiều mã → chia đều 1/N; trừ đổi trả
3. **Tính kết quả**: Doanh thu thuần - Giá vốn thuần - Chi phí cố định = Phần đóng góp
4. **Đối soát & xuất**: Kiểm tra tổng khớp → Xuất Excel 4 sheet
