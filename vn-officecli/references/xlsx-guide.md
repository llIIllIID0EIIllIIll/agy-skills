# Hướng dẫn Xử lý Bảng tính Excel (.xlsx) với OfficeCLI

Tài liệu hướng dẫn tạo mới, nhập dữ liệu, tính toán công thức và lập báo cáo tài chính bằng `officecli`.

---

## 1. Tạo mới & Cấu trúc cơ bản

```bash
# Tạo file Excel trống
officecli create data.xlsx

# Xem cấu trúc workbook và các sheet
officecli get data.xlsx /

# Đổi tên sheet
officecli set data.xlsx "/sheet[1]" --prop name="TongHop"

# Thêm sheet mới
officecli add data.xlsx / --type sheet --prop name="ChiTiet"
```

---

## 2. Nhập dữ liệu & Công thức (Formulas)

```bash
# Ghi giá trị vào ô (sử dụng tọa độ A1 hoặc r{R}c{C})
officecli set data.xlsx "/sheet[1]/cell[A1]" --prop value="BÁO CÁO TÀI CHÍNH NĂM 2026" --prop bold=true --prop size="14pt"

# Tiêu đề cột
officecli set data.xlsx "/sheet[1]/cell[A3]" --prop value="Hạng mục" --prop bold=true
officecli set data.xlsx "/sheet[1]/cell[B3]" --prop value="Số lượng" --prop bold=true
officecli set data.xlsx "/sheet[1]/cell[C3]" --prop value="Đơn giá (VNĐ)" --prop bold=true
officecli set data.xlsx "/sheet[1]/cell[D3]" --prop value="Thành tiền (VNĐ)" --prop bold=true

# Dữ liệu & Công thức nhân
officecli set data.xlsx "/sheet[1]/cell[A4]" --prop value="Máy chủ Dell PowerEdge"
officecli set data.xlsx "/sheet[1]/cell[B4]" --prop value=2 --prop type=number
officecli set data.xlsx "/sheet[1]/cell[C4]" --prop value=85000000 --prop type=number --prop numFmt="#,##0"
officecli set data.xlsx "/sheet[1]/cell[D4]" --prop formula="=B4*C4" --prop numFmt="#,##0"

# Tổng cộng dùng SUM
officecli set data.xlsx "/sheet[1]/cell[A5]" --prop value="Tổng cộng" --prop bold=true
officecli set data.xlsx "/sheet[1]/cell[D5]" --prop formula="=SUM(D4:D4)" --prop bold=true --prop numFmt="#,##0"

# Đóng và lưu
officecli close data.xlsx
```

---

## 3. Import nhanh từ file CSV

```bash
# Đưa toàn bộ dữ liệu CSV vào sheet
officecli import report.xlsx "/sheet[1]" "data.csv"
```
