# Hướng dẫn Thiết kế Slide PowerPoint (.pptx) với OfficeCLI

Tài liệu hướng dẫn tạo bài thuyết trình, pitch deck và báo cáo trình chiếu bằng `officecli`.

---

## 1. Khởi tạo & Quản lý Slide

```bash
# Tạo file PowerPoint trống
officecli create presentation.pptx

# Xem danh sách các slide
officecli get presentation.pptx /

# Thêm một slide mới
officecli add presentation.pptx / --type slide
```

---

## 2. Thêm Hình khối, Văn bản & Bố cục

```bash
# Thêm tiêu đề chính trên Slide 1
officecli add presentation.pptx /slide[1] --type shape --prop shapeType=rectangle --prop x="2cm" --prop y="3cm" --prop width="20cm" --prop height="3cm" --prop text="KẾ HOẠCH PHÁT TRIỂN 2026" --prop bold=true --prop size="32pt" --prop color="003366"

# Thêm phụ đề (Subtitle)
officecli add presentation.pptx /slide[1] --type shape --prop shapeType=rectangle --prop x="2cm" --prop y="6cm" --prop width="20cm" --prop height="2cm" --prop text="Báo cáo Hội đồng Quản trị và Ban Lãnh đạo" --prop size="18pt" --prop color="666666"

# Thêm slide nội dung (Slide 2)
officecli add presentation.pptx / --type slide
officecli add presentation.pptx /slide[2] --type shape --prop shapeType=rectangle --prop x="2cm" --prop y="1.5cm" --prop width="20cm" --prop height="1.5cm" --prop text="I. MỤC TIÊU CHIẾN LƯỢC" --prop bold=true --prop size="24pt"

# Thêm hộp nội dung chi tiết
officecli add presentation.pptx /slide[2] --type shape --prop shapeType=rectangle --prop x="2cm" --prop y="3.5cm" --prop width="18cm" --prop height="8cm" --prop text="• Tăng trưởng doanh thu 25% so với cùng kỳ.\n• Hoàn tất chuyển đổi số toàn diện các phòng ban.\n• Tối ưu chi phí vận hành 15% thông qua tự động hóa." --prop size="16pt"

# Đóng và lưu
officecli close presentation.pptx
```
