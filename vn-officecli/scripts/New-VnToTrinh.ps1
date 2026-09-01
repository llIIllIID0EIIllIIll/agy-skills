<#
.SYNOPSIS
    Tạo văn bản Tờ trình chuẩn thể thức hành chính Việt Nam (.docx) bằng OfficeCLI.

.DESCRIPTION
    Tạo tài liệu Word tuân thủ Nghị định 30/2020/NĐ-CP và tiêu chuẩn quy chế doanh nghiệp v3.0:
    - Khổ giấy A4 (210 x 297mm)
    - Căn lề: Trên 25mm, Dưới 20mm, Trái 30mm, Phải 15mm
    - Font Times New Roman, cỡ chữ 13pt, giãn dòng 1.3x, thụt đầu dòng 1.25cm
    - Header và Chân trang bảng 2 cột ẩn viền (chống lệch trang)
    - Đầy đủ 3 phần chuẩn của Tờ trình: Sự cần thiết -> Nội dung đề xuất -> Kiến nghị phê duyệt

.PARAMETER OutputPath
    Đường dẫn file .docx đầu ra (Mặc định: ToTrinh.docx)

.PARAMETER Title
    Trích yếu nội dung tờ trình (Ví dụ: "V/v phê duyệt kế hoạch chuyển đổi số năm 2026")

.PARAMETER Recipient
    Kính gửi (Ví dụ: "Hội đồng Quản trị / Ban Tổng Giám đốc")

.PARAMETER Organization
    Tên đơn vị soạn thảo / ban hành (Ví dụ: "PHÒNG CÔNG NGHỆ THÔNG TIN")

.PARAMETER ParentOrg
    Cơ quan / Công ty chủ quản (Ví dụ: "CÔNG TY CỔ PHẦN TẬP ĐOÀN ...")

.PARAMETER DocNumber
    Số ký hiệu văn bản (Ví dụ: "01/TTr-CNTT")

.PARAMETER Location
    Địa danh (Ví dụ: "Hà Nội" hoặc "TP. Hồ Chí Minh")

.PARAMETER SignerTitle
    Chức danh người ký (Ví dụ: "TRƯỞNG PHÒNG")

.PARAMETER SignerName
    Họ tên người ký (Ví dụ: "Nguyễn Văn A")
#>

[CmdletBinding()]
param(
    [string]$OutputPath = "ToTrinh_Chuan.docx",
    [string]$Title = "V/v phê duyệt chủ trương và kế hoạch triển khai dự án năm 2026",
    [string]$Recipient = "Hội đồng Quản trị / Ban Tổng Giám đốc Công ty",
    [string]$Organization = "PHÒNG KẾ HOẠCH & CHIẾN LƯỢC",
    [string]$ParentOrg = "CÔNG TY CỔ PHẦN TẬP ĐOÀN ABC",
    [string]$DocNumber = "01/TTr-KHCL",
    [string]$Location = "Hà Nội",
    [string]$SignerTitle = "TRƯỞNG PHÒNG",
    [string]$SignerName = "Nguyễn Văn A",
    [string[]]$Receivers = @("Như kính gửi;", "Ban Tổng Giám đốc;", "Phòng Tài chính - Kế toán;", "Lưu: VT, KHCL.")
)

$ErrorActionPreference = "Stop"

# Kiểm tra officecli
$cli = Get-Command officecli -ErrorAction SilentlyContinue
if (-not $cli) {
    $alt = "$env:LOCALAPPDATA\OfficeCLI\officecli.exe"
    if (Test-Path $alt) {
        $cliPath = $alt
    } else {
        throw "OfficeCLI chưa được cài đặt. Vui lòng chạy lệnh: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCli/main/install.ps1 | iex"
    }
} else {
    $cliPath = $cli.Source
}

$fullPath = [System.IO.Path]::GetFullPath($OutputPath)
if (Test-Path $fullPath) {
    & $cliPath close $fullPath 2>$null
    Remove-Item $fullPath -Force
}

Write-Host "[*] Khởi tạo tài liệu: $fullPath" -ForegroundColor Cyan
& $cliPath create $fullPath

# 1. Thiết lập khổ giấy A4, Lề và Font mặc định
Write-Host "[+] Thiết lập quy chuẩn khổ giấy và lề A4..." -ForegroundColor Green
& $cliPath set $fullPath / `
    --prop defaultFont="Times New Roman" `
    --prop defaultFontSize="13pt" `
    --prop pageWidth="11906" `
    --prop pageHeight="16838" `
    --prop marginTop="1417" `
    --prop marginBottom="1134" `
    --prop marginLeft="1701" `
    --prop marginRight="850"

# 2. Tạo Header Table (2 cột ẩn viền)
Write-Host "[+] Thêm bảng tiêu đề Quốc hiệu & Đơn vị ban hành..." -ForegroundColor Green
& $cliPath add $fullPath /body --type table --prop cols=2 --prop rows=1 --prop border.all=none --prop width="100%"

$currentDate = Get-Date
$dateString = "$Location, ngày $($currentDate.Day.ToString('00')) tháng $($currentDate.Month.ToString('00')) năm $($currentDate.Year)"

# Cột 1 (Trái): Đơn vị ban hành & Số hiệu
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[1]" --type p --prop text="$ParentOrg" --prop alignment=center --prop size="12pt" --prop lineSpacing="1.15x"
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[1]" --type p --prop text="$Organization" --prop alignment=center --prop bold=true --prop size="12pt" --prop lineSpacing="1.15x"
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[1]" --type p --prop text="Số: $DocNumber" --prop alignment=center --prop size="12pt" --prop lineSpacing="1.15x"

# Cột 2 (Phải): Quốc hiệu, Tiêu ngữ & Ngày tháng
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[2]" --type p --prop text="CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" --prop alignment=center --prop bold=true --prop size="12pt" --prop lineSpacing="1.15x"
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[2]" --type p --prop text="Độc lập - Tự do - Hạnh phúc" --prop alignment=center --prop bold=true --prop size="13pt" --prop lineSpacing="1.15x"
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[2]" --type p --prop text="$dateString" --prop alignment=center --prop italic=true --prop size="13pt" --prop lineSpacing="1.15x"

# 3. Tiêu đề Tờ trình & Trích yếu nội dung
Write-Host "[+] Thêm tiêu đề TỜ TRÌNH và Trích yếu..." -ForegroundColor Green
& $cliPath add $fullPath /body --type p --prop text="TỜ TRÌNH" --prop alignment=center --prop bold=true --prop size="16pt" --prop spaceBefore="18pt" --prop spaceAfter="6pt"
& $cliPath add $fullPath /body --type p --prop text="$Title" --prop alignment=center --prop bold=true --prop italic=true --prop size="13pt" --prop spaceAfter="14pt"

# 4. Kính gửi
& $cliPath add $fullPath /body --type p --prop text="Kính gửi: $Recipient" --prop bold=true --prop size="13pt" --prop spaceAfter="12pt"

# 5. Nội dung Tờ trình (3 Phần chuẩn)
Write-Host "[+] Khởi tạo cấu trúc 3 phần chuẩn của Tờ trình..." -ForegroundColor Green

# Căn cứ ban đầu
& $cliPath add $fullPath /body --type p `
    --prop text="Căn cứ vào định hướng chiến lược phát triển và tình hình hoạt động sản xuất kinh doanh thực tế của Công ty; căn cứ vào chức năng, nhiệm vụ được giao, $Organization kính trình $Recipient xem xét và phê duyệt nội dung sau:" `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"

# Phần I
& $cliPath add $fullPath /body --type p --prop text="I. SỰ CẦN THIẾT VÀ CĂN CỨ THỰC TIỄN" --prop bold=true --prop size="13pt" --prop spaceBefore="10pt" --prop spaceAfter="4pt"
& $cliPath add $fullPath /body --type p `
    --prop text="1. Bối cảnh và lý do đề xuất: Nêu rõ thực trạng hiện tại, những khó khăn bất cập hoặc cơ hội mới đòi hỏi phải thực hiện dự án/chính sách mới." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"
& $cliPath add $fullPath /body --type p `
    --prop text="2. Mục tiêu hướng tới: Xác định rõ mục tiêu định lượng và định tính nhằm giải quyết triệt để các vấn đề nêu trên, nâng cao hiệu quả vận hành của đơn vị." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"

# Phần II
& $cliPath add $fullPath /body --type p --prop text="II. NỘI DUNG ĐỀ XUẤT VÀ PHƯƠNG ÁN TRIỂN KHAI" --prop bold=true --prop size="13pt" --prop spaceBefore="10pt" --prop spaceAfter="4pt"
& $cliPath add $fullPath /body --type p `
    --prop text="1. Nội dung chi tiết phương án: Trình bày giải pháp cụ thể, quy trình áp dụng và phạm vi triển khai trên toàn hệ thống." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"
& $cliPath add $fullPath /body --type p `
    --prop text="2. Dự toán ngân sách và nguồn lực: Bảng phân bổ chi phí, nhân sự tham gia và lộ trình giải ngân từng giai đoạn." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"
& $cliPath add $fullPath /body --type p `
    --prop text="3. Tiến độ thực hiện: Kế hoạch mốc thời gian từ khâu chuẩn bị, nghiệm thu đến vận hành chính thức." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"

# Phần III
& $cliPath add $fullPath /body --type p --prop text="III. KIẾN NGHỊ VÀ ĐỀ XUẤT PHÊ DUYỆT" --prop bold=true --prop size="13pt" --prop spaceBefore="10pt" --prop spaceAfter="4pt"
& $cliPath add $fullPath /body --type p `
    --prop text="Để đảm bảo tiến độ và hiệu quả công việc, $Organization kính trình $Recipient xem xét và quyết định:" `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"
& $cliPath add $fullPath /body --type p `
    --prop text="1. Phê duyệt chủ trương và kế hoạch triển khai chi tiết theo phương án nêu tại Phần II." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="4pt"
& $cliPath add $fullPath /body --type p `
    --prop text="2. Phê duyệt dự toán ngân sách thực hiện và giao các bộ phận liên quan phối hợp triển khai." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"
& $cliPath add $fullPath /body --type p `
    --prop text="Kính trình Lãnh đạo xem xét, phê duyệt./." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="4pt" --prop spaceAfter="18pt"

# 6. Chân trang Bảng Chữ ký & Nơi nhận (Bảng 2 cột ẩn viền)
Write-Host "[+] Thêm bảng chữ ký và nơi nhận..." -ForegroundColor Green
& $cliPath add $fullPath /body --type table --prop cols=2 --prop rows=1 --prop border.all=none --prop width="100%"

# Cột 1 (Trái): Nơi nhận
& $cliPath add $fullPath "/body/tbl[2]/tr[1]/tc[1]" --type p --prop text="Nơi nhận:" --prop bold=true --prop italic=true --prop size="11pt" --prop spaceAfter="2pt"
foreach ($rcv in $Receivers) {
    & $cliPath add $fullPath "/body/tbl[2]/tr[1]/tc[1]" --type p --prop text="- $rcv" --prop size="10pt" --prop lineSpacing="1.15x" --prop spaceAfter="1pt"
}

# Cột 2 (Phải): Chức danh người ký & Họ tên
& $cliPath add $fullPath "/body/tbl[2]/tr[1]/tc[2]" --type p --prop text="$SignerTitle" --prop alignment=center --prop bold=true --prop size="13pt" --prop spaceAfter="40pt"
& $cliPath add $fullPath "/body/tbl[2]/tr[1]/tc[2]" --type p --prop text="$SignerName" --prop alignment=center --prop bold=true --prop size="13pt"

& $cliPath close $fullPath
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "  TẠO THÀNH CÔNG VĂN BẢN TỜ TRÌNH CHUẨN THỂ THỨC!" -ForegroundColor Green
Write-Host "  File: $fullPath" -ForegroundColor White
Write-Host "===================================================" -ForegroundColor Cyan
