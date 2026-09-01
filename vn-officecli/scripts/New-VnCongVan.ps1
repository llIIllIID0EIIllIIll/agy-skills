<#
.SYNOPSIS
    Tạo văn bản Công văn (CV) chuẩn thể thức doanh nghiệp Hoàn Cầu theo Quy định 01/2025/QyĐ-PC.

.DESCRIPTION
    Tạo văn bản Công văn tuân thủ Mẫu số 01 (Phụ lục 05 - Trang 25):
    - Khổ A4, lề chuẩn (Trên 25mm, Dưới 20mm, Trái 30mm, Phải 15mm)
    - Font Times New Roman 13pt, giãn dòng 1.3x
    - Trích yếu V/v đặt ngay dưới Số/Ký hiệu
    - Header và Chân trang bảng ẩn viền chống vỡ trang
#>

[CmdletBinding()]
param(
    [string]$OutputPath = "CongVan_HoanCau_Chuan.docx",
    [string]$Title = "V/v phối hợp triển khai công tác chuyển đổi số và chuẩn hóa văn bản",
    [string]$Recipient = "CÁC CÔNG TY THÀNH VIÊN VÀ CÁC ĐƠN VỊ TRỰC THUỘC",
    [string]$Organization = "TẬP ĐOÀN HOÀN CẦU",
    [string]$DocNumber = "01/2025/CV-HCGROUP",
    [string]$Location = "Tp. Hồ Chí Minh",
    [string]$SignerTitle = "TỔNG GIÁM ĐỐC",
    [string]$SignerName = "Nguyễn Thị Thanh Đào",
    [string[]]$Receivers = @("Như kính gửi;", "Ban Tổng Giám đốc;", "Khối Pháp chế & Vận hành;", "Lưu: VT, PC.")
)

$ErrorActionPreference = "Stop"

$cli = Get-Command officecli -ErrorAction SilentlyContinue
if (-not $cli) {
    $cliPath = "$env:LOCALAPPDATA\OfficeCLI\officecli.exe"
    if (-not (Test-Path $cliPath)) { $cliPath = "$env:LOCALAPPDATA\agy\bin\officecli.exe" }
} else {
    $cliPath = $cli.Source
}

$fullPath = [System.IO.Path]::GetFullPath($OutputPath)
$parentDir = Split-Path -Parent $fullPath
if ($parentDir -and !(Test-Path $parentDir)) {
    New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
}
if (Test-Path $fullPath) {
    & $cliPath close $fullPath 2>$null
    Remove-Item $fullPath -Force
}

Write-Host "[*] Khởi tạo tài liệu Công văn: $fullPath" -ForegroundColor Cyan
& $cliPath create $fullPath

# 1. Thiết lập A4 và căn lề
& $cliPath set $fullPath / `
    --prop defaultFont="Times New Roman" `
    --prop defaultFontSize="13pt" `
    --prop pageWidth="11906" `
    --prop pageHeight="16838" `
    --prop marginTop="1417" `
    --prop marginBottom="1134" `
    --prop marginLeft="1701" `
    --prop marginRight="850"

# 2. Header Table (2 cột ẩn viền)
$currentDate = Get-Date
$dateString = "$Location, ngày $($currentDate.Day.ToString('00')) tháng $($currentDate.Month.ToString('00')) năm $($currentDate.Year)"

& $cliPath add $fullPath /body --type table --prop cols=2 --prop rows=1 --prop border.all=none --prop width="100%"

# Cột 1 (Trái): Tên Tập đoàn, Số CV, Trích yếu
$subText = if ($Title.StartsWith("V/v:")) { $Title } else { "V/v: $Title" }
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[1]" --type p --prop text="$Organization" --prop alignment=center --prop bold=true --prop size="12pt"
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[1]" --type p --prop text="Số: $DocNumber" --prop alignment=center --prop size="12pt"
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[1]" --type p --prop text="$subText" --prop alignment=center --prop italic=true --prop size="12pt"

# Cột 2 (Phải): Quốc hiệu, Tiêu ngữ, Ngày tháng
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[2]" --type p --prop text="CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" --prop alignment=center --prop bold=true --prop size="12pt"
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[2]" --type p --prop text="Độc lập - Tự do - Hạnh phúc" --prop alignment=center --prop bold=true --prop size="13pt"
& $cliPath add $fullPath "/body/tbl[1]/tr[1]/tc[2]" --type p --prop text="$dateString" --prop alignment=center --prop italic=true --prop size="13pt"

# 3. Kính gửi
& $cliPath add $fullPath /body --type p --prop text="Kính gửi: $Recipient" --prop bold=true --prop size="13pt" --prop spaceBefore="18pt" --prop spaceAfter="12pt"

# 4. Nội dung công văn
& $cliPath add $fullPath /body --type p `
    --prop text="- Căn cứ Quy định số 01/2025/QyĐ-PC ngày 01/10/2025 của Tập đoàn Hoàn Cầu về Thể thức và Kỹ thuật trình bày văn bản;" `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="2pt" --prop spaceAfter="4pt"
& $cliPath add $fullPath /body --type p `
    --prop text="- Nhằm đảm bảo tính đồng bộ, chuyên nghiệp và nâng cao hiệu quả công tác văn thư lưu trữ trên toàn hệ thống Tập đoàn và các đơn vị thành viên;" `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="2pt" --prop spaceAfter="6pt"

& $cliPath add $fullPath /body --type p `
    --prop text="Tổng Giám đốc Tập đoàn Hoàn Cầu đề nghị Thủ trưởng các Đơn vị trực thuộc và Giám đốc các Công ty thành viên tập trung triển khai thực hiện các nội dung sau:" `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="4pt" --prop spaceAfter="6pt"

& $cliPath add $fullPath /body --type p `
    --prop text="1. Tổ chức phổ biến, quán triệt toàn diện nội dung Quy định thể thức văn bản số 01/2025/QyĐ-PC đến toàn thể cán bộ, nhân viên trong đơn vị." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="4pt"

& $cliPath add $fullPath /body --type p `
    --prop text="2. Nghiêm túc áp dụng chuẩn lề A4 (Trên 25mm, Dưới 20mm, Trái 30mm, Phải 15mm), font Times New Roman cỡ 13pt và hệ thống mã viết tắt 33 loại văn bản theo đúng Phụ lục 01." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="4pt"

& $cliPath add $fullPath /body --type p `
    --prop text="3. Giao Khối Công nghệ & Vận hành phối hợp cùng Ban Pháp chế thường xuyên kiểm tra, rà soát và hỗ trợ kỹ thuật định dạng văn bản tự động cho các phòng ban." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="3pt" --prop spaceAfter="6pt"

& $cliPath add $fullPath /body --type p `
    --prop text="Đề nghị các Đơn vị nghiêm túc phối hợp và triển khai thực hiện./." `
    --prop alignment=both --prop size="13pt" --prop lineSpacing="1.3x" --prop firstLineIndent="709" --prop spaceBefore="4pt" --prop spaceAfter="20pt"

# 5. Chân trang Nơi nhận & Ký duyệt
& $cliPath add $fullPath /body --type table --prop cols=2 --prop rows=1 --prop border.all=none --prop width="100%"
& $cliPath add $fullPath "/body/tbl[2]/tr[1]/tc[1]" --type p --prop text="Nơi nhận:" --prop bold=true --prop italic=true --prop size="10pt" --prop spaceAfter="2pt"
foreach ($rcv in $Receivers) {
    & $cliPath add $fullPath "/body/tbl[2]/tr[1]/tc[1]" --type p --prop text="- $rcv" --prop size="10pt" --prop lineSpacing="1.15x" --prop spaceAfter="1pt"
}

& $cliPath add $fullPath "/body/tbl[2]/tr[1]/tc[2]" --type p --prop text="$SignerTitle" --prop alignment=center --prop bold=true --prop size="13pt" --prop spaceAfter="45pt"
& $cliPath add $fullPath "/body/tbl[2]/tr[1]/tc[2]" --type p --prop text="$SignerName" --prop alignment=center --prop bold=true --prop size="13pt"

& $cliPath close $fullPath
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "  TẠO THÀNH CÔNG CÔNG VĂN CHUẨN HOÀN CẦU!" -ForegroundColor Green
Write-Host "  File: $fullPath" -ForegroundColor White
Write-Host "===================================================" -ForegroundColor Cyan
