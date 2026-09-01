<#
.SYNOPSIS
    Chuẩn hóa định dạng file Word (.docx) sang thể thức hành chính Việt Nam bằng OfficeCLI.

.DESCRIPTION
    Tự động hiệu chỉnh:
    - Khổ giấy A4 (210 x 297mm)
    - Lề: Trên 25mm, Dưới 20mm, Trái 30mm, Phải 15mm
    - Font Times New Roman 13pt
    - Giãn dòng 1.3x, căn đều 2 bên (Justified), thụt đầu dòng 1.25cm
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$FilePath
)

$ErrorActionPreference = "Stop"

$cli = Get-Command officecli -ErrorAction SilentlyContinue
if (-not $cli) {
    $cliPath = "$env:LOCALAPPDATA\OfficeCLI\officecli.exe"
} else {
    $cliPath = $cli.Source
}

if (!(Test-Path $cliPath)) {
    throw "Không tìm thấy officecli. Vui lòng cài đặt bằng lệnh: irm https://raw.githubusercontent.com/iOfficeAI/OfficeCli/main/install.ps1 | iex"
}

$fullPath = [System.IO.Path]::GetFullPath($FilePath)
if (!(Test-Path $fullPath)) {
    throw "Không tìm thấy file: $fullPath"
}

Write-Host "[*] Đang chuẩn hóa văn bản: $fullPath..." -ForegroundColor Cyan

# 1. Cập nhật khổ giấy và căn lề A4
& $cliPath set $fullPath / `
    --prop defaultFont="Times New Roman" `
    --prop defaultFontSize="13pt" `
    --prop pageWidth="11906" `
    --prop pageHeight="16838" `
    --prop marginTop="1417" `
    --prop marginBottom="1134" `
    --prop marginLeft="1701" `
    --prop marginRight="850"

# 2. Lưu và đóng
& $cliPath close $fullPath
Write-Host "[✔] Hoàn tất chuẩn hóa thể thức file: $fullPath" -ForegroundColor Green
