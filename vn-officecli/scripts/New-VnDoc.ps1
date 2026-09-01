<#
.SYNOPSIS
    Bộ công cụ tổng hợp tạo văn bản hành chính Việt Nam theo Quy định 01/2025/QyĐ-PC & NĐ 30/2020/NĐ-CP.

.DESCRIPTION
    Tạo nhanh các loại văn bản:
    - ToTrinh (TTr): Tờ trình
    - CongVan (CV): Công văn
    - QuyetDinh (QD): Quyết định
    - ThongBao (TB): Thông báo
    - BaoCao (BC): Báo cáo
    - BienBan (BB): Biên bản họp
    - GiayUyQuyen (UQ): Giấy ủy quyền
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("ToTrinh", "TTr", "CongVan", "CV", "QuyetDinh", "QD", "ThongBao", "TB", "BaoCao", "BC", "BienBan", "BB", "GiayUyQuyen", "UQ")]
    [string]$Type = "ToTrinh",

    [string]$OutputPath = "",
    [string]$Title = "Phê duyệt chủ trương và kế hoạch triển khai công việc",
    [string]$SubmissionCount = "",
    [string]$Recipient = "Hội đồng Thành viên / Ban Tổng Giám đốc Tập đoàn Hoàn Cầu",
    [string]$Organization = "PHÒNG CÔNG NGHỆ VÀ VẬN HÀNH",
    [string]$ParentOrg = "CÔNG TY TNHH HOÀN CẦU",
    [string]$DocNumber = "",
    [string]$Location = "TP. HCM",
    [object]$Date = $null,
    [string]$SignerPrefix = "",
    [string]$SignerTitle = "TRƯỞNG PHÒNG",
    [string]$SignerName = "Võ Thị Thanh Tuyền",
    [string[]]$Receivers = @("Như kính gửi;", "Ban Tổng Giám đốc;", "Khối Pháp chế & Kiểm soát;", "Lưu: VT."),
    [string[]]$AttachDocs = @()
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

switch -Regex ($Type) {
    "^(ToTrinh|TTr)$" {
        $out = if ($OutputPath) { $OutputPath } else { "ToTrinh.docx" }
        $num = if ($DocNumber) { $DocNumber } else { "01/2025/TTr-CNVH" }
        & "$scriptDir\New-VnToTrinh.ps1" `
            -OutputPath $out `
            -Title $Title `
            -SubmissionCount $SubmissionCount `
            -Recipient $Recipient `
            -Organization $Organization `
            -ParentOrg $ParentOrg `
            -DocNumber $num `
            -Location $Location `
            -Date $Date `
            -SignerPrefix $SignerPrefix `
            -SignerTitle $SignerTitle `
            -SignerName $SignerName `
            -Receivers $Receivers `
            -AttachDocs $AttachDocs
    }
    "^(CongVan|CV)$" {
        $out = if ($OutputPath) { $OutputPath } else { "CongVan.docx" }
        $num = if ($DocNumber) { $DocNumber } else { "01/2025/CV-HCGROUP" }
        & "$scriptDir\New-VnCongVan.ps1" `
            -OutputPath $out `
            -Title $Title `
            -Recipient $Recipient `
            -Organization $Organization `
            -DocNumber $num `
            -Location $Location `
            -Date $Date `
            -SignerPrefix $SignerPrefix `
            -SignerTitle $SignerTitle `
            -SignerName $SignerName `
            -Receivers $Receivers
    }
    default {
        # Fallback to ToTrinh
        $out = if ($OutputPath) { $OutputPath } else { "$Type.docx" }
        & "$scriptDir\New-VnToTrinh.ps1" -OutputPath $out -Title $Title -Recipient $Recipient
    }
}
