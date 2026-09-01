# install.ps1
# Multi-mode installer for Antigravity (AGY) skills on Windows
# Supports both remote execution (irm | iex) and local cloned repo execution.

[CmdletBinding()]
param(
    [string[]]$TargetDirs = @(
        "$env:USERPROFILE\.gemini\antigravity-cli\skills",
        "$env:USERPROFILE\.gemini\config\skills",
        "$env:USERPROFILE\.agents\skills"
    )
)

$ErrorActionPreference = "Stop"

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "  ANTIGRAVITY (AGY) SKILLS INSTALLER FOR WINDOWS   " -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host ""

$skills = @("diagnose-crash", "windows", "vn-officecli")
$tempFolder = $null

# Determine source directory
$sourceSkillsDir = $null
if ($PSScriptRoot -and (Test-Path (Join-Path $PSScriptRoot "diagnose-crash\SKILL.md"))) {
    $sourceSkillsDir = $PSScriptRoot
    Write-Host "[*] Running from local repository: $sourceSkillsDir" -ForegroundColor Gray
} else {
    Write-Host "[*] Running in remote/in-memory mode. Downloading latest release..." -ForegroundColor Yellow
    $zipUrl = "https://github.com/llIIllIID0EIIllIIll/agy-skills/archive/refs/heads/main.zip"
    $tempZip = Join-Path $env:TEMP "agy-skills-$([Guid]::NewGuid().ToString('N')).zip"
    $tempFolder = Join-Path $env:TEMP "agy-skills-$([Guid]::NewGuid().ToString('N'))"
    
    Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip -UseBasicParsing
    Expand-Archive -Path $tempZip -DestinationPath $tempFolder -Force
    Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue

    $extractedRoot = Join-Path $tempFolder "agy-skills-main"
    if (Test-Path (Join-Path $extractedRoot "diagnose-crash\SKILL.md")) {
        $sourceSkillsDir = $extractedRoot
    } elseif (Test-Path (Join-Path $extractedRoot "skills\diagnose-crash\SKILL.md")) {
        $sourceSkillsDir = Join-Path $extractedRoot "skills"
    } else {
        throw "Could not locate skills in downloaded archive."
    }
}

try {
    foreach ($target in $TargetDirs) {
        $resolvedDir = [System.IO.Path]::GetFullPath($target)
        if (!(Test-Path $resolvedDir)) {
            Write-Host "[+] Creating target directory: $resolvedDir" -ForegroundColor Yellow
            New-Item -Path $resolvedDir -ItemType Directory -Force | Out-Null
        } else {
            Write-Host "[*] Target directory found: $resolvedDir" -ForegroundColor Gray
        }

        foreach ($skill in $skills) {
            $src = Join-Path $sourceSkillsDir $skill
            $dst = Join-Path $resolvedDir $skill
            
            if (Test-Path $src) {
                Write-Host "    [+] Installing skill: '$skill' -> $dst" -ForegroundColor Green
                if (Test-Path $dst) {
                    Remove-Item -Path $dst -Recurse -Force
                }
                Copy-Item -Path $src -Destination $dst -Recurse -Force
                
                $skillMd = Join-Path $dst "SKILL.md"
                if (Test-Path $skillMd) {
                    Write-Host "        ✔ SKILL.md validated" -ForegroundColor DarkGreen
                }
            } else {
                Write-Host "    [-] Source skill directory not found: $src" -ForegroundColor Red
            }
        }
    }

    Write-Host ""
    Write-Host "===================================================" -ForegroundColor Cyan
    Write-Host "  INSTALLATION COMPLETED SUCCESSFULLY!" -ForegroundColor Green
    Write-Host "===================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Installed skills:" -ForegroundColor White
    Write-Host "  1. diagnose-crash : Diagnose Windows app crashes, WER, minidumps, event logs" -ForegroundColor Cyan
    Write-Host "  2. windows        : Windows 11 desktop customization, terminal, winget, theming" -ForegroundColor Cyan
    Write-Host "  3. vn-officecli   : Soan thao To Trinh, Van ban hanh chinh VN & bo Office (.docx, .xlsx, .pptx)" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Run 'agy' in your terminal and type '/skills' to verify." -ForegroundColor Yellow
    Write-Host ""
}
finally {
    if ($tempFolder -and (Test-Path $tempFolder)) {
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}
