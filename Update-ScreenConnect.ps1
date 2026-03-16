#Requires -RunAsAdministrator
# ==============================================================
# ICT Hero - ScreenConnect Installer Cache Updater
# icthero.co.uk
# ==============================================================
# Reads config.txt from the script directory and downloads a
# fresh ScreenConnect installer to the same folder, replacing
# any existing cached copy.
# ==============================================================

Unblock-File -Path $MyInvocation.MyCommand.Definition -ErrorAction SilentlyContinue
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ConfigFile = Join-Path $ScriptDir "config.txt"

# Defaults
$Company    = ""
$Site       = ""
$Department = ""
$DeviceType = ""
$Tag        = ""
$PreferMSI  = $false
$InstallerName = "ScreenConnect.ClientSetup.exe"

# Read config.txt
if (Test-Path $ConfigFile) {
    Get-Content $ConfigFile | ForEach-Object {
        $Line = $_.Trim()
        if ($Line -eq "" -or $Line -match "^[#;]") { return }
        $Parts = $Line -split "=", 2
        if ($Parts.Count -eq 2) {
            switch ($Parts[0].Trim()) {
                "Company"       { $Company       = $Parts[1].Trim() }
                "Site"          { $Site          = $Parts[1].Trim() }
                "Department"    { $Department    = $Parts[1].Trim() }
                "DeviceType"    { $DeviceType    = $Parts[1].Trim() }
                "Tag"           { $Tag           = $Parts[1].Trim() }
                "InstallerName" { $InstallerName = $Parts[1].Trim() }
                "PreferMSI"     { $PreferMSI     = ($Parts[1].Trim() -eq "true") }
            }
        }
    }
} else {
    Write-Host "WARNING: config.txt not found - using defaults." -ForegroundColor Yellow
}

# Fall back to domain if Company is blank
if ([string]::IsNullOrWhiteSpace($Company)) {
    $Company = $env:USERDNSDOMAIN
    if ([string]::IsNullOrWhiteSpace($Company)) {
        try { $Company = (Get-WmiObject Win32_ComputerSystem).Domain } catch {}
    }
    if ([string]::IsNullOrWhiteSpace($Company) -or $Company -eq "WORKGROUP") {
        $Company = "$env:COMPUTERNAME-$env:USERNAME"
    }
    Write-Host "WARNING: Company not set in config.txt - using '$Company'" -ForegroundColor Yellow
}

# Build download URL
Add-Type -AssemblyName System.Web
$Enc = [System.Web.HttpUtility]

if ($PreferMSI) {
    $DestFile    = Join-Path $ScriptDir ([System.IO.Path]::ChangeExtension($InstallerName, ".msi"))
    $DownloadUrl = "https://icthero.screenconnect.com/Bin/ScreenConnect.ClientSetup.msi?e=Access&y=Guest&c={0}&c={1}&c={2}&c={3}&c=&c=&c=&c={4}" -f `
        $Enc::UrlEncode($Company), $Enc::UrlEncode($Site), $Enc::UrlEncode($Department), $Enc::UrlEncode($DeviceType), $Enc::UrlEncode($Tag)
} else {
    $DestFile    = Join-Path $ScriptDir ([System.IO.Path]::ChangeExtension($InstallerName, ".exe"))
    $DownloadUrl = "https://icthero.screenconnect.com/Bin/ScreenConnect.ClientSetup.exe?e=Access&y=Guest&c={0}&c={1}&c={2}&c={3}&c=&c=&c=&c={4}" -f `
        $Enc::UrlEncode($Company), $Enc::UrlEncode($Site), $Enc::UrlEncode($Department), $Enc::UrlEncode($DeviceType), $Enc::UrlEncode($Tag)
}

# Remove existing cached installer
if (Test-Path $DestFile) {
    Write-Host "Removing existing cached installer..."
    Remove-Item $DestFile -Force
}

# Download fresh installer
Write-Host "Downloading installer for: $Company"
Write-Host "Destination: $DestFile"

try {
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent", "ICTHero-CacheUpdater/1.0")
    $wc.DownloadFile($DownloadUrl, $DestFile)

    $FileSize = (Get-Item $DestFile).Length
    Write-Host "Done. File size: $([math]::Round($FileSize / 1KB, 1)) KB" -ForegroundColor Green
} catch {
    Write-Host "Download failed: $_" -ForegroundColor Red
    if (Test-Path $DestFile) { Remove-Item $DestFile -Force -ErrorAction SilentlyContinue }
    Exit 1
}
