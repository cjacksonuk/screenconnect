#Requires -RunAsAdministrator
# ==============================================================
# ICT Hero - ScreenConnect Client Installer
# icthero.co.uk
# ==============================================================
# - Reads config.txt from the script directory
# - Checks if ScreenConnect is already installed
# - Uses a pre-downloaded installer if present, otherwise downloads
# - Enforces a configurable timeout throughout
# - Logs to a configurable path (default: C:\Logs\screenconnect)
# - Auto-creates log folder if missing
# - Cleans up previous run logs, preserving the first (install) log
# ==============================================================

# Self-unblock and bypass execution policy for this session only
Unblock-File -Path $MyInvocation.MyCommand.Definition -ErrorAction SilentlyContinue
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Record script start time for timeout enforcement
$ScriptStartTime = Get-Date

# Resolve the directory this script is running from
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# ==============================================================
# LOGGING
# ==============================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $Line -Encoding UTF8
    switch ($Level) {
        "WARN"  { Write-Host $Line -ForegroundColor Yellow }
        "ERROR" { Write-Host $Line -ForegroundColor Red }
        default { Write-Host $Line -ForegroundColor Cyan }
    }
}

# ==============================================================
# TIMEOUT CHECK
# ==============================================================

function Test-Timeout {
    param([int]$TimeoutSeconds)
    $Elapsed = (Get-Date) - $ScriptStartTime
    if ($Elapsed.TotalSeconds -ge $TimeoutSeconds) {
        Write-Log "Timeout of $TimeoutSeconds seconds exceeded (elapsed: $([math]::Round($Elapsed.TotalSeconds))s). Aborting." -Level ERROR
        Exit 1
    }
}

# ==============================================================
# READ CONFIG.TXT
# ==============================================================

$ConfigFile = Join-Path $ScriptDir "config.txt"

$Config = @{
    TimeoutSeconds  = 300
    Company         = ""
    Site            = ""
    Department      = ""
    DeviceType      = ""
    Tag             = ""
    InstallerName   = "ScreenConnect.ClientSetup.exe"
    PreferMSI       = $false
    LogPath         = "C:\Logs\screenconnect"
}

if (Test-Path $ConfigFile) {
    Write-Log "Reading configuration from: $ConfigFile"
    Get-Content $ConfigFile | ForEach-Object {
        $Line = $_.Trim()
        if ($Line -eq "" -or $Line -match "^[#;]") { return }
        $Parts = $Line -split "=", 2
        if ($Parts.Count -eq 2) {
            $Key   = $Parts[0].Trim()
            $Value = $Parts[1].Trim()
            switch ($Key) {
                "TimeoutSeconds" {
                    if ($Value -match "^\d+$") { $Config.TimeoutSeconds = [int]$Value }
                    else { Write-Log "Invalid TimeoutSeconds value '$Value' - using default $($Config.TimeoutSeconds)s" -Level WARN }
                }
                "Company"       { $Config.Company       = $Value }
                "Site"          { $Config.Site           = $Value }
                "Department"    { $Config.Department     = $Value }
                "DeviceType"    { $Config.DeviceType     = $Value }
                "Tag"           { $Config.Tag            = $Value }
                "InstallerName" { $Config.InstallerName  = $Value }
                "PreferMSI"     { $Config.PreferMSI      = ($Value -eq "true") }
                "LogPath"       { $Config.LogPath        = $Value }
                default         { Write-Log "Unknown config key '$Key' - ignored" -Level WARN }
            }
        }
    }
    Write-Log "Configuration loaded. Timeout: $($Config.TimeoutSeconds)s"
} else {
    Write-Log "config.txt not found at '$ConfigFile'. Using defaults." -Level WARN
}

# ==============================================================
# LOG FOLDER SETUP AND CLEANUP
# ==============================================================

# Create log folder if it does not exist
if (-not (Test-Path $Config.LogPath)) {
    try {
        New-Item -ItemType Directory -Path $Config.LogPath -Force | Out-Null
        Write-Host "Created log folder: $($Config.LogPath)"
    } catch {
        # Fall back to script directory if we cannot create the log folder
        Write-Host "WARNING: Could not create log folder '$($Config.LogPath)': $_. Falling back to script directory." -ForegroundColor Yellow
        $Config.LogPath = $ScriptDir
    }
}

# Now we know the log folder exists - define the log file path
$LogFile = Join-Path $Config.LogPath ("ScreenConnect-Install-" + (Get-Date -Format "yyyy-MM-dd_HHmm") + ".log")

# Clean up previous run logs, keeping only the chronologically first log
# (the original install log) plus the current run
$ExistingLogs = Get-ChildItem -Path $Config.LogPath -Filter "ScreenConnect-Install-*.log" -ErrorAction SilentlyContinue |
                Sort-Object Name

if ($ExistingLogs.Count -gt 0) {
    # The oldest log is the original install record - always preserve it
    $OldestLog = $ExistingLogs[0]

    # Remove everything except the oldest log (current run log does not exist yet)
    $LogsToRemove = $ExistingLogs | Where-Object { $_.FullName -ne $OldestLog.FullName }
    foreach ($OldLog in $LogsToRemove) {
        try {
            Remove-Item $OldLog.FullName -Force -ErrorAction Stop
            Write-Host "Removed previous log: $($OldLog.Name)"
        } catch {
            Write-Host "WARNING: Could not remove old log '$($OldLog.Name)': $_" -ForegroundColor Yellow
        }
    }
    Write-Host "Preserved original install log: $($OldestLog.Name)"
}

# ==============================================================
# RESOLVE COMPANY NAME
# ==============================================================

if ([string]::IsNullOrWhiteSpace($Config.Company)) {
    Write-Log "Company not set in config.txt - attempting to determine from domain..." -Level WARN

    $Domain = $env:USERDNSDOMAIN
    if ([string]::IsNullOrWhiteSpace($Domain)) {
        try {
            $CS = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
            $Domain = $CS.Domain
        } catch {
            Write-Log "WMI query failed: $_" -Level WARN
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Domain) -and $Domain -ne "WORKGROUP") {
        $Config.Company = $Domain
        Write-Log "Using domain as Company name: $Domain"
    } else {
        Write-Log "Cannot determine domain - building fallback Company identifier..." -Level WARN
        $PublicIP = ""
        try {
            $PublicIP = (Invoke-RestMethod -Uri "https://api.ipify.org" -TimeoutSec 10 -ErrorAction Stop).Trim()
        } catch {
            $PublicIP = "UnknownIP"
            Write-Log "Could not retrieve public IP: $_" -Level WARN
        }
        $Config.Company = "$PublicIP-$env:COMPUTERNAME-$env:USERNAME"
        Write-Log "Fallback Company name: $($Config.Company)" -Level WARN
    }
}

Write-Log "Company    : $($Config.Company)"
Write-Log "Site       : $($Config.Site)"
Write-Log "Department : $($Config.Department)"
Write-Log "DeviceType : $($Config.DeviceType)"
Write-Log "Tag        : $($Config.Tag)"

# ==============================================================
# CHECK IF ALREADY INSTALLED
# ==============================================================

Test-Timeout -TimeoutSeconds $Config.TimeoutSeconds

function Get-ScreenConnectInstalled {
    $RegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($Path in $RegPaths) {
        $Entry = Get-ItemProperty $Path -ErrorAction SilentlyContinue |
                 Where-Object { $_.DisplayName -match "ScreenConnect" }
        if ($Entry) { return $true }
    }
    $Service = Get-Service -Name "ScreenConnect*" -ErrorAction SilentlyContinue
    if ($Service) { return $true }
    return $false
}

if (Get-ScreenConnectInstalled) {
    Write-Log "ScreenConnect is already installed. Nothing to do."
    Exit 0
}

Write-Log "ScreenConnect is NOT installed. Proceeding..."

# ==============================================================
# LOCATE OR DOWNLOAD INSTALLER
# ==============================================================

Test-Timeout -TimeoutSeconds $Config.TimeoutSeconds

if ($Config.PreferMSI) {
    $InstallerFilename = [System.IO.Path]::ChangeExtension($Config.InstallerName, ".msi")
    $UrlTemplate = "https://icthero.screenconnect.com/Bin/ScreenConnect.ClientSetup.msi?e=Access&y=Guest&c={0}&c={1}&c={2}&c={3}&c=&c=&c=&c={4}"
} else {
    $InstallerFilename = [System.IO.Path]::ChangeExtension($Config.InstallerName, ".exe")
    $UrlTemplate = "https://icthero.screenconnect.com/Bin/ScreenConnect.ClientSetup.exe?e=Access&y=Guest&c={0}&c={1}&c={2}&c={3}&c=&c=&c=&c={4}"
}

$LocalInstaller = Join-Path $ScriptDir $InstallerFilename

Add-Type -AssemblyName System.Web
$UrlCompany  = [System.Web.HttpUtility]::UrlEncode($Config.Company)
$UrlSite     = [System.Web.HttpUtility]::UrlEncode($Config.Site)
$UrlDept     = [System.Web.HttpUtility]::UrlEncode($Config.Department)
$UrlDevType  = [System.Web.HttpUtility]::UrlEncode($Config.DeviceType)
$UrlTag      = [System.Web.HttpUtility]::UrlEncode($Config.Tag)
$DownloadUrl = $UrlTemplate -f $UrlCompany, $UrlSite, $UrlDept, $UrlDevType, $UrlTag

if (Test-Path $LocalInstaller) {
    Write-Log "Pre-downloaded installer found: $LocalInstaller"
    Write-Log "NOTE: Pre-downloaded installer may not contain Company/Site metadata - delete it to force a fresh download if labelling matters."
} else {
    Write-Log "No pre-downloaded installer found. Downloading..."
    Write-Log "URL: $DownloadUrl"

    $Elapsed = (Get-Date) - $ScriptStartTime
    $RemainingSeconds = $Config.TimeoutSeconds - [int]$Elapsed.TotalSeconds

    if ($RemainingSeconds -le 10) {
        Write-Log "Insufficient time remaining before timeout. Aborting." -Level ERROR
        Exit 1
    }

    try {
        $DownloadJob = Start-Job -ScriptBlock {
            param($Url, $Dest)
            $wc = New-Object System.Net.WebClient
            $wc.Headers.Add("User-Agent", "ICTHero-Installer/1.0")
            $wc.DownloadFile($Url, $Dest)
        } -ArgumentList $DownloadUrl, $LocalInstaller

        $Finished = Wait-Job $DownloadJob -Timeout $RemainingSeconds

        if ($Finished -eq $null) {
            Stop-Job $DownloadJob
            Remove-Job $DownloadJob
            Write-Log "Download timed out after $RemainingSeconds seconds." -Level ERROR
            Exit 1
        }

        Receive-Job $DownloadJob -ErrorAction Stop
        Remove-Job $DownloadJob

        if (-not (Test-Path $LocalInstaller)) {
            throw "Installer file not found after download."
        }

        $FileSize = (Get-Item $LocalInstaller).Length
        if ($FileSize -lt 10000) {
            throw "Downloaded file is suspiciously small ($FileSize bytes) - likely an error response."
        }

        Write-Log "Download complete. File size: $([math]::Round($FileSize / 1KB, 1)) KB"

    } catch {
        Write-Log "Download failed: $_" -Level ERROR
        if (Test-Path $LocalInstaller) { Remove-Item $LocalInstaller -Force -ErrorAction SilentlyContinue }
        Exit 1
    }
}

# ==============================================================
# INSTALL
# ==============================================================

Test-Timeout -TimeoutSeconds $Config.TimeoutSeconds

Write-Log "Starting installation from: $LocalInstaller"

$Elapsed        = (Get-Date) - $ScriptStartTime
$InstallTimeout = $Config.TimeoutSeconds - [int]$Elapsed.TotalSeconds

if ($InstallTimeout -le 5) {
    Write-Log "Insufficient time remaining to run installer. Aborting." -Level ERROR
    Exit 1
}

try {
    if ($Config.PreferMSI) {
        $MsiArgs = @("/i", "`"$LocalInstaller`"", "/qn", "/norestart", "/l*v", "`"$LogFile.msi.log`"")
        Write-Log "Running: msiexec $($MsiArgs -join ' ')"
        $InstallProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList $MsiArgs -Wait:$false -PassThru
    } else {
        $ExeArgs = @("/silent", "/qn", "/norestart")
        Write-Log "Running: $LocalInstaller $($ExeArgs -join ' ')"
        $InstallProcess = Start-Process -FilePath $LocalInstaller -ArgumentList $ExeArgs -Wait:$false -PassThru
    }

    $Completed = $InstallProcess.WaitForExit($InstallTimeout * 1000)

    if (-not $Completed) {
        Write-Log "Installer did not complete within the timeout window. Killing process..." -Level ERROR
        $InstallProcess.Kill()
        Exit 1
    }

    $ExitCode = $InstallProcess.ExitCode
    Write-Log "Installer exited with code: $ExitCode"

    if ($ExitCode -eq 0 -or $ExitCode -eq 3010) {
        if ($ExitCode -eq 3010) {
            Write-Log "Installation succeeded but a REBOOT IS REQUIRED." -Level WARN
        } else {
            Write-Log "Installation completed successfully."
        }
    } else {
        Write-Log "Installer returned a non-zero exit code ($ExitCode). Installation may have failed." -Level ERROR
        Exit $ExitCode
    }

} catch {
    Write-Log "Exception during installation: $_" -Level ERROR
    Exit 1
}

# ==============================================================
# POST-INSTALL VERIFICATION
# ==============================================================

Start-Sleep -Seconds 5

if (Get-ScreenConnectInstalled) {
    Write-Log "Post-install check PASSED: ScreenConnect is now installed."
} else {
    Write-Log "Post-install check FAILED: ScreenConnect does not appear to be installed. Check logs." -Level WARN
}

$TotalElapsed = [math]::Round(((Get-Date) - $ScriptStartTime).TotalSeconds, 1)
Write-Log "Script completed in $TotalElapsed seconds."
Write-Log "Log saved to: $LogFile"
Exit 0
