# === CONFIGURATION ===
$DailyPath = "F:\HyperV-B\00_Daily"
$WeeklyPath = "F:\HyperV-B\01_Weekly"
$WeeksToKeep = 4
$LogFile = "E:\HyperV-Backup.log"
$Source = "Rotate-Weekly"

# Mail settings
. "C:\Script\MailConfig.ps1"
. "C:\Script\MailFunctions.ps1"
# =====================

function Log {
    param([string]$Message, [string]$Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Type] [$Source] $Message"
    switch ($Type) {
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }
    Add-Content -Path $LogFile -Value $logEntry
}

# Task header (two empty lines + title)
Add-Content -Path $LogFile -Value "`r`n"
Log "=== TASK: $Source ===" "INFO"

$script:Errors = @()
function Add-Error {
    param([string]$Message)
    $script:Errors += $Message
    Log $Message "ERROR"
}

if (-not (Test-Path $WeeklyPath)) {
    New-Item -ItemType Directory -Path $WeeklyPath -Force | Out-Null
    Log "Created weekly folder: $WeeklyPath" "INFO"
}

Log "Starting weekly rotation (keeping last $WeeksToKeep weekly archives)..." "INFO"

$latestDaily = Get-ChildItem -Path $DailyPath -Filter "*.7z" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $latestDaily) {
    Log "No daily archives found to create weekly copy." "WARNING"
    exit 0
}

$destFile = Join-Path $WeeklyPath $latestDaily.Name

if (-not (Test-Path $destFile)) {
    try {
        Copy-Item -Path $latestDaily.FullName -Destination $destFile -Force -ErrorAction Stop
        Log "Weekly copy created: $destFile" "SUCCESS"
    } catch {
        Add-Error "Failed to copy weekly archive: $_"
    }
} else {
    Log "Weekly copy for $($latestDaily.Name) already exists. Skipping." "INFO"
}

$weeklyFiles = Get-ChildItem -Path $WeeklyPath -Filter "*.7z" | Sort-Object LastWriteTime -Descending
$toDelete = $weeklyFiles | Select-Object -Skip $WeeksToKeep
if ($toDelete) {
    Log "Deleting $($toDelete.Count) old weekly archive(s):" "INFO"
    foreach ($file in $toDelete) {
        try {
            Remove-Item -Path $file.FullName -Force -ErrorAction Stop
            Log "  Removed: $($file.FullName)" "INFO"
        } catch {
            Add-Error "Failed to delete $($file.FullName): $_"
        }
    }
}
Log "Weekly rotation completed." "SUCCESS"

if ($script:Errors.Count -gt 0) {
    Send-ErrorNotification -Subject "Hyper-V Rotate Weekly ERROR on $env:COMPUTERNAME" -Body ($script:Errors -join "`r`n")
}