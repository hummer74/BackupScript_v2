# === CONFIGURATION ===
$WeeklyPath = "F:\HyperV-B\01_Weekly"
$MonthlyPath = "F:\HyperV-B\02_Monthly"
$MonthsToKeep = 6
$LogFile = "E:\HyperV-Backup.log"
$Source = "Rotate-Monthly"

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

if (-not (Test-Path $MonthlyPath)) {
    New-Item -ItemType Directory -Path $MonthlyPath -Force | Out-Null
    Log "Created monthly folder: $MonthlyPath" "INFO"
}

Log "Starting monthly rotation (keeping last $MonthsToKeep monthly archives)..." "INFO"

$latestWeekly = Get-ChildItem -Path $WeeklyPath -Filter "*.7z" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $latestWeekly) {
    Log "No weekly archives found to create monthly copy." "WARNING"
    exit 0
}

$destFile = Join-Path $MonthlyPath $latestWeekly.Name

if (-not (Test-Path $destFile)) {
    try {
        Copy-Item -Path $latestWeekly.FullName -Destination $destFile -Force -ErrorAction Stop
        Log "Monthly copy created: $destFile" "SUCCESS"
    } catch {
        Add-Error "Failed to copy monthly archive: $_"
    }
} else {
    Log "Monthly copy for $($latestWeekly.Name) already exists. Skipping." "INFO"
}

$monthlyFiles = Get-ChildItem -Path $MonthlyPath -Filter "*.7z" | Sort-Object LastWriteTime -Descending
$toDelete = $monthlyFiles | Select-Object -Skip $MonthsToKeep
if ($toDelete) {
    Log "Deleting $($toDelete.Count) old monthly archive(s):" "INFO"
    foreach ($file in $toDelete) {
        try {
            Remove-Item -Path $file.FullName -Force -ErrorAction Stop
            Log "  Removed: $($file.FullName)" "INFO"
        } catch {
            Add-Error "Failed to delete $($file.FullName): $_"
        }
    }
}
Log "Monthly rotation completed." "SUCCESS"

if ($script:Errors.Count -gt 0) {
    Send-ErrorNotification -Subject "Hyper-V Rotate Monthly ERROR on $env:COMPUTERNAME" -Body ($script:Errors -join "`r`n")
}