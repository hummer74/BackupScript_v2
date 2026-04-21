# === CONFIGURATION ===
# All paths and settings are centralized here
$VMName = "1c-azimut"
$BackupRoot = "F:\HyperV-B"
$DailyPath = Join-Path $BackupRoot "00_Daily"
$TempRoot = "E:\Z-TEMP\HyperV-Backup"
$SevenZipPath = "C:\Program Files\7-Zip\7z.exe"
$LogFile = "E:\HyperV-Backup.log"
$ScriptSource = "Backup-and-Rotate"
$DaysToKeep = 5

# External dependencies for mail notifications
. "C:\Script\MailConfig.ps1"
. "C:\Script\MailFunctions.ps1"
# =====================

function Log {
    param([string]$Message, [string]$Type = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Type] [$ScriptSource] $Message"
    switch ($Type) {
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }
    Add-Content -Path $LogFile -Value $logEntry
}

# Task header in the log file
Add-Content -Path $LogFile -Value "`r`n"
Log "=== STARTING COMPREHENSIVE TASK: $ScriptSource ===" "INFO"

$script:Errors = @()
function Add-Error {
    param([string]$Message)
    $script:Errors += $Message
    Log $Message "ERROR"
}

function Test-DiskSpace {
    param(
        [string]$Path,
        [long]$RequiredBytes,
        [string]$Label = "Destination"
    )
    # Get the root drive even if the folder path doesn't exist yet
    try {
        $root = [System.IO.Path]::GetPathRoot($Path)
        $drive = Get-PSDrive -Name $root.TrimEnd(':\') -ErrorAction Stop
        
        $free = $drive.Free
        if ($free -lt $RequiredBytes) {
            $freeGB = [math]::Round($free/1GB, 2)
            $reqGB = [math]::Round($RequiredBytes/1GB, 2)
            return $false, "Insufficient space for $Label (Drive $($drive.Name): at $Path). Free: $freeGB GB, Need: $reqGB GB"
        }
        return $true, $null
    }
    catch {
        return $false, "Could not access drive for $Label (Path: $Path). Error: $_"
    }
}

function Wait-VMState {
    param(
        [string]$VMName,
        [string]$TargetState,
        [int]$TimeoutSeconds = 60,
        [int]$CheckIntervalSeconds = 2
    )
    $elapsed = 0
    while ($elapsed -lt $TimeoutSeconds) {
        $state = (Get-VM -Name $VMName).State
        if ($state -eq $TargetState) { return $true }
        Start-Sleep -Seconds $CheckIntervalSeconds
        $elapsed += $CheckIntervalSeconds
    }
    return $false
}

# --- Phase 1: Initialization & Global Checks ---
$DateStr = Get-Date -Format "yyyy-MM-dd_HH-mm"
$TempExportFolder = Join-Path $TempRoot "$VMName-$DateStr"
$ArchiveFileName = Join-Path $DailyPath "$VMName-$DateStr.7z"

$earlyExit = $false

# 1. Check if 7-Zip exists
if (-not (Test-Path $SevenZipPath)) {
    Add-Error "7-Zip not found at $SevenZipPath."
    $earlyExit = $true
}

# 2. Check if VM exists
$VM = Get-VM -Name $VMName -ErrorAction SilentlyContinue
if (-not $VM) {
    Add-Error "Virtual Machine '$VMName' not found on this host."
    $earlyExit = $true
}

# 3. Ensure all base directories exist (Create them if missing)
if (-not $earlyExit) {
    $PathsToVerify = @($BackupRoot, $DailyPath, $TempRoot)
    foreach ($P in $PathsToVerify) {
        if (-not (Test-Path $P)) {
            try {
                Log "Creating missing directory: $P" "INFO"
                New-Item -ItemType Directory -Path $P -Force -ErrorAction Stop | Out-Null
            } catch {
                Add-Error "Failed to create directory ${P}: $($_.Exception.Message)"
                $earlyExit = $true
            }
        }
    }
}

# 4. Space Checks
if (-not $earlyExit) {
    $vhdSize = 0
    Get-VMHardDiskDrive -VMName $VMName | ForEach-Object {
        if (Test-Path $_.Path) { $vhdSize += (Get-Item $_.Path).Length }
    }
    $requiredTemp = [math]::Round($vhdSize * 1.1)

    $spaceOk, $spaceMsg = Test-DiskSpace -Path $TempRoot -RequiredBytes $requiredTemp -Label "Temporary Export Folder"
    if (-not $spaceOk) { Add-Error $spaceMsg; $earlyExit = $true }

    $spaceOk, $spaceMsg = Test-DiskSpace -Path $BackupRoot -RequiredBytes $vhdSize -Label "Final Backup Storage"
    if (-not $spaceOk) { Add-Error $spaceMsg; $earlyExit = $true }
}

# Stop if early checks failed
if ($earlyExit) {
    if ($script:Errors.Count -gt 0) {
        Send-ErrorNotification -Subject "Hyper-V Backup ERROR: $VMName" -Body ($script:Errors -join "`r`n")
    }
    exit 1
}

# Cleanup temporary root before export
if (Test-Path $TempRoot) {
    try {
        Log "Cleaning up temporary root directory: $TempRoot" "INFO"
        Get-ChildItem -Path $TempRoot -Recurse -Force | Remove-Item -Recurse -Force -ErrorAction Stop
        Log "Temporary root cleaned successfully." "SUCCESS"
    }
    catch {
        Log "Warning: Could not fully clean temporary root. Continuing anyway. Error: $_" "WARNING"
    }
}

# --- Phase 2: Execution Block ---
$WasRunning = ($VM.State -eq 'Running')

try {
    # Step 1: Stop VM
    if ($WasRunning) {
        Log "Attempting graceful shutdown of '$VMName' (Timeout: 600s)..."
        Stop-VM -VM $VM -Confirm:$false
        if (-not (Wait-VMState -VMName $VMName -TargetState 'Off' -TimeoutSeconds 600)) {
            Log "Graceful shutdown timed out. Attempting 'Turn Off' (Force)..." "WARNING"
            Stop-VM -VM $VM -Force -Confirm:$false
            if (-not (Wait-VMState -VMName $VMName -TargetState 'Off' -TimeoutSeconds 60)) {
                throw "CRITICAL: VM '$VMName' failed to stop. Aborting backup to prevent corruption."
            }
        }
        Log "VM is stopped." "SUCCESS"
    }

    # Step 2: Export VM
    if (-not (Test-Path $TempExportFolder)) { New-Item -ItemType Directory -Path $TempExportFolder -Force | Out-Null }
    Log "Exporting VM to $TempExportFolder..."
    Export-VM -Name $VMName -Path $TempExportFolder -ErrorAction Stop
    Log "Export completed successfully." "SUCCESS"

    # Step 3: Restart VM (Minimize downtime)
    if ($WasRunning) {
        Log "Starting VM '$VMName'..."
        Start-VM -Name $VMName
        # Increased timeout to 5 minutes with 15-second check interval
        if (-not (Wait-VMState -VMName $VMName -TargetState 'Running' -TimeoutSeconds 300 -CheckIntervalSeconds 15)) {
            Log "VM failed to reach 'Running' state within 300s. Manual check required." "ERROR"
            $script:Errors += "VM '$VMName' failed to restart after export."
        } else {
            Log "VM is running." "SUCCESS"
        }
    }

    # Step 4: Compress Exported Folder
    Log "Starting compression (mx5)..."
    & $SevenZipPath a -t7z -mx5 "$ArchiveFileName" "$TempExportFolder\*" -bsp1
    if ($LASTEXITCODE -ne 0) {
        throw "7z compression failed with exit code $LASTEXITCODE"
    }
    Log "Archive created: $ArchiveFileName" "SUCCESS"

    # Step 4b: Verify archive integrity
    Log "Verifying archive integrity..."
    & $SevenZipPath t "$ArchiveFileName" -bsp1
    if ($LASTEXITCODE -ne 0) {
        throw "Archive integrity check failed with exit code $LASTEXITCODE. Archive may be corrupted."
    }
    Log "Archive integrity verified successfully." "SUCCESS"

    # Step 5: Remove temporary export folder (only after successful verification)
    if (Test-Path $TempExportFolder) {
        Log "Removing temporary export folder..."
        Remove-Item -Path $TempExportFolder -Recurse -Force
        Log "Temporary folder removed." "SUCCESS"
    }

    # Step 6: Daily Rotation (Keep archives from the last $DaysToKeep calendar days based on filename date)
    Log "Performing rotation (Keep archives from the last $DaysToKeep days)..."
    $cutoffDate = (Get-Date).Date.AddDays(-$DaysToKeep)
    $allFiles = Get-ChildItem -Path $DailyPath -Filter "*.7z"
    $oldFiles = @()

    foreach ($file in $allFiles) {
        # Extract date from filename pattern: VMName-YYYY-MM-DD_HH-MM.7z
        if ($file.Name -match '(\d{4}-\d{2}-\d{2})_\d{2}-\d{2}\.7z$') {
            try {
                $fileDate = [datetime]::ParseExact($matches[1], 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                if ($fileDate -lt $cutoffDate) {
                    $oldFiles += $file
                }
            }
            catch {
                Log "  WARNING: Could not parse date from filename '$($file.Name)'. Using LastWriteTime as fallback." "WARNING"
                if ($file.LastWriteTime -lt $cutoffDate) {
                    $oldFiles += $file
                }
            }
        }
        else {
            Log "  WARNING: Filename '$($file.Name)' does not match expected pattern. Using LastWriteTime as fallback." "WARNING"
            if ($file.LastWriteTime -lt $cutoffDate) {
                $oldFiles += $file
            }
        }
    }

    if ($oldFiles.Count -gt 0) {
        foreach ($file in $oldFiles) {
            try {
                Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                Log "  Deleted old archive (older than $DaysToKeep days): $($file.Name)"
            } catch {
                Add-Error "Failed to delete old archive $($file.Name): $_"
            }
        }
    } else {
        Log "No archives older than $DaysToKeep days found."
    }
}
catch {
    Add-Error "Process failed: $_"
    # Temporary folder is intentionally left for troubleshooting
    Log "Temporary export folder preserved at $TempExportFolder for manual inspection." "WARNING"
}
finally {
    # Final safety: Ensure VM is running if it was running initially
    if ($WasRunning) {
        $currentState = (Get-VM -Name $VMName).State
        if ($currentState -ne 'Running') {
            Log "Final safety check: VM is not running. Attempting to start..." "WARNING"
            Start-VM -Name $VMName -ErrorAction SilentlyContinue
        }
    }

    # Final Notification if errors occurred
    if ($script:Errors.Count -gt 0) {
        Send-ErrorNotification -Subject "Hyper-V Backup/Rotate Issues: $VMName" -Body ($script:Errors -join "`r`n")
    }

    Log "=== TASK FINISHED ===" "INFO"
}