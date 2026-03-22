# register_task.ps1
# Called automatically by install.bat (runs as Administrator via UAC).
# Can also be run manually as Administrator to re-register the task.

param(
    [string]$InstallDir = "C:\ProgramData\ETIMSBackup"
)

$ExePath    = "$InstallDir\main.exe"
$ConfigPath = "$InstallDir\config.json"
$taskName   = "ETIMSBackupUploader"

if (-not (Test-Path $ConfigPath)) {
    Write-Host "[ERROR] config.json not found in $InstallDir"
    exit 1
}

# Read schedule time from config.json
$config       = Get-Content $ConfigPath | ConvertFrom-Json
$scheduleTime = $config.schedule_time

# Remove old task if exists
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Removed existing task."
}

$action   = New-ScheduledTaskAction -Execute $ExePath -WorkingDirectory $InstallDir
$trigger  = New-ScheduledTaskTrigger -Daily -At $scheduleTime
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable

$principal = New-ScheduledTaskPrincipal `
    -UserId "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

Register-ScheduledTask `
    -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Uploads latest ETIMS backup to Google Drive. Installed in $InstallDir"

Write-Host "Task '$taskName' registered - runs daily at $scheduleTime."
Write-Host "StartWhenAvailable is ON: missed runs fire as soon as the PC comes back online."
