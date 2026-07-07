# setup_scheduler.ps1
# Creates a Windows Task Scheduler entry that runs the pipeline daily at 8:00 AM.
# Run once as Administrator: powershell -ExecutionPolicy Bypass -File E:\GLIM\setup_scheduler.ps1

$TaskName    = "GL-Dashboard-Pipeline"
$ScriptPath  = "E:\GLIM\run_pipeline.py"
$PythonPath  = (Get-Command python).Source
$WorkDir     = "E:\GLIM"
$TriggerTime = "08:00"

# Remove existing task if it exists
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed existing task: $TaskName"
}

$Action  = New-ScheduledTaskAction `
    -Execute $PythonPath `
    -Argument $ScriptPath `
    -WorkingDirectory $WorkDir

$Trigger = New-ScheduledTaskTrigger -Daily -At $TriggerTime

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action   $Action `
    -Trigger  $Trigger `
    -Settings $Settings `
    -Description "GL Dashboard — auto-fetch email, parse Excel, update dashboard" `
    -RunLevel Highest

Write-Host ""
Write-Host "Task '$TaskName' scheduled to run daily at $TriggerTime"
Write-Host "To run it manually: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "To view logs:       Get-Content E:\GLIM\pipeline.log -Tail 50"
