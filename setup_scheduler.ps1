# setup_scheduler.ps1
# Creates a Windows Task Scheduler entry that runs the pipeline every hour
# from 8:00 AM to 11:00 PM daily - catches whichever hour a source report
# actually lands by email, instead of only picking it up once overnight.
# Run once as Administrator: powershell -ExecutionPolicy Bypass -File .\setup_scheduler.ps1
# (from inside the GLIM folder - works on any drive/path, self-locating)

$TaskName        = "GL-Dashboard-Pipeline"
$WorkDir         = $PSScriptRoot
$ScriptPath      = Join-Path $WorkDir "run_pipeline.py"
$PythonPath      = (Get-Command python).Source
$WindowStart     = "08:00"
$RepeatEvery     = New-TimeSpan -Hours 1
$WindowDuration  = New-TimeSpan -Hours 15   # 08:00 -> 23:00

# Remove existing task if it exists
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed existing task: $TaskName"
}

$Action = New-ScheduledTaskAction -Execute $PythonPath -Argument $ScriptPath -WorkingDirectory $WorkDir

$Trigger = New-ScheduledTaskTrigger -Daily -At $WindowStart
$Trigger.Repetition = (New-ScheduledTaskTrigger -Once -At $WindowStart -RepetitionInterval $RepeatEvery -RepetitionDuration $WindowDuration).Repetition

$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Description "GL Dashboard - auto-fetch email, parse Excel, update dashboard (hourly, 8 AM-11 PM)" -RunLevel Highest

Write-Host ""
Write-Host "Task '$TaskName' scheduled: every hour from $WindowStart, for $($WindowDuration.TotalHours) hours (through approx 23:00)"
Write-Host "Working directory: $WorkDir"
Write-Host "To run it manually: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "To view logs:       Get-Content '$WorkDir\pipeline.log' -Tail 50"
