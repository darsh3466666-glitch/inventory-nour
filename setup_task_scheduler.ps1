$TaskName = "InventoryWatcher_Nour"
$BatPath = "C:\Users\GoldenTech\.openclaw-autoclaw\workspace\.cluster\DELIVERY\start_watcher.bat"

Write-Host "=== Setup Task Scheduler ===" -ForegroundColor Cyan

if (-not (Test-Path $BatPath)) {
    Write-Host "ERROR: Bat file not found" -ForegroundColor Red
    exit 1
}

$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Removing old task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

$Action = New-ScheduledTaskAction -Execute $BatPath
$Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 1)
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -DontStopOnIdleEnd -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest

Write-Host "Creating task..." -ForegroundColor Yellow
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description "Inventory Watcher Nour - updates every 1 min" -Force | Out-Null

Write-Host "Task created successfully!" -ForegroundColor Green

$task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($task) {
    $info = $task | Get-ScheduledTaskInfo
    Write-Host "State: $($task.State)" -ForegroundColor White
    Write-Host "LastRun: $($info.LastRunTime)" -ForegroundColor White
    Write-Host "NextRun: $($info.NextRunTime)" -ForegroundColor White
}
