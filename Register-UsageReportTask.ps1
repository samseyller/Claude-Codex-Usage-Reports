#Requires -Version 7.0
<#
.SYNOPSIS
  Registers (or updates) a Windows scheduled task that runs Invoke-UsageReport.ps1 weekly.

.DESCRIPTION
  Runs as the current user, only while logged on (no stored password). If the computer is off or
  asleep at the scheduled time, the task runs as soon as possible afterwards. Also removes the
  older Claude-only task if it is still registered.

.EXAMPLE
  ./Register-UsageReportTask.ps1
  ./Register-UsageReportTask.ps1 -DayOfWeek Friday -At 17:00
  ./Register-UsageReportTask.ps1 -DataFolder "$HOME/OneDrive/AIUsage" -CollectOnly   # secondary machine
  ./Register-UsageReportTask.ps1 -Unregister
#>
[CmdletBinding()]
param(
    [string]$TaskName = 'AI Coding Usage Weekly Report',
    [DayOfWeek]$DayOfWeek = 'Monday',
    [datetime]$At = '09:00',
    [string]$DataFolder,
    [switch]$CollectOnly,
    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'

foreach ($old in 'Claude Code Weekly Usage Report') {
    if ($old -ne $TaskName -and (Get-ScheduledTask -TaskName $old -ErrorAction SilentlyContinue)) {
        Unregister-ScheduledTask -TaskName $old -Confirm:$false
        Write-Host "Removed old scheduled task '$old'."
    }
}

if ($Unregister) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed scheduled task '$TaskName'."
    return
}

$script    = Join-Path $PSScriptRoot 'Invoke-UsageReport.ps1'
$pwsh      = (Get-Command pwsh -ErrorAction Stop).Source
$arguments = "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$script`""
if ($DataFolder)  { $arguments += " -DataFolder `"$DataFolder`"" }
if ($CollectOnly) { $arguments += ' -CollectOnly' }

$action    = New-ScheduledTaskAction -Execute $pwsh -Argument $arguments -WorkingDirectory $PSScriptRoot
$trigger   = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek -At $At
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                 -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId ([Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal `
    -Description 'Collects Claude Code and Codex token usage from local session logs and writes weekly and lifetime reports.' -Force | Out-Null

$task = Get-ScheduledTask -TaskName $TaskName
Write-Host "Registered '$TaskName': every $DayOfWeek at $($At.ToString('HH:mm')), next run $((($task | Get-ScheduledTaskInfo).NextRunTime))."
Write-Host "Run it now with: Start-ScheduledTask -TaskName '$TaskName'"
