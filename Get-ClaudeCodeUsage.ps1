<#
.SYNOPSIS
  Summarises Claude Code token usage per project per week from local session logs.

.DESCRIPTION
  Reads the JSONL session logs Claude Code writes under ~/.claude/projects,
  totals token usage by project (working folder) and ISO week, and writes a CSV.
  Run it on each computer; point -OutputFolder at a synced folder (OneDrive etc.)
  so every machine's CSV lands in one place for the weekly report.

.EXAMPLE
  ./Get-ClaudeCodeUsage.ps1 -Days 7
  ./Get-ClaudeCodeUsage.ps1 -Days 90 -OutputFolder "$HOME/OneDrive/ClaudeUsage"
#>
[CmdletBinding()]
param(
    [int]$Days = 7,
    [string]$LogRoot = (Join-Path $HOME '.claude/projects'),
    [string]$OutputFolder = (Get-Location).Path
)

if (-not (Test-Path $LogRoot)) { throw "No Claude Code logs found at $LogRoot" }

$cutoff   = (Get-Date).ToUniversalTime().AddDays(-$Days)
$computer = [System.Net.Dns]::GetHostName()
$seen     = [System.Collections.Generic.HashSet[string]]::new()
$rows     = [System.Collections.Generic.List[object]]::new()

Get-ChildItem -Path $LogRoot -Recurse -Filter *.jsonl -File |
    Where-Object { $_.LastWriteTimeUtc -ge $cutoff } |
    ForEach-Object {
        foreach ($line in [System.IO.File]::ReadLines($_.FullName)) {
            if ($line -notmatch '"usage"') { continue }
            try { $e = $line | ConvertFrom-Json -ErrorAction Stop } catch { continue }
            if ($e.type -ne 'assistant' -or -not $e.message.usage) { continue }

            $ts = [datetime]::Parse($e.timestamp).ToUniversalTime()
            if ($ts -lt $cutoff) { continue }

            # Streaming writes the same response more than once; count it once.
            $key = "$($e.message.id)|$($e.requestId)"
            if (-not $seen.Add($key)) { continue }

            $u = $e.message.usage
            $project = if ($e.cwd) { Split-Path $e.cwd -Leaf } else { $_.Directory.Name }
            # ISO week via the Thursday rule (works in Windows PowerShell 5.1 and 7)
            $thu  = $ts.Date.AddDays(3 - (([int]$ts.DayOfWeek + 6) % 7))
            $wk   = [System.Globalization.CultureInfo]::InvariantCulture.Calendar.GetWeekOfYear(
                        $thu, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
            $week = '{0}-W{1:00}' -f $thu.Year, $wk

            $rows.Add([pscustomobject]@{
                Project    = $project
                Week       = $week
                Session    = $e.sessionId
                Model      = $e.message.model
                Input      = [long]$u.input_tokens
                Output     = [long]$u.output_tokens
                CacheWrite = [long]$u.cache_creation_input_tokens
                CacheRead  = [long]$u.cache_read_input_tokens
            })
        }
    }

$summary = $rows | Group-Object Project, Week | ForEach-Object {
    $g = $_.Group
    [pscustomobject]@{
        Computer    = $computer
        Project     = $g[0].Project
        Week        = $g[0].Week
        Sessions    = ($g.Session | Sort-Object -Unique).Count
        Requests    = $g.Count
        InputTokens = ($g | Measure-Object Input      -Sum).Sum
        OutputTokens= ($g | Measure-Object Output     -Sum).Sum
        CacheWrite  = ($g | Measure-Object CacheWrite -Sum).Sum
        CacheRead   = ($g | Measure-Object CacheRead  -Sum).Sum
        Models      = ($g.Model | Where-Object { $_ } | Sort-Object -Unique) -join '; '
    }
} | Sort-Object Week, @{ Expression = 'OutputTokens'; Descending = $true }

New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$out = Join-Path $OutputFolder ("claude-usage-{0}-{1:yyyy-MM-dd}.csv" -f $computer, (Get-Date))
$summary | Export-Csv -Path $out -NoTypeInformation -Encoding utf8
$summary | Format-Table Project, Week, Sessions, InputTokens, OutputTokens, CacheRead -AutoSize
Write-Host "Saved $out"
