#Requires -Version 7.0
<#
.SYNOPSIS
  Weekly Claude Code + Codex usage reports with lifetime statistics.

.DESCRIPTION
  1. Collect - reads local session logs and upserts per day / session / project / model token totals
     into <DataFolder>/<tool>/history/usage-<computer>.csv:
       Claude Code  ~/.claude/projects/**/*.jsonl   (Claude Code deletes these after ~30 days)
       Codex        ~/.codex/sessions/**/*.jsonl and ~/.codex/archived_sessions
     The history files are the lifetime record. A row is only replaced by one with at least as many
     requests, so logs expiring never shrink it.

  2. Report - writes three report folders, each with the same layout:
       <DataFolder>/claude/    Claude Code only
       <DataFolder>/codex/     Codex only
       <DataFolder>/combined/  both tools together
     Each holds weekly/<week>/<folder>-usage-<week>.csv|html and lifetime.json|html. Every history
     file in a tool's history folder is included, so pointing DataFolder at a synced folder merges
     several computers.

  By default the report covers the last completed ISO week (Monday-Sunday, local time). Missing
  weeks back to each tool's first recorded week are backfilled.

.EXAMPLE
  ./Invoke-UsageReport.ps1
  ./Invoke-UsageReport.ps1 -Week 2026-W42          # a specific (or in-progress) week
  ./Invoke-UsageReport.ps1 -DataFolder "$HOME/OneDrive/AIUsage" -CollectOnly
#>
[CmdletBinding()]
param(
    [string]$DataFolder = (Join-Path $PSScriptRoot 'data'),
    [string]$ClaudeLogRoot = (Join-Path $HOME '.claude/projects'),
    [string]$CodexHome = (Join-Path $HOME '.codex'),
    [ValidatePattern('^\d{4}-W\d{2}$')]
    [string]$Week,
    [int]$LookbackDays = 35,
    [int]$TopProjects = 25,
    [switch]$FullScan,
    [switch]$CollectOnly,
    [switch]$ReportOnly
)

$ErrorActionPreference = 'Stop'
$Inv      = [cultureinfo]::InvariantCulture
$Computer = [System.Net.Dns]::GetHostName()
$Metrics  = 'Requests', 'InputTokens', 'OutputTokens', 'CacheWriteTokens', 'CacheReadTokens', 'ReasoningTokens'
$HistCols = @('Computer', 'Project', 'Date', 'Week', 'Session', 'Model') + $Metrics

# Tool order fixes the chart colour of each tool (series-1 blue, series-2 orange) on every page.
$Tools = [ordered]@{
    claude = @{ Name = 'Claude Code'; Series = 's1' }
    codex  = @{ Name = 'Codex';       Series = 's2' }
}
$Views = [ordered]@{
    claude   = @{ Title = 'Claude Code';         Tools = @('claude') }
    codex    = @{ Title = 'Codex';               Tools = @('codex') }
    combined = @{ Title = 'Claude Code + Codex'; Tools = @('claude', 'codex') }
}

#region Helpers ------------------------------------------------------------------------------

function Get-IsoWeekLabel([datetime]$Date) {
    '{0}-W{1:00}' -f [Globalization.ISOWeek]::GetYear($Date), [Globalization.ISOWeek]::GetWeekOfYear($Date)
}

function Get-IsoWeekStart([string]$Label) {
    $null = $Label -match '^(\d{4})-W(\d{2})$'
    [Globalization.ISOWeek]::ToDateTime([int]$Matches[1], [int]$Matches[2], [DayOfWeek]::Monday)
}

function Get-JProp([System.Text.Json.JsonElement]$Element, [string]$Name) {
    $out = [System.Text.Json.JsonElement]::new()
    if ($Element.ValueKind -eq 'Object' -and $Element.TryGetProperty($Name, [ref]$out)) { return $out }
    $null
}

function Get-JString([System.Text.Json.JsonElement]$Element, [string]$Name) {
    $p = Get-JProp $Element $Name
    if ($null -ne $p -and $p.ValueKind -eq 'String') { $p.GetString() } else { '' }
}

function Get-JLong([System.Text.Json.JsonElement]$Element, [string]$Name) {
    $p = Get-JProp $Element $Name
    if ($null -ne $p -and $p.ValueKind -eq 'Number') { $p.GetInt64() } else { 0L }
}

function Get-Leaf([string]$Path) {
    $parts = $Path -split '[\\/]' | Where-Object { $_ }
    if ($parts) { $parts[-1] } else { '' }
}

# Adds one API request to the collection aggregate (keyed by project, local date, session, model).
function Add-Usage($Agg, [string]$Project, [datetime]$Local, [string]$Session, [string]$Model,
                   [long]$In, [long]$Out, [long]$CacheWrite, [long]$CacheRead, [long]$Reasoning) {
    $date = $Local.ToString('yyyy-MM-dd', $Inv)
    $k = "$Project|$date|$Session|$Model"
    $row = $Agg[$k]
    if (-not $row) {
        $row = [pscustomobject]@{
            Computer = $Computer; Project = $Project; Date = $date; Week = Get-IsoWeekLabel $Local.Date
            Session = $Session; Model = $Model; Requests = 0L; InputTokens = 0L; OutputTokens = 0L
            CacheWriteTokens = 0L; CacheReadTokens = 0L; ReasoningTokens = 0L
        }
        $Agg[$k] = $row
    }
    $row.Requests++
    $row.InputTokens += $In; $row.OutputTokens += $Out; $row.CacheWriteTokens += $CacheWrite
    $row.CacheReadTokens += $CacheRead; $row.ReasoningTokens += $Reasoning
}

# Sums a set of history rows into one totals object.
function Get-Totals([object[]]$Set) {
    $t = [ordered]@{}; foreach ($m in $Metrics) { $t[$m] = 0L }
    $sessions  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $days      = [System.Collections.Generic.SortedSet[string]]::new()
    $projects  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $models    = [System.Collections.Generic.SortedSet[string]]::new()
    $computers = [System.Collections.Generic.SortedSet[string]]::new()
    $byTool    = @{}; foreach ($k in $Tools.Keys) { $byTool[$k] = 0L }
    foreach ($r in $Set) {
        if ($null -eq $r) { continue }
        foreach ($m in $Metrics) { $t[$m] += $r.$m }
        $null = $sessions.Add("$($r.Tool)|$($r.Computer)|$($r.Session)")
        $null = $days.Add($r.Date)
        $null = $projects.Add($r.Project)
        if ($r.Model) { $null = $models.Add($r.Model) }
        $null = $computers.Add($r.Computer)
        $byTool[$r.Tool] += $r.OutputTokens
    }
    $t.TotalTokens = $t.InputTokens + $t.OutputTokens + $t.CacheWriteTokens + $t.CacheReadTokens
    $t.Sessions    = $sessions.Count
    $t.ActiveDays  = $days.Count
    $t.Projects    = $projects.Count
    $t.FirstDate   = if ($days.Count) { $days.Min } else { $null }
    $t.LastDate    = if ($days.Count) { $days.Max } else { $null }
    $t.Models      = @($models)
    $t.Computers   = @($computers)
    $t.ByTool      = $byTool
    [pscustomobject]$t
}

# Totals per value of one property, biggest output first, each with a Name.
function Group-Totals([object[]]$Set, [string]$Property) {
    @(foreach ($g in ($Set | Group-Object $Property)) {
        (Get-Totals $g.Group) | Add-Member Name $g.Group[0].$Property -PassThru
    }) | Sort-Object OutputTokens -Descending
}

function Format-Compact([double]$n) {
    $a = [math]::Abs($n)
    if     ($a -ge 1e9) { ($n / 1e9).ToString('0.##', $Inv) + 'B' }
    elseif ($a -ge 1e6) { ($n / 1e6).ToString('0.##', $Inv) + 'M' }
    elseif ($a -ge 1e3) { ($n / 1e3).ToString('0.#',  $Inv) + 'K' }
    else                { $n.ToString('0', $Inv) }
}

function Format-N([double]$n) { $n.ToString('N0', $Inv) }
function Format-Svg([double]$n) { $n.ToString('0.##', $Inv) }
function Encode([object]$s) { [System.Net.WebUtility]::HtmlEncode([string]$s) }

function Format-Delta([double]$Now, [double]$Before) {
    if ($Before -eq 0) { return $(if ($Now -eq 0) { '<span class="delta">no change</span>' } else { '<span class="delta">none the week before</span>' }) }
    $pct = ($Now - $Before) / $Before * 100
    $arrow = if ($pct -gt 0.5) { '&#9650;' } elseif ($pct -lt -0.5) { '&#9660;' } else { '&#9644;' }
    '<span class="delta">{0} {1}% vs last week</span>' -f $arrow, ([math]::Abs($pct)).ToString('0', $Inv)
}

function Get-NiceStep([double]$Raw) {
    if ($Raw -le 0) { return 1 }
    $mag = [math]::Pow(10, [math]::Floor([math]::Log10($Raw)))
    foreach ($f in 1, 2, 2.5, 5, 10) { if ($f * $mag -ge $Raw) { return $f * $mag } }
}

# Column chart as inline SVG, stacked when given several series (@{ Series = 's1'; Values = ... }).
# Columns carry data-tip for the hover tooltip; the tallest column gets a value label at its cap.
function New-ColumnChart {
    param([string[]]$Labels, [object[]]$Series, [string[]]$Tips, [int]$LabelEvery = 1,
          [int]$Width = 760, [int]$Height = 240, [string]$AriaLabel)
    $padL = 52; $padR = 8; $padT = 22; $padB = 28
    $plotW = $Width - $padL - $padR; $plotH = $Height - $padT - $padB
    $n = $Labels.Count
    $totals = [double[]]::new($n)
    foreach ($s in $Series) { for ($i = 0; $i -lt $n; $i++) { $totals[$i] += $s.Values[$i] } }
    $max  = [math]::Max(1, ($totals | Measure-Object -Maximum).Maximum)
    $step = Get-NiceStep ($max / 4)
    $top  = [math]::Ceiling($max / $step) * $step
    $base = $padT + $plotH
    $sb = [System.Text.StringBuilder]::new()
    $null = $sb.Append(('<svg class="chart" viewBox="0 0 {0} {1}" role="img" aria-label="{2}">' -f $Width, $Height, (Encode $AriaLabel)))
    for ($v = 0; $v -le $top + 1e-9; $v += $step) {
        $y = Format-Svg ($base - $v / $top * $plotH)
        $null = $sb.Append(('<line class="grid" x1="{0}" x2="{1}" y1="{2}" y2="{2}"/>' -f $padL, ($Width - $padR), $y))
        $null = $sb.Append(('<text class="axis" x="{0}" y="{1}" text-anchor="end" dy="0.32em">{2}</text>' -f ($padL - 8), $y, (Format-Compact $v)))
    }
    $band = $plotW / [math]::Max(1, $n)
    $bw = [math]::Min(24, $band * 0.62)
    $maxIdx = [array]::IndexOf($totals, [double]($totals | Measure-Object -Maximum).Maximum)
    for ($i = 0; $i -lt $n; $i++) {
        $x = $padL + $i * $band + ($band - $bw) / 2
        $null = $sb.Append(('<g class="bar" data-tip="{0}"><rect class="hit" x="{1}" y="{2}" width="{3}" height="{4}"/>' -f (Encode $Tips[$i]), (Format-Svg ($padL + $i * $band)), $padT, (Format-Svg $band), $plotH))
        $segs = @($Series | Where-Object { $_.Values[$i] -gt 0 })
        $cum = 0.0
        for ($j = 0; $j -lt $segs.Count; $j++) {
            $y0 = $base - $cum / $top * $plotH
            $cum += $segs[$j].Values[$i]
            $y1 = [math]::Min($y0 - 1.5, $base - $cum / $top * $plotH)
            if ($j -gt 0 -and $y0 - $y1 -gt 3) { $y0 -= 2 }   # 2px surface gap between stacked segments
            if ($j -eq $segs.Count - 1) {
                $r = [math]::Min(4, [math]::Min($y0 - $y1, $bw / 2))
                $d = 'M{0},{1}L{0},{2}Q{0},{3} {4},{3}L{5},{3}Q{6},{3} {6},{2}L{6},{1}Z' -f
                    (Format-Svg $x), (Format-Svg $y0), (Format-Svg ($y1 + $r)), (Format-Svg $y1),
                    (Format-Svg ($x + $r)), (Format-Svg ($x + $bw - $r)), (Format-Svg ($x + $bw))
            } else {
                $d = 'M{0},{1}L{0},{2}L{3},{2}L{3},{1}Z' -f (Format-Svg $x), (Format-Svg $y0), (Format-Svg $y1), (Format-Svg ($x + $bw))
            }
            $null = $sb.Append("<path class=`"mark $($segs[$j].Series)`" d=`"$d`"/>")
            if ($i -eq $maxIdx -and $j -eq $segs.Count - 1) {
                # Keep the label clear of the y-axis labels and the right edge.
                $text = Format-Compact $totals[$i]; $half = $text.Length * 3.5
                $cx = $x + $bw / 2; $anchor = 'middle'
                if ($cx - $half -lt $padL) { $cx = $x; $anchor = 'start' }
                elseif ($cx + $half -gt $Width - $padR) { $cx = $x + $bw; $anchor = 'end' }
                $null = $sb.Append(('<text class="cap" x="{0}" y="{1}" text-anchor="{2}">{3}</text>' -f (Format-Svg $cx), (Format-Svg ($y1 - 6)), $anchor, $text))
            }
        }
        $null = $sb.Append('</g>')
        if ($i % $LabelEvery -eq 0) {
            $null = $sb.Append(('<text class="axis" x="{0}" y="{1}" text-anchor="middle">{2}</text>' -f (Format-Svg ($padL + $i * $band + $band / 2)), ($base + 18), (Encode $Labels[$i])))
        }
    }
    $null = $sb.Append(('<line class="baseline" x1="{0}" x2="{1}" y1="{2}" y2="{2}"/></svg>' -f $padL, ($Width - $padR), $base))
    $sb.ToString()
}

function New-Legend([string[]]$ToolKeys) {
    if ($ToolKeys.Count -lt 2) { return '' }
    '<div class="legend">' + (($ToolKeys | ForEach-Object { '<span><i class="sw {0}"></i>{1}</span>' -f $Tools[$_].Series, $Tools[$_].Name }) -join '') + '</div>'
}

function New-ShareCell([double]$Value, [double]$Max) {
    $pct = if ($Max -gt 0) { [math]::Round($Value / $Max * 100, 1) } else { 0 }
    '<td class="num" data-sort="{0}"><div class="share"><span>{1}</span><i style="--w:{2}%"></i></div></td>' -f
        $Value.ToString('0', $Inv), (Format-Compact $Value), $pct.ToString('0.#', $Inv)
}

# Sortable table. Each column: @{ H = heading; T = share|compact|n|text|wrap|html|rank; V = { param($r) ... } }.
# The first 'share' column is the default sort (descending).
function New-Table([object[]]$Items, [object[]]$Columns) {
    $max = @{}
    $head = foreach ($c in $Columns) {
        if ($c.T -eq 'share') { $max[$c.H] = [double](@($Items | ForEach-Object { & $c.V $_ }) | Measure-Object -Maximum).Maximum }
        $cls  = if ($c.T -in 'share', 'compact', 'n', 'rank') { ' class="num"' } else { '' }
        $sort = if ($c.T -eq 'share' -and -not $sorted) { $sorted = $true; ' aria-sort="descending"' } else { '' }
        "<th$cls$sort>$(Encode $c.H)</th>"
    }
    $rows = foreach ($it in $Items) {
        $cells = foreach ($c in $Columns) {
            $v = & $c.V $it
            switch ($c.T) {
                'share'   { New-ShareCell $v $max[$c.H] }
                'compact' { '<td class="num" data-sort="{0}">{1}</td>' -f ([double]$v).ToString('0', $Inv), (Format-Compact $v) }
                'n'       { '<td class="num" data-sort="{0}">{1}</td>' -f ([double]$v).ToString('0', $Inv), (Format-N $v) }
                'rank'    { '<td class="num rank" data-sort="{0}">{0}</td>' -f $v }
                'wrap'    { '<td class="wrap">{0}</td>' -f (Encode $v) }
                'html'    { "<td>$v</td>" }
                default   { '<td>{0}</td>' -f (Encode $v) }
            }
        }
        '<tr>' + ($cells -join '') + '</tr>'
    }
    '<div class="scroll"><table class="sortable"><thead><tr>' + ($head -join '') + '</tr></thead><tbody>' + ($rows -join "`n") + '</tbody></table></div>'
}

function New-Tiles([object[]]$Tiles) {
    $html = foreach ($t in $Tiles) {
        '<div class="tile"><div class="label">{0}</div><div class="value">{1}</div>{2}</div>' -f $t[0], $t[1], $t[2]
    }
    '<section class="tiles">' + ($html -join '') + '</section>'
}

# Header switcher between the three report folders. $Rel maps a view key to the href for that view.
function New-ViewNav([string]$Current, [scriptblock]$Rel) {
    $links = foreach ($k in $Views.Keys) {
        if ($k -eq $Current) { "<b>$(Encode $Views[$k].Title)</b>" } else { "<a href=`"$(& $Rel $k)`">$(Encode $Views[$k].Title)</a>" }
    }
    '<div class="views">' + ($links -join '') + '</div>'
}

$PageCss = @'
:root {
  color-scheme: light;
  --bg: #f4f4f2; --surface: #fcfcfb; --border: #e4e3df; --grid: #ebeae6;
  --text: #0b0b0b; --text-2: #52514e; --muted: #73726d; --link: #1c5cab;
  --s1: #2a78d6; --s2: #eb6834; --neutral: #8a8984; --track: rgba(115,114,109,.16);
}
@media (prefers-color-scheme: dark) {
  :root {
    color-scheme: dark;
    --bg: #111110; --surface: #1a1a19; --border: #2e2e2b; --grid: #2a2a27;
    --text: #f0efec; --text-2: #c3c2b7; --muted: #94938b; --link: #86b6ef;
    --s1: #3987e5; --s2: #d95926; --neutral: #8f8e86; --track: rgba(195,194,183,.14);
  }
}
body.v-claude { --series: var(--s1); } body.v-codex { --series: var(--s2); } body.v-combined { --series: var(--neutral); }
* { box-sizing: border-box; }
body { margin: 0; background: var(--bg); color: var(--text);
  font: 14px/1.5 "Segoe UI Variable Text", "Segoe UI", system-ui, -apple-system, sans-serif; }
main { max-width: 1040px; margin: 0 auto; padding: 24px 16px 64px; }
a { color: var(--link); }
.views { display: inline-flex; gap: 4px; padding: 3px; border: 1px solid var(--border); border-radius: 999px; background: var(--surface); margin-bottom: 20px; font-size: 13px; }
.views a, .views b { padding: 4px 12px; border-radius: 999px; text-decoration: none; }
.views a { color: var(--text-2); } .views a:hover { color: var(--text); }
.views b { background: var(--text); color: var(--bg); font-weight: 600; }
header { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: end; gap: 8px 24px; margin-bottom: 24px; }
header h1 { margin: 0; font-size: 26px; font-weight: 600; letter-spacing: -.01em; }
header p { margin: 4px 0 0; color: var(--text-2); }
nav { display: flex; gap: 16px; flex-wrap: wrap; font-size: 13px; }
h2 { font-size: 16px; font-weight: 600; margin: 0 0 4px; }
.sub { color: var(--muted); font-size: 13px; margin: 0 0 16px; }
.card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px; margin-bottom: 16px; }
.tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin-bottom: 16px; }
.tile { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px; }
.tile .label { color: var(--text-2); font-size: 13px; }
.tile .value { font-size: 26px; font-weight: 600; line-height: 1.25; margin-top: 2px; }
.tile .hint, .delta { color: var(--muted); font-size: 12px; }
.records { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px 24px; }
.records div span { display: block; color: var(--text-2); font-size: 13px; }
.records div b { font-size: 16px; font-weight: 600; }
.records div em { font-style: normal; color: var(--muted); font-size: 12px; display: block; }
.legend { display: flex; gap: 16px; font-size: 13px; color: var(--text-2); margin: -8px 0 8px; }
.legend span { display: inline-flex; align-items: center; gap: 6px; }
.sw { display: inline-block; width: 10px; height: 10px; border-radius: 3px; vertical-align: -1px; margin-right: 6px; }
.legend .sw { margin-right: 0; }
.sw.s1 { background: var(--s1); } .sw.s2 { background: var(--s2); }
.chart { width: 100%; height: auto; display: block; }
.chart .grid { stroke: var(--grid); stroke-width: 1; }
.chart .baseline { stroke: var(--border); stroke-width: 1; }
.chart .axis { fill: var(--muted); font-size: 11px; }
.chart .cap { fill: var(--text-2); font-size: 11px; font-weight: 600; }
.chart .mark.s1 { fill: var(--s1); } .chart .mark.s2 { fill: var(--s2); }
.chart .hit { fill: transparent; }
.chart .bar:hover .hit { fill: var(--track); }
.scroll { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { padding: 8px 10px; text-align: left; border-bottom: 1px solid var(--border); white-space: nowrap; }
th { color: var(--text-2); font-weight: 600; cursor: pointer; user-select: none; }
th[aria-sort="descending"]::after { content: " \25BE"; } th[aria-sort="ascending"]::after { content: " \25B4"; }
td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
td.wrap { white-space: normal; color: var(--text-2); min-width: 140px; }
td.rank { color: var(--muted); width: 1%; }
tbody tr:hover { background: var(--track); }
.share { display: flex; align-items: center; justify-content: flex-end; gap: 8px; }
.share i { display: block; width: 90px; height: 6px; border-radius: 3px;
  background: linear-gradient(90deg, var(--series) var(--w), var(--track) var(--w)); }
.empty { color: var(--text-2); padding: 24px 0; text-align: center; }
.note { color: var(--muted); font-size: 12px; margin-top: 24px; }
#tip { position: fixed; pointer-events: none; background: var(--text); color: var(--bg); font-size: 12px;
  padding: 6px 9px; border-radius: 6px; white-space: pre; opacity: 0; transition: opacity .1s; z-index: 10; }
@media (max-width: 600px) { .share i { width: 48px; } header h1 { font-size: 22px; } }
'@

$PageJs = @'
const tip = document.getElementById('tip');
document.querySelectorAll('[data-tip]').forEach(el => {
  el.addEventListener('mousemove', e => {
    tip.textContent = el.dataset.tip; tip.style.opacity = 1;
    const x = Math.min(e.clientX + 14, innerWidth - tip.offsetWidth - 8);
    tip.style.left = x + 'px'; tip.style.top = (e.clientY - tip.offsetHeight - 10) + 'px';
  });
  el.addEventListener('mouseleave', () => tip.style.opacity = 0);
});
document.querySelectorAll('table.sortable th').forEach((th, col) => th.addEventListener('click', () => {
  const tbody = th.closest('table').tBodies[0];
  const desc = th.getAttribute('aria-sort') !== 'descending';
  th.closest('tr').querySelectorAll('th').forEach(h => h.removeAttribute('aria-sort'));
  th.setAttribute('aria-sort', desc ? 'descending' : 'ascending');
  const key = td => td.dataset.sort ?? td.textContent.trim().toLowerCase();
  [...tbody.rows].sort((a, b) => {
    let x = key(a.cells[col]), y = key(b.cells[col]);
    const nx = parseFloat(x), ny = parseFloat(y);
    const c = (!isNaN(nx) && !isNaN(ny)) ? nx - ny : x.localeCompare(y);
    return desc ? -c : c;
  }).forEach(r => tbody.appendChild(r));
}));
'@

function New-Page([string]$Title, [string]$View, [string]$Body) {
    @"
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(Encode $Title)</title>
<style>$PageCss</style>
</head>
<body class="v-$View">
<main>
$Body
<p class="note">Generated $((Get-Date).ToString('yyyy-MM-dd HH:mm', $Inv)) on $(Encode $Computer) from local session logs. Output tokens include reasoning tokens; Claude Code's synthetic (non-API) messages are excluded. Click a column heading to sort.</p>
</main>
<div id="tip" role="tooltip"></div>
<script>$PageJs</script>
</body>
</html>
"@
}

#endregion

#region Collect ------------------------------------------------------------------------------

function Read-ClaudeLogs([datetime]$Cutoff) {
    $agg = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not (Test-Path $ClaudeLogRoot)) { Write-Warning "No Claude Code logs at $ClaudeLogRoot"; return $agg }
    $root  = (Resolve-Path $ClaudeLogRoot).Path.TrimEnd('\', '/')
    $seen  = [System.Collections.Generic.HashSet[string]]::new()
    $files = @(Get-ChildItem -Path $root -Recurse -Filter *.jsonl -File | Where-Object LastWriteTime -ge $Cutoff)
    Write-Host "Claude Code: scanning $($files.Count) log file(s)..."

    foreach ($file in $files) {
        # Fallback project name: the log folder directly under the root (subagent logs sit deeper).
        $fallback = $file.FullName.Substring($root.Length).TrimStart('\', '/').Split([char[]]'\/')[0]
        foreach ($line in [System.IO.File]::ReadLines($file.FullName)) {
            if (-not $line.Contains('"usage"')) { continue }
            try { $doc = [System.Text.Json.JsonDocument]::Parse($line) } catch { continue }
            try {
                $e = $doc.RootElement
                if ((Get-JString $e 'type') -ne 'assistant') { continue }
                $msg = Get-JProp $e 'message'
                if ($null -eq $msg) { continue }
                $u = Get-JProp $msg 'usage'
                if ($null -eq $u) { continue }
                $model = Get-JString $msg 'model'
                if ($model -eq '<synthetic>') { continue }

                # Streaming writes the same response more than once; count it once.
                $key = '{0}|{1}' -f (Get-JString $msg 'id'), (Get-JString $e 'requestId')
                if ($key -eq '|') { $key = Get-JString $e 'uuid' }
                if (-not $seen.Add($key)) { continue }

                $ts = [datetimeoffset]::Parse((Get-JString $e 'timestamp'), $Inv).LocalDateTime
                if ($ts -lt $Cutoff) { continue }
                $cwd = Get-JString $e 'cwd'
                $project = if ($cwd) { Get-Leaf $cwd } else { $fallback }

                Add-Usage $agg $project $ts (Get-JString $e 'sessionId') $model `
                    (Get-JLong $u 'input_tokens') (Get-JLong $u 'output_tokens') `
                    (Get-JLong $u 'cache_creation_input_tokens') (Get-JLong $u 'cache_read_input_tokens') 0
            }
            finally { $doc.Dispose() }
        }
    }
    $agg
}

# Codex rollout files log one token_usage_record per API response. Older files only have
# token_count events: each carries the last response's usage plus a running total, and is sometimes
# repeated, so a token_count whose running total hasn't changed is a duplicate. (The running total
# itself resets on context compaction, so it can't be used directly.) Codex input_tokens include
# cached tokens; they are split out so the columns mean the same as Claude Code's.
function Read-CodexLogs([datetime]$Cutoff) {
    $agg   = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $roots = @('sessions', 'archived_sessions' | ForEach-Object { Join-Path $CodexHome $_ } | Where-Object { Test-Path $_ })
    if (-not $roots) { Write-Warning "No Codex sessions under $CodexHome"; return $agg }
    $seen  = [System.Collections.Generic.HashSet[string]]::new()
    $files = @(Get-ChildItem -Path $roots -Recurse -Filter *.jsonl -File | Where-Object LastWriteTime -ge $Cutoff)
    Write-Host "Codex: scanning $($files.Count) session file(s)..."

    foreach ($file in $files) {
        $session = ''; $cwd = ''; $model = ''; $prevTotal = $null
        $records = [System.Collections.Generic.List[object]]::new()
        $counts  = [System.Collections.Generic.List[object]]::new()
        foreach ($line in [System.IO.File]::ReadLines($file.FullName)) {
            if (-not ($line.Contains('"token_usage_record"') -or $line.Contains('"token_count"') -or
                      $line.Contains('"turn_context"') -or $line.Contains('"session_meta"'))) { continue }
            try { $doc = [System.Text.Json.JsonDocument]::Parse($line) } catch { continue }
            try {
                $e = $doc.RootElement
                $p = Get-JProp $e 'payload'
                if ($null -eq $p) { continue }
                $ts = Get-JString $e 'timestamp'
                switch (Get-JString $e 'type') {
                    'session_meta' {
                        if (-not $session) { $session = Get-JString $p 'id'; $cwd = Get-JString $p 'cwd' }
                    }
                    'turn_context' {
                        $m = Get-JString $p 'model'; if ($m) { $model = $m }
                        $c = Get-JString $p 'cwd';   if ($c) { $cwd = $c }
                    }
                    'token_usage_record' {
                        $u = Get-JProp $p 'usage'
                        if ($null -ne $u) {
                            $records.Add(@{ Key = 'r|' + (Get-JString $p 'response_id'); Ts = $ts; Model = $model; Cwd = $cwd
                                In = Get-JLong $u 'input_tokens'; Cached = Get-JLong $u 'cached_input_tokens'
                                Write = Get-JLong $u 'cache_write_input_tokens'; Out = Get-JLong $u 'output_tokens'
                                Reason = Get-JLong $u 'reasoning_output_tokens' })
                        }
                    }
                    'event_msg' {
                        if ((Get-JString $p 'type') -ne 'token_count') { break }
                        $info = Get-JProp $p 'info'
                        if ($null -eq $info) { break }
                        $total = Get-JProp $info 'total_token_usage'
                        $u     = Get-JProp $info 'last_token_usage'
                        if ($null -eq $total -or $null -eq $u) { break }
                        $raw = $total.GetRawText()
                        if ($raw -eq $prevTotal) { break }
                        $prevTotal = $raw
                        $counts.Add(@{ Key = "t|$session|$ts|$raw"; Ts = $ts; Model = $model; Cwd = $cwd
                            In = Get-JLong $u 'input_tokens'; Cached = Get-JLong $u 'cached_input_tokens'
                            Write = Get-JLong $u 'cache_write_input_tokens'; Out = Get-JLong $u 'output_tokens'
                            Reason = Get-JLong $u 'reasoning_output_tokens' })
                    }
                }
            }
            finally { $doc.Dispose() }
        }

        if (-not $session) { $session = $file.BaseName }
        foreach ($r in $(if ($records.Count) { $records } else { $counts })) {
            if (-not $seen.Add($r.Key)) { continue }
            $local = [datetimeoffset]::Parse($r.Ts, $Inv).LocalDateTime
            if ($local -lt $Cutoff) { continue }
            $project = if ($r.Cwd) { Get-Leaf $r.Cwd } else { '(unknown)' }
            $uncached = [math]::Max(0L, $r.In - $r.Cached - $r.Write)
            Add-Usage $agg $project $local $session $(if ($r.Model) { $r.Model } else { '(unknown)' }) `
                $uncached $r.Out $r.Write $r.Cached $r.Reason
        }
    }
    $agg
}

function Update-History([string]$Tool) {
    $histDir  = Join-Path $DataFolder "$Tool/history"
    $histFile = Join-Path $histDir "usage-$Computer.csv"
    New-Item -ItemType Directory -Force -Path $histDir | Out-Null

    $full   = $FullScan -or -not (Test-Path $histFile)
    $cutoff = if ($full) { [datetime]::MinValue } else { (Get-Date).Date.AddDays(-$LookbackDays) }
    $agg    = if ($Tool -eq 'claude') { Read-ClaudeLogs $cutoff } else { Read-CodexLogs $cutoff }

    # Upsert. Logs only ever disappear, so a fresh total with fewer requests than the stored one
    # means part of that day's log has expired: keep the stored row.
    $history = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    if (Test-Path $histFile) {
        foreach ($r in Import-Csv $histFile) { $history["$($r.Project)|$($r.Date)|$($r.Session)|$($r.Model)"] = $r }
    }
    $added = 0; $updated = 0
    foreach ($k in $agg.Keys) {
        $new = $agg[$k]; $old = $history[$k]
        if (-not $old) { $history[$k] = $new; $added++ }
        elseif ($new.Requests -ge [long]$old.Requests) {
            if ($new.Requests -ne [long]$old.Requests -or $new.OutputTokens -ne [long]$old.OutputTokens) { $updated++ }
            $history[$k] = $new
        }
    }
    if ($history.Count) {
        $history.Values | Sort-Object Date, Project, Session, Model | Select-Object $HistCols |
            Export-Csv -Path $histFile -NoTypeInformation -Encoding utf8
    }
    Write-Host "$($Tools[$Tool].Name) history: $added new and $updated updated row(s) -> $histFile"
}

#endregion

#region Report -------------------------------------------------------------------------------

function Import-History([string]$Tool) {
    $dir = Join-Path $DataFolder "$Tool/history"
    if (-not (Test-Path $dir)) { return }
    foreach ($f in Get-ChildItem $dir -Filter 'usage-*.csv' -File) {
        foreach ($r in Import-Csv $f.FullName) {
            [pscustomobject]@{
                Tool = $Tool; Computer = $r.Computer; Project = $r.Project; Date = $r.Date; Week = $r.Week
                Session = $r.Session; Model = $r.Model
                Requests = [long]$r.Requests; InputTokens = [long]$r.InputTokens; OutputTokens = [long]$r.OutputTokens
                CacheWriteTokens = [long]$r.CacheWriteTokens; CacheReadTokens = [long]$r.CacheReadTokens
                ReasoningTokens = [long]$r.ReasoningTokens
            }
        }
    }
}

function Get-WeekFile([string]$View, [string]$Label, [string]$Ext) { "$View-usage-$Label.$Ext" }
function Get-WeekPath([string]$View, [string]$Label, [string]$Ext) {
    Join-Path $DataFolder "$View/weekly/$Label/$(Get-WeekFile $View $Label $Ext)"
}

# Columns shared by the project tables: a per-tool output split on the combined view, a reasoning
# column when the data has any.
function Get-SplitColumns([string[]]$ToolKeys, [object[]]$Rows) {
    $cols = @()
    if ($ToolKeys.Count -gt 1) {
        foreach ($k in $ToolKeys) { $cols += @{ H = $Tools[$k].Name; T = 'compact'; V = [scriptblock]::Create("param(`$r) `$r.ByTool['$k']") } }
    }
    if (($Rows | Measure-Object ReasoningTokens -Sum).Sum -gt 0) {
        $cols += @{ H = 'Reasoning'; T = 'compact'; V = { param($r) $r.ReasoningTokens } }
    }
    $cols
}

function Write-WeeklyReport([string]$View, [hashtable]$ByWeek, [string]$Label, [hashtable]$Span) {
    $v        = $Views[$View]
    $start    = Get-IsoWeekStart $Label
    $end      = $start.AddDays(6)
    $prev     = Get-IsoWeekLabel $start.AddDays(-7)
    $next     = Get-IsoWeekLabel $start.AddDays(7)
    $inFlight = (Get-Date).Date -le $end
    $wRows    = @($ByWeek[$Label] | Where-Object { $_ })
    $t        = Get-Totals $wRows
    $p        = Get-Totals @($ByWeek[$prev] | Where-Object { $_ })
    $csvPath  = Get-WeekPath $View $Label 'csv'
    New-Item -ItemType Directory -Force -Path (Split-Path $csvPath) | Out-Null

    # CSV: one row per tool + computer + project.
    $csvRows = foreach ($g in ($wRows | Group-Object Tool, Computer, Project)) {
        $gt = Get-Totals $g.Group; $f = $g.Group[0]
        [pscustomobject]@{
            Tool = $Tools[$f.Tool].Name; Computer = $f.Computer; Project = $f.Project; Week = $Label
            Sessions = $gt.Sessions; Requests = $gt.Requests; ActiveDays = $gt.ActiveDays
            InputTokens = $gt.InputTokens; OutputTokens = $gt.OutputTokens; ReasoningTokens = $gt.ReasoningTokens
            CacheWriteTokens = $gt.CacheWriteTokens; CacheReadTokens = $gt.CacheReadTokens; TotalTokens = $gt.TotalTokens
            Models = $gt.Models -join '; '
        }
    }
    if ($csvRows) { @($csvRows) | Sort-Object OutputTokens -Descending | Export-Csv $csvPath -NoTypeInformation -Encoding utf8 }
    else {
        '"Tool","Computer","Project","Week","Sessions","Requests","ActiveDays","InputTokens","OutputTokens","ReasoningTokens","CacheWriteTokens","CacheReadTokens","TotalTokens","Models"' |
            Set-Content $csvPath -Encoding utf8
    }

    # Daily chart, stacked by tool.
    $days  = 0..6 | ForEach-Object { $start.AddDays($_) }
    $byDay = @{}; foreach ($g in ($wRows | Group-Object Date)) { $byDay[$g.Name] = Get-Totals $g.Group }
    $dayTotals = [object[]]::new(7)
    for ($i = 0; $i -lt 7; $i++) { $dayTotals[$i] = $byDay[$days[$i].ToString('yyyy-MM-dd', $Inv)] }
    $series = @(foreach ($k in $v.Tools) {
        $vals = [double[]]::new(7)
        for ($i = 0; $i -lt 7; $i++) { if ($dayTotals[$i]) { $vals[$i] = $dayTotals[$i].ByTool[$k] } }
        @{ Series = $Tools[$k].Series; Values = $vals }
    })
    $tips = for ($i = 0; $i -lt 7; $i++) {
        $dt = $dayTotals[$i]; $name = $days[$i].ToString('dddd d MMM', $Inv)
        if (-not $dt) { "$name`nNo activity"; continue }
        $lines = @($name, "$(Format-N $dt.OutputTokens) output tokens")
        if ($v.Tools.Count -gt 1) { foreach ($k in $v.Tools) { $lines += "  $($Tools[$k].Name): $(Format-N $dt.ByTool[$k])" } }
        $lines += "$(Format-N $dt.Requests) requests · $($dt.Projects) project(s)"
        $lines -join "`n"
    }
    $chart = New-ColumnChart -Labels ($days | ForEach-Object { $_.ToString('ddd d', $Inv) }) -Series $series -Tips $tips -Height 220 `
        -AriaLabel "Output tokens per day, $Label"

    $projects = Group-Totals $wRows 'Project'
    $projCols = @(@{ H = 'Project'; T = 'text'; V = { param($r) $r.Name } }
                  @{ H = 'Output tokens'; T = 'share'; V = { param($r) $r.OutputTokens } }) +
                (Get-SplitColumns $v.Tools $wRows) +
                @(@{ H = 'Requests'; T = 'n'; V = { param($r) $r.Requests } }
                  @{ H = 'Sessions'; T = 'n'; V = { param($r) $r.Sessions } }
                  @{ H = 'Days'; T = 'n'; V = { param($r) $r.ActiveDays } }
                  @{ H = 'Cache read'; T = 'compact'; V = { param($r) $r.CacheReadTokens } }
                  @{ H = 'Total tokens'; T = 'compact'; V = { param($r) $r.TotalTokens } })
    if ($t.Computers.Count -gt 1) { $projCols += @{ H = 'Computers'; T = 'text'; V = { param($r) $r.Computers -join ', ' } } }
    if ($v.Tools.Count -eq 1) { $projCols += @{ H = 'Models'; T = 'wrap'; V = { param($r) $r.Models -join ', ' } } }

    $groupCols = { param($h, [switch]$Html) @(
        @{ H = $h; T = $(if ($Html) { 'html' } else { 'text' }); V = { param($r) $r.Name } }
        @{ H = 'Output tokens'; T = 'share'; V = { param($r) $r.OutputTokens } }
        @{ H = 'Requests'; T = 'n'; V = { param($r) $r.Requests } }
        @{ H = 'Sessions'; T = 'n'; V = { param($r) $r.Sessions } }
        @{ H = 'Projects'; T = 'n'; V = { param($r) $r.Projects } }) }

    $range    = '{0} – {1}' -f $start.ToString('d MMM', $Inv), $end.ToString('d MMM yyyy', $Inv)
    $status   = if ($inFlight) { ' · <b>week in progress</b>' } else { '' }
    $prevLink = if ($prev -ge $Span.First) { "<a href=`"../$prev/$(Get-WeekFile $View $prev 'html')`">&larr; $prev</a>" } else { '' }
    $nextLink = if ($next -le $Span.Last)  { "<a href=`"../$next/$(Get-WeekFile $View $next 'html')`">$next &rarr;</a>" } else { '' }
    $viewNav  = New-ViewNav $View {
        param($k) if ($Label -ge $script:Spans[$k].First -and $Label -le $script:Spans[$k].Last) { "../../../$k/weekly/$Label/$(Get-WeekFile $k $Label 'html')" } else { "../../../$k/lifetime.html" }
    }
    $where = if ($t.Computers) { $t.Computers -join ', ' } else { $Computer }

    $body = @"
$viewNav
<header>
  <div><h1>$(Encode $v.Title) usage · $Label</h1><p>$range$status · $(Encode $where)</p></div>
  <nav>$prevLink $nextLink <a href="../../lifetime.html">Lifetime stats</a> <a href="$(Get-WeekFile $View $Label 'csv')">CSV</a></nav>
</header>
$(New-Tiles @(
    @('Output tokens', (Format-Compact $t.OutputTokens), (Format-Delta $t.OutputTokens $p.OutputTokens)),
    @('Requests', (Format-N $t.Requests), (Format-Delta $t.Requests $p.Requests)),
    @('Sessions', (Format-N $t.Sessions), (Format-Delta $t.Sessions $p.Sessions)),
    @('Projects', (Format-N $t.Projects), (Format-Delta $t.Projects $p.Projects)),
    @('Active days', "$($t.ActiveDays) / 7", '<span class="hint">days with any request</span>'),
    @('Total tokens', (Format-Compact $t.TotalTokens), '<span class="hint">incl. cached input</span>')
))
<section class="card">
  <h2>Output tokens by day</h2>
  <p class="sub">Tokens generated each day. Hover a day for details.</p>
  $(New-Legend $v.Tools)
  $chart
</section>
$(if ($v.Tools.Count -gt 1 -and $wRows) { @"
<section class="card">
  <h2>By tool</h2>
  $(New-Table @(Group-Totals $wRows 'Tool' | ForEach-Object { $_.Name = '<i class="sw {0}"></i>{1}' -f $Tools[$_.Name].Series, $Tools[$_.Name].Name; $_ }) (& $groupCols 'Tool' -Html))
</section>
"@ })
<section class="card">
  <h2>Projects</h2>
  <p class="sub">Ranked by output tokens. Sessions spanning several projects count once in each.</p>
  $(if ($projects) { New-Table $projects $projCols } else { '<p class="empty">No activity was recorded this week.</p>' })
</section>
$(if ($wRows) { @"
<section class="card">
  <h2>Models</h2>
  $(New-Table (Group-Totals $wRows 'Model') (& $groupCols 'Model'))
</section>
"@ })
"@
    New-Page "$($v.Title) $Label" $View $body | Set-Content (Get-WeekPath $View $Label 'html') -Encoding utf8
}

function Write-LifetimeReport([string]$View, [object[]]$Rows, [hashtable]$ByWeek) {
    $v = $Views[$View]
    $dir = Join-Path $DataFolder $View
    $t = Get-Totals $Rows
    $currentWeek = Get-IsoWeekLabel (Get-Date).Date

    # Weekly series from the first tracked week to this week, zero-filled.
    $weeks = [System.Collections.Generic.List[object]]::new()
    if ($t.FirstDate) {
        $w = Get-IsoWeekStart (Get-IsoWeekLabel ([datetime]::ParseExact($t.FirstDate, 'yyyy-MM-dd', $Inv)))
        while ($w -le (Get-IsoWeekStart $currentWeek)) {
            $label = Get-IsoWeekLabel $w; $wt = Get-Totals @($ByWeek[$label] | Where-Object { $_ })
            $byTool = [ordered]@{}; foreach ($k in $v.Tools) { $byTool[$Tools[$k].Name] = $wt.ByTool[$k] }
            $weeks.Add([pscustomobject]@{
                Week = $label; Start = $w.ToString('yyyy-MM-dd', $Inv)
                OutputTokens = $wt.OutputTokens; TotalTokens = $wt.TotalTokens; Requests = $wt.Requests
                Sessions = $wt.Sessions; Projects = $wt.Projects; ActiveDays = $wt.ActiveDays; OutputByTool = $byTool
                Report = if (Test-Path (Get-WeekPath $View $label 'html')) { "weekly/$label/$(Get-WeekFile $View $label 'html')" } else { $null }
            })
            $w = $w.AddDays(7)
        }
    }

    $projects = Group-Totals $Rows 'Project'
    $top = @($projects | Select-Object -First $TopProjects)
    $days = @(foreach ($g in ($Rows | Group-Object Date)) { [pscustomobject]@{ Date = $g.Name; OutputTokens = (Get-Totals $g.Group).OutputTokens } })
    $busiestDay  = $days | Sort-Object OutputTokens -Descending | Select-Object -First 1
    $busiestWeek = $weeks | Sort-Object OutputTokens -Descending | Select-Object -First 1
    $streak = 0; $run = 0; $prevDay = $null
    foreach ($d in ($days.Date | Sort-Object)) {
        $cur = [datetime]::ParseExact($d, 'yyyy-MM-dd', $Inv)
        $run = if ($prevDay -and ($cur - $prevDay).Days -eq 1) { $run + 1 } else { 1 }
        $streak = [math]::Max($streak, $run); $prevDay = $cur
    }
    $models    = Group-Totals $Rows 'Model'
    $computers = Group-Totals $Rows 'Computer'
    $byTool    = Group-Totals $Rows 'Tool'
    $activeWeeks = @($weeks | Where-Object Requests).Count

    # JSON: the machine-readable lifetime summary.
    $slim = { param($o) [ordered]@{ Name = $o.Name; OutputTokens = $o.OutputTokens; ReasoningTokens = $o.ReasoningTokens
        InputTokens = $o.InputTokens; CacheWriteTokens = $o.CacheWriteTokens; CacheReadTokens = $o.CacheReadTokens
        TotalTokens = $o.TotalTokens; Requests = $o.Requests; Sessions = $o.Sessions; ActiveDays = $o.ActiveDays
        FirstDate = $o.FirstDate; LastDate = $o.LastDate } }
    $json = [ordered]@{
        Generated = (Get-Date).ToString('o', $Inv); View = $v.Title
        TrackingSince = $t.FirstDate; LastActivity = $t.LastDate
        Totals = [ordered]@{ OutputTokens = $t.OutputTokens; ReasoningTokens = $t.ReasoningTokens; InputTokens = $t.InputTokens
            CacheWriteTokens = $t.CacheWriteTokens; CacheReadTokens = $t.CacheReadTokens; TotalTokens = $t.TotalTokens
            Requests = $t.Requests; Sessions = $t.Sessions; Projects = $t.Projects; ActiveDays = $t.ActiveDays; ActiveWeeks = $activeWeeks }
        Records = [ordered]@{
            BusiestWeek = if ($busiestWeek) { [ordered]@{ Week = $busiestWeek.Week; OutputTokens = $busiestWeek.OutputTokens } }
            BusiestDay  = if ($busiestDay)  { [ordered]@{ Date = $busiestDay.Date; OutputTokens = $busiestDay.OutputTokens } }
            LongestDailyStreak = $streak }
        Tools       = @($byTool | ForEach-Object { $o = & $slim $_; $o.Name = $Tools[$_.Name].Name; $o })
        TopProjects = @($top | ForEach-Object {
            $o = & $slim $_; $o.Computers = $_.Computers; $o.Models = $_.Models
            if ($v.Tools.Count -gt 1) { $o.OutputByTool = [ordered]@{}; foreach ($k in $v.Tools) { $o.OutputByTool[$Tools[$k].Name] = $_.ByTool[$k] } }
            $o })
        Models    = @($models | ForEach-Object { & $slim $_ })
        Computers = @($computers | ForEach-Object { & $slim $_ })
        Weeks     = @($weeks | Select-Object Week, Start, OutputTokens, TotalTokens, Requests, Sessions, Projects, ActiveDays,
                        @{ n = 'OutputByTool'; e = { if ($v.Tools.Count -gt 1) { $_.OutputByTool } } }, Report)
    }
    $json | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $dir 'lifetime.json') -Encoding utf8

    # Weekly trend chart: last 52 weeks, stacked by tool.
    $chartWeeks = @($weeks | Select-Object -Last 52)
    $chart = if ($chartWeeks) {
        $series = @(foreach ($k in $v.Tools) { @{ Series = $Tools[$k].Series; Values = [double[]]@($chartWeeks | ForEach-Object { $_.OutputByTool[$Tools[$k].Name] }) } })
        $tips = foreach ($w in $chartWeeks) {
            $lines = @("$($w.Week) (from $($w.Start))", "$(Format-N $w.OutputTokens) output tokens")
            if ($v.Tools.Count -gt 1) { foreach ($name in $w.OutputByTool.Keys) { $lines += "  ${name}: $(Format-N $w.OutputByTool[$name])" } }
            $lines += "$(Format-N $w.Requests) requests · $($w.Projects) project(s)"
            $lines -join "`n"
        }
        New-ColumnChart -Labels ($chartWeeks | ForEach-Object { 'W' + $_.Week.Substring(6) }) -Series $series -Tips $tips `
            -LabelEvery ([math]::Max(1, [math]::Ceiling($chartWeeks.Count / 13))) -AriaLabel 'Output tokens per week'
    } else { '<p class="empty">No history yet.</p>' }

    $rank = 0
    $topCols = @(@{ H = '#'; T = 'rank'; V = { param($r) $r.Rank } }
                 @{ H = 'Project'; T = 'text'; V = { param($r) $r.Name } }
                 @{ H = 'Output tokens'; T = 'share'; V = { param($r) $r.OutputTokens } }) +
               (Get-SplitColumns $v.Tools $Rows) +
               @(@{ H = 'Total tokens'; T = 'compact'; V = { param($r) $r.TotalTokens } }
                 @{ H = 'Requests'; T = 'n'; V = { param($r) $r.Requests } }
                 @{ H = 'Sessions'; T = 'n'; V = { param($r) $r.Sessions } }
                 @{ H = 'Days'; T = 'n'; V = { param($r) $r.ActiveDays } }
                 @{ H = 'First seen'; T = 'text'; V = { param($r) $r.FirstDate } }
                 @{ H = 'Last seen'; T = 'text'; V = { param($r) $r.LastDate } })
    $top | ForEach-Object { $_ | Add-Member Rank (++$rank) }

    $groupCols = { param($h, [switch]$Html) @(
        @{ H = $h; T = $(if ($Html) { 'html' } else { 'text' }); V = { param($r) $r.Name } }
        @{ H = 'Output tokens'; T = 'share'; V = { param($r) $r.OutputTokens } }
        @{ H = 'Requests'; T = 'n'; V = { param($r) $r.Requests } }
        @{ H = 'Sessions'; T = 'n'; V = { param($r) $r.Sessions } }
        @{ H = 'Projects'; T = 'n'; V = { param($r) $r.Projects } }
        @{ H = 'Last used'; T = 'text'; V = { param($r) $r.LastDate } }) }

    $weekItems = @($weeks | Where-Object { $_.Requests -or $_.Report } | Sort-Object Week -Descending)
    $weekCols = @(@{ H = 'Week'; T = 'html'; V = { param($w)
                        $s = if ($w.Report) { "<a href=`"$($w.Report)`">$($w.Week)</a>" } else { $w.Week }
                        if ($w.Week -eq $currentWeek) { $s += ' <span class="delta">in progress</span>' }; $s } }
                  @{ H = 'Starts'; T = 'text'; V = { param($w) $w.Start } }
                  @{ H = 'Output tokens'; T = 'compact'; V = { param($w) $w.OutputTokens } })
    if ($v.Tools.Count -gt 1) {
        foreach ($k in $v.Tools) { $weekCols += @{ H = $Tools[$k].Name; T = 'compact'; V = [scriptblock]::Create("param(`$w) `$w.OutputByTool['$($Tools[$k].Name)']") } }
    }
    $weekCols += @(@{ H = 'Requests'; T = 'n'; V = { param($w) $w.Requests } }
                   @{ H = 'Sessions'; T = 'n'; V = { param($w) $w.Sessions } }
                   @{ H = 'Projects'; T = 'n'; V = { param($w) $w.Projects } }
                   @{ H = 'Days'; T = 'n'; V = { param($w) $w.ActiveDays } })

    $since  = if ($t.FirstDate) { [datetime]::ParseExact($t.FirstDate, 'yyyy-MM-dd', $Inv).ToString('d MMMM yyyy', $Inv) } else { 'no data yet' }
    $latest = $weeks | Where-Object Report | Select-Object -Last 1
    $viewNav = New-ViewNav $View { param($k) "../$k/lifetime.html" }
    $toolSince = if ($v.Tools.Count -gt 1) {
        ' · ' + (($v.Tools | ForEach-Object { $s = $script:Spans[$_]; if ($s.HasData) { "$($Tools[$_].Name) from $($s.FirstDate)" } }) -join ', ')
    }

    $body = @"
$viewNav
<header>
  <div><h1>$(Encode $v.Title) lifetime usage</h1><p>Tracking since $since$toolSince · $(Encode ($t.Computers -join ', '))</p></div>
  <nav>$(if ($latest) { "<a href=`"$($latest.Report)`">Latest weekly report ($($latest.Week))</a>" }) <a href="lifetime.json">JSON</a></nav>
</header>
$(New-Tiles @(
    @('Output tokens', (Format-Compact $t.OutputTokens), "<span class=`"hint`">$(Format-N $t.OutputTokens)</span>"),
    @('Total tokens', (Format-Compact $t.TotalTokens), '<span class="hint">incl. cached input</span>'),
    @('Requests', (Format-N $t.Requests), ''),
    @('Sessions', (Format-N $t.Sessions), ''),
    @('Projects', (Format-N $t.Projects), ''),
    @('Active days', (Format-N $t.ActiveDays), "<span class=`"hint`">across $activeWeeks active week(s)</span>")
))
<section class="card">
  <h2>Records</h2>
  <div class="records">
    <div><span>Busiest week</span><b>$(if ($busiestWeek -and $busiestWeek.OutputTokens) { "$($busiestWeek.Week) · $(Format-Compact $busiestWeek.OutputTokens)" } else { '–' })</b><em>output tokens</em></div>
    <div><span>Busiest day</span><b>$(if ($busiestDay) { "$($busiestDay.Date) · $(Format-Compact $busiestDay.OutputTokens)" } else { '–' })</b><em>output tokens</em></div>
    <div><span>Longest daily streak</span><b>$streak day(s)</b><em>consecutive days with activity</em></div>
    <div><span>Top project</span><b>$(if ($top) { Encode $top[0].Name } else { '–' })</b><em>$(if ($top -and $t.OutputTokens) { ($top[0].OutputTokens / $t.OutputTokens * 100).ToString('0', $Inv) + '% of all output tokens' })</em></div>
  </div>
</section>
<section class="card">
  <h2>Output tokens by week</h2>
  <p class="sub">$(if ($weeks.Count -gt 52) { 'Last 52 weeks.' } else { 'Every week since tracking began.' }) Hover a week for details.</p>
  $(New-Legend $v.Tools)
  $chart
</section>
$(if ($v.Tools.Count -gt 1 -and $Rows) { @"
<section class="card">
  <h2>By tool</h2>
  $(New-Table @($byTool | ForEach-Object { $_.Name = '<i class="sw {0}"></i>{1}' -f $Tools[$_.Name].Series, $Tools[$_.Name].Name; $_ }) (& $groupCols 'Tool' -Html))
</section>
"@ })
<section class="card">
  <h2>Top $TopProjects projects of all time</h2>
  <p class="sub">Ranked by output tokens. $(if ($projects.Count -gt $TopProjects) { "$($projects.Count - $TopProjects) more project(s) not shown." } else { "$($projects.Count) project(s) tracked." })</p>
  $(if ($top) { New-Table $top $topCols } else { '<p class="empty">No history yet.</p>' })
</section>
$(if ($models) { @"
<section class="card">
  <h2>Models</h2>
  $(New-Table $models (& $groupCols 'Model'))
</section>
"@ })
$(if ($computers.Count -gt 1) { @"
<section class="card">
  <h2>Computers</h2>
  $(New-Table $computers (& $groupCols 'Computer'))
</section>
"@ })
<section class="card">
  <h2>Weekly reports</h2>
  $(if ($weekItems) { New-Table $weekItems $weekCols } else { '<p class="empty">No weekly reports yet.</p>' })
</section>
"@
    New-Page "$($v.Title) lifetime usage" $View $body | Set-Content (Join-Path $dir 'lifetime.html') -Encoding utf8
    Write-Host "$($v.Title): $(Join-Path $dir 'lifetime.html')"
}

#endregion

#region Main ---------------------------------------------------------------------------------

New-Item -ItemType Directory -Force -Path (Join-Path $DataFolder 'logs') | Out-Null
$logFile = Join-Path $DataFolder ("logs/run-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $logFile | Out-Null
try {
    if (-not $ReportOnly) { foreach ($k in $Tools.Keys) { Update-History $k } }
    if (-not $CollectOnly) {
        $rowsByTool = @{}; foreach ($k in $Tools.Keys) { $rowsByTool[$k] = @(Import-History $k) }
        $currentWeek = Get-IsoWeekLabel (Get-Date).Date
        $target = if ($Week) { $Week } else { Get-IsoWeekLabel (Get-Date).Date.AddDays(-7) }

        # Which weeks each view reports on: from its first recorded week to the target week (or the
        # latest report already written, if later).
        $script:Spans = @{}
        foreach ($view in $Views.Keys) {
            $rows = @($Views[$view].Tools | ForEach-Object { $rowsByTool[$_] })
            $first = ($rows.Week | Sort-Object | Select-Object -First 1) ?? $target
            if ($target -lt $first) { $first = $target }
            $existing = Get-ChildItem (Join-Path $DataFolder "$view/weekly") -Directory -ErrorAction SilentlyContinue |
                Where-Object Name -match '^\d{4}-W\d{2}$' | Sort-Object Name | Select-Object -Last 1 -ExpandProperty Name
            $last = @($target, $existing) | Where-Object { $_ } | Sort-Object | Select-Object -Last 1
            $script:Spans[$view] = @{ First = $first; Last = $last; HasData = [bool]$rows; FirstDate = ($rows.Date | Sort-Object | Select-Object -First 1) }
        }

        foreach ($view in $Views.Keys) {
            New-Item -ItemType Directory -Force -Path (Join-Path $DataFolder $view) | Out-Null
            $rows = @($Views[$view].Tools | ForEach-Object { $rowsByTool[$_] })
            $byWeek = @{}; foreach ($g in ($rows | Group-Object Week)) { $byWeek[$g.Name] = $g.Group }

            # Write the target week and the one before it (so its "next" link appears), plus any
            # finished week with no final report yet: none, or one written while still in progress.
            $todo = [System.Collections.Generic.List[string]]::new()
            $prevOfTarget = Get-IsoWeekLabel (Get-IsoWeekStart $target).AddDays(-7)
            for ($w = Get-IsoWeekStart $script:Spans[$view].First; $w -le (Get-IsoWeekStart $target); $w = $w.AddDays(7)) {
                $label = Get-IsoWeekLabel $w
                $file = Get-WeekPath $view $label 'html'
                $stale = $label -lt $currentWeek -and (-not (Test-Path $file) -or (Get-Item $file).LastWriteTime -lt $w.AddDays(7))
                if ($stale -or $label -eq $target -or $label -eq $prevOfTarget) { $todo.Add($label) }
            }
            foreach ($label in $todo) { Write-WeeklyReport $view $byWeek $label $script:Spans[$view] }
            Write-Host "$($Views[$view].Title): wrote $($todo.Count) weekly report(s), $($todo[0]) to $($todo[-1])"
            Write-LifetimeReport $view $rows $byWeek
        }

        # Landing page: the combined lifetime report.
        '<!doctype html><meta charset="utf-8"><meta http-equiv="refresh" content="0; url=combined/lifetime.html"><title>AI coding usage</title><a href="combined/lifetime.html">Combined usage report</a>' |
            Set-Content (Join-Path $DataFolder 'index.html') -Encoding utf8
    }
}
finally {
    Stop-Transcript | Out-Null
    Get-ChildItem (Join-Path $DataFolder 'logs') -Filter 'run-*.log' | Sort-Object Name -Descending | Select-Object -Skip 26 | Remove-Item
}

#endregion
