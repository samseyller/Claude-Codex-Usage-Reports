![Claude Code + Codex Usage Reports](banner.png)

# Claude Code + Codex Usage Reports

PowerShell scripts that read the local session logs of **Claude Code** and **Codex**, keep a lifetime token-usage history, and write weekly and lifetime CSV/HTML reports for each tool and for both combined.

## Requirements

- PowerShell 7+
- Windows (only for the scheduled task)

## Usage

```powershell
# Collect usage and write reports for the last completed ISO week
./Invoke-UsageReport.ps1

# A specific (or in-progress) week
./Invoke-UsageReport.ps1 -Week 2026-W42

# Run automatically every Monday at 09:00
./Register-UsageReportTask.ps1
```

Open `data/index.html` to browse the reports.

| Parameter | Description |
| --- | --- |
| `-DataFolder` | Where history and reports are stored (default `./data`) |
| `-Week` | ISO week to report, e.g. `2026-W42` |
| `-LookbackDays` | Days of logs to rescan on each run (default `35`) |
| `-TopProjects` | Projects shown in the lifetime report (default `25`) |
| `-FullScan` | Rescan all logs, not just the lookback window |
| `-CollectOnly` / `-ReportOnly` | Run only one of the two steps |

## Multiple computers

Point `-DataFolder` at a synced folder (OneDrive etc.) on every machine. Each computer writes its own `usage-<computer>.csv` history file and the reports merge them all. Secondary machines can use `-CollectOnly`:

```powershell
./Register-UsageReportTask.ps1 -DataFolder "$HOME/OneDrive/AIUsage" -CollectOnly
```

## Output

```
data/
├── index.html
├── claude/                      Claude Code only
│   ├── history/                 per-computer history
│   ├── weekly/<week>/*.csv|html
│   └── lifetime.json|html
├── codex/                       Codex only (same layout)
└── combined/                    both tools (weekly + lifetime only)
```

The history files are the only lasting record, because Claude Code deletes its logs after about 30 days. Back them up (for example in a synced folder). Everything else is regenerated on each run. The whole `data/` folder is git-ignored because it holds your personal usage data, so don't commit it.

## License

[MIT](LICENSE)
