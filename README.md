# Stock Dashboard

Local stock-summary dashboard. No backend, no build step — just PowerShell + a static HTML/JS template.

## Files

- `Update-StockDashboard.ps1` — fetches live quotes (Yahoo Finance chart API, no key) and recent
  headlines (Google News RSS, no key) for the tickers listed near the top of the script, then
  renders `template.html` into `dashboard.html`.
- `template.html` — the dashboard UI: per-card 1M/3M/6M/1Y range toggle, sparkline with hover
  tooltip, a "최근 이슈" news list (real article titles/links, not AI-written summaries).
- `run.bat` — double-click launcher (bypasses PowerShell execution-policy prompts).
- `dashboard.html` — generated output, opened automatically after each run. Not tracked in git.

## Usage

Double-click `run.bat`, or:

```powershell
powershell -ExecutionPolicy Bypass -File Update-StockDashboard.ps1
```

To change tracked tickers, edit the `$tickers` array at the top of `Update-StockDashboard.ps1`.

## Notes

- Requires only Windows PowerShell 5.1 — no Python/Node/npm.
- `.ps1` files must stay saved as **UTF-8 with BOM**, or Windows PowerShell 5.1 misreads the
  Korean text and the ₩ sign and fails to parse the script.
- Email automation (scheduled send) is not implemented yet — deferred until SMTP account and
  send time are decided.
