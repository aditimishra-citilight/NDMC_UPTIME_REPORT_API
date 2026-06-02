# NDMC Uptime Report — Run Guide

Generates `NDMC_UptimeReport_<Month><Year>.xlsx` with 6 zone sheets (SP, City, Civil_Lines, Karol_Bagh, Narela, Rohini) populated from the Citilight smartlight portal.

> **No more cookies.** The tool now logs in by itself using your portal username and
> password — you never touch the browser or DevTools. Just run it and answer 3 questions.

## Prerequisites (one-time)

- **Node.js** installed (check with `node --version`). Download: https://nodejs.org
- This folder (`D:\CityReport`) with all the JS files. (`run-report.bat` installs the
  `node_modules` automatically the first time.)
- A **portal login** (username + password) for `https://smartlight.citilight.co:446`.

---

## The easy way — double-click

1. In `D:\CityReport`, **double-click `run-report.bat`**.
2. A black window opens and asks 3 things — type the answer and press Enter:
   ```
   Which month do you want the report for? Enter 1-12 (1=Jan … 12=Dec): 5
   Which year? Press Enter for 2026:
   Portal username: admin
   Portal password: ********        (hidden as you type)
   ```
3. It logs in, fetches all 6 zones (~10–15 min), and **opens the Excel automatically**.
4. The window stays open at the end so you can read the summary. Close it when done.

That's it. The report `NDMC_UptimeReport_<Month><Year>.xlsx` is saved in `D:\CityReport`.

> **Tip:** If a file with that name is already open in Excel, close it first — otherwise
> the tool saves a `_NEW.xlsx` copy and tells you to rename it.

---

## The manual way — PowerShell (optional)

If you prefer the command line instead of the `.bat`:

```powershell
cd D:\CityReport
node NdmcUptimeReport.js
```

It asks the same 3 questions. To skip the prompts (e.g. for automation), set environment
variables first:

```powershell
$env:NDMC_MONTH = 5
$env:NDMC_YEAR  = 2026
$env:NDMC_USER  = "admin"
$env:NDMC_PASS  = "your-password"
node NdmcUptimeReport.js
```

---

## What you'll see while it runs

Live progress prints one zone at a time, with chunked API calls visible:
```
Report: May-26   (2026-05-01 → 2026-05-31)
Logging in…
Login OK — session acquired.

[SP] cityId=2
  live data: 133 switches
    op chunk 2026-05-01 → 2026-05-04
    ...
  uptime: 1995 daily rows (0 failed chunks)
```

When complete, look for the `--- Summary ---` block with all six zones showing `OK`, and
the final line:
```
Written: NDMC_UptimeReport_May2026.xlsx
```

---

## Approximate run times

The script chunks long date ranges into short windows and fetches them concurrently
(up to `MAX_CONCURRENCY` requests at once — default 3, the "safe" level for this portal).
Operational and uptime are fetched in parallel within each zone.

| Date range | Approximate time |
|---|---|
| 4 days (test) | 2–3 minutes |
| 15 days | 5–8 minutes |
| Full month (29–31 days) | 10–15 minutes |

**Speed knob:** `MAX_CONCURRENCY` near the top of the script (default `3`). If a full run
finishes with **0 failed chunks**, you can try raising it to `5`. If you see `failed
chunks` or `FAIL: 502`, lower it back to `3`.

**While it runs:** don't close the window, and don't open the Excel file until it's done.

---

## Verification checklist (after each run)

Open the output Excel. On each sheet (or just SP — the others use the same logic), confirm:

- Row 1: title `Monthly uptime and Energy Consumption Report <Zone> Zone - <Month><Year>`.
- Row 2: 12 column headers (A: SNo. … L: Actual kWh %), frozen so they stay visible.
- Row 3+: one row per switch, sorted by Switch ID alphabetically (A* first, then H*).
- Column B: switch IDs like `CCMS A009339`, never blank.
- Column C: month label (e.g. `May-26`) in every row.
- Column D: values > 0 (zeros are auto-replaced with random 0.10–0.80).
- Column E: same value in every row of a sheet.
- Column F = E − G (spot check with `=E3-G3`).
- Column G: mostly 0; ~5% of rows have a small value, all ≤ ~1.50.
- Column H: always `0.00`.
- Column I = F / E formatted as % (typically 97–100%).
- Column L = K / J formatted as % (often 100% when uptime is 100%).

Switch counts per zone (May 2026 baseline, ±2):
SP 133, City 143, Civil_Lines 823, Karol_Bagh 341, Narela 884, Rohini 962.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `Login failed or session invalid — check the username/password` | Wrong username/password | Re-run and re-type credentials carefully |
| `login: server did not return a JSESSIONID cookie` | Login endpoint/portal changed or unreachable | Check the portal loads in a browser; confirm credentials |
| Zone says `FAIL: 502` | Portal server overloaded | Wait 5 min, retry. Persistent 502s mean portal is down. |
| Zone says `FAIL: timeout` | API call took >5 min | Chunking should prevent this. If it persists, retry. |
| `EBUSY: resource busy or locked` | Excel has the file open | Close Excel, re-run (it auto-saves a `_NEW.xlsx` otherwise) |
| `ETIMEDOUT` / `ECONNREFUSED` | Portal down or no internet | Check portal in browser; wait and retry |
| Zone shows `0 switches OK` | API returned empty | Retry; check if cityId still valid |
| `Cannot parse CCMS ID` | Unexpected switch ID format | Inspect that switch on the portal |
| `Node.js is not installed` (from the .bat) | Node missing | Install from https://nodejs.org and re-run |
| `Cannot find module 'axios'`/`exceljs` | Dependencies missing | `cd D:\CityReport; npm install` (the .bat does this automatically) |

---

## Column source reference

| Col | Header | Source |
|---|---|---|
| A | SNo. | Sequence number |
| B | Switch ID | Live Data Feed → `name` field |
| C | Month | Constant (e.g. `May-26`) |
| D | Connected Load (KW) | Live Data Feed → `totalwattage` field (0 → random 0.10–0.80) |
| E | Night Duration / Expected Hours | Operational report → mean of per-switch `expected_on` sums |
| F | Actual Hours of Lamps Operated | Formula: `E − G` |
| G | Lamps OFF due to Power Failure (Hours) | Operational report → sum of `output_off` (>2.0 → random 0.80–1.50; plus ~5% scatter) |
| H | Lamps OFF due to Abnormalities (Hours) | Always `0.00` |
| I | Load Uptime % | Formula: `F / E` |
| J | Desired kWh Consumption | Uptime report → per-day `expected_kwh` rounded to 2dp, then summed per switch |
| K | Actual kWh Consumption | Uptime report → per-day `actual_kwh` rounded to 2dp, then summed per switch |
| L | Actual kWh Consumption % | Formula: `K / J` |

---

## Implementation notes

- **Auto-login:** the script POSTs username/password to `/smartlight/login`, captures the
  `JSESSIONID`, and uses it for all data calls. Credentials are typed at runtime (or read
  from `NDMC_USER`/`NDMC_PASS`) and are never stored in the code.
- The script sorts switches alphabetically by Switch ID (A* first, then H*) to match historical NDMC reports.
- Connected Load comes from the `totalwattage` JSON field (a near-live meter reading), not from rated power.
- Operational and Uptime API calls are chunked into short (4-day) windows because the portal cannot handle large date ranges in a single call (it times out or returns 502).
- Chunks are fetched **concurrently**, throttled by a global semaphore to at most `MAX_CONCURRENCY` (default 3) in-flight requests at any moment.
- Operational and uptime for a zone run **in parallel** (they're independent), sharing the same concurrency budget.
- Substitution rules:
  - Connected Load = 0 → random 0.10–0.80 (per `connected-load-zero-rule`).
  - Power Failure hours > 2.0, plus ~1 in every 15–20 switches → random 0.80–1.50 (per `power-failure-substitution-rule`).
- Output is written with **ExcelJS** (styled: merged title, bold wrapped headers, borders, frozen header, 2-decimal number formats).
