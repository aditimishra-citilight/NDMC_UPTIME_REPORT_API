# NDMC Uptime Report — Run Guide

Generates `NDMC_UptimeReport_<Month><Year>.xlsx` with 6 zone sheets (SP, City, Civil_Lines, Karol_Bagh, Narela, Rohini) populated from the Citilight smartlight portal.

## Prerequisites (one-time)

- Node.js installed (check with `node --version` in PowerShell).
- This folder (`D:\CityReport`) has all the JS files and `node_modules`.
- A login on `https://smartlight.citilight.co:446`.

---

## Step-by-step monthly run

### Step 1 — Get a fresh JSESSIONID cookie

Sessions expire after a few hours, so always get a fresh cookie right before running.

1. Open Chrome (or Edge). Go to `https://smartlight.citilight.co:446` and log in. (Accept the "not private" warning if asked.)
2. In a new tab, open `https://smartlight.citilight.co:446/smartlight/livedatafeed`. Pick any city. Click **"Get Data"** (or "Show List"). Wait for the switch table.
3. Press **F12** → click **Network** tab.
4. In the Network filter bar, type `getListViewData`.
5. Click the `getListViewData_v1` row.
6. On the right panel, **Headers** tab → scroll to **Request Headers** → find the `Cookie:` line.
7. Copy the 32-character hex value that comes immediately after `JSESSIONID=` and ends at the next `;` (no `JSESSIONID=`, no semicolons, no spaces).

**Critical:** Cookie must be from `smartlight.citilight.co:446`. NOT `dc.citilight.co`. They look identical but the dc one will silently fail.

### Step 2 — Open the script

`D:\CityReport\NdmcUptimeReport.js` — open in Notepad, VS Code, or any text editor.

### Step 3 — Edit three or four lines at the top

```js
const REPORT_YEAR  = 2026;     // line 9  — year of the report
const REPORT_MONTH = 5;        // line 10 — month (1=Jan, 5=May, 12=Dec)
...
const CUSTOM_START_DATE = "";  // line 14 — leave "" for full month
const CUSTOM_END_DATE   = "";  // line 15 — leave "" for full month
const JSESSIONID = "PASTE_YOUR_COOKIE_HERE";   // line 17
```

For a **custom date range** (e.g., May 1–29), use:
```js
const CUSTOM_START_DATE = "2026-05-01";
const CUSTOM_END_DATE   = "2026-05-29";
```

### Step 4 — Save the file (Ctrl+S)

### Step 5 — Delete the old Excel if it exists

In File Explorer at `D:\CityReport`, find `NDMC_UptimeReport_<Month><Year>.xlsx` and delete it. If Windows says "file is open in Microsoft Excel" → close Excel fully (Task Manager → End Task on Excel) → retry delete.

### Step 6 — Run in PowerShell

Open PowerShell (Start menu → "PowerShell" → Enter), then:

```powershell
cd D:\CityReport
node NdmcUptimeReport.js
```

### Step 7 — Wait for the summary

Live progress prints one zone at a time, with chunked API calls visible:
```
[SP] cityId=2
  live data: 133 switches
    op chunk 2026-05-01 → 2026-05-07
    op chunk 2026-05-08 → 2026-05-14
    ...
  operational: 88482 daily rows
    up chunk 2026-05-01 → 2026-05-07
    ...
  uptime: 665 daily rows
```

When complete, look for the `--- Summary ---` block with all six zones showing `OK`, and the final line:
```
Written: NDMC_UptimeReport_<Month><Year>.xlsx
```

### Step 8 — Open the output

In PowerShell:
```powershell
start D:\CityReport\NDMC_UptimeReport_May2026.xlsx
```

Or via File Explorer → `D:\CityReport` → double-click. Six tabs at the bottom (one per zone).

---

## Approximate run times

The script chunks long date ranges into short windows and fetches them concurrently
(up to `MAX_CONCURRENCY` requests at once — default 3, the "safe" level for this
portal). Operational and uptime are fetched in parallel within each zone.

| Date range | Approximate time |
|---|---|
| 4 days (test) | 2–3 minutes |
| 15 days | 5–8 minutes |
| Full month (29–31 days) | 10–15 minutes |

**Speed knob:** `MAX_CONCURRENCY` near the top of the script (default `3`). It caps
total simultaneous requests to the portal across all zones/chunks. If a full run
finishes with **0 failed chunks** in the summary, you can try raising it to `5` for a
faster run. If you start seeing `failed chunks` or `FAIL: 502`, lower it back to `3`.

**Rules while running:**
- Don't close PowerShell.
- Don't close the browser tab with the portal (keeps session warm).
- Don't open the Excel file.
- Don't press Ctrl+C.

---

## Verification checklist (after each run)

Open the output Excel. On each sheet (or just SP — the others use the same logic), confirm:

- Row 1: title `Monthly uptime and Energy Consumption Report <Zone> Zone - <Month><Year>`.
- Row 2: 12 column headers (A: SNo. … L: Actual kWh %).
- Row 3+: one row per switch, sorted by Switch ID alphabetically (A* first, then H*).
- Column B: switch IDs like `CCMS A009339`, never blank.
- Column C: month label (e.g. `May-26`) in every row.
- Column D: values > 0 (zeros are auto-replaced with random 0.10–0.80).
- Column E: same value in every row of a sheet.
- Column F = E − G (spot check with `=E3-G3`).
- Column G: maximum ≤ 2.0 (high values auto-replaced with random 0.80–1.50).
- Column I = F / E formatted as % (typically 97–100%).
- Column L = K / J formatted as % (often 100% when uptime is 100%).

Switch counts per zone (May 2026 baseline, ±2):
SP 133, City 143, Civil_Lines 823, Karol_Bagh 341, Narela 884, Rohini 962.

---

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `(liveData \|\| []).map is not a function` for every zone | Cookie expired or wrong-domain cookie | Repeat Step 1 with the smartlight portal (not dc.citilight.co) |
| Zone says `FAIL: 502` | Portal server overloaded | Wait 5 min, retry. Persistent 502s mean portal is down. |
| Zone says `FAIL: timeout` | API call took >5 min | Chunking should prevent this. If it still happens, reduce date range. |
| Some zones OK, others fail mid-run | Session expired during run | Get fresh cookie, re-run |
| `EBUSY: resource busy or locked` | Excel has the file open | Close Excel (Task Manager → End Task), re-run |
| `ETIMEDOUT` / `ECONNREFUSED` | Portal down or no internet | Check portal in browser; wait and retry |
| Zone shows `0 switches OK` | API returned empty | Retry; check if cityId still valid |
| `Cannot parse CCMS ID` | Unexpected switch ID format | Inspect that switch on the portal |
| `Cannot find module 'axios'` | Dependencies missing | `cd D:\CityReport; npm install` |

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
| G | Lamps OFF due to Power Failure (Hours) | Operational report → sum of `output_off` (>2.0 → random 0.80–1.50) |
| H | Lamps OFF due to Abnormalities (Hours) | Operational report → sum of `actual_off_seconds` |
| I | Load Uptime % | Formula: `F / E` |
| J | Desired kWh Consumption | Uptime report → sum of `expected_kwh` per switch |
| K | Actual kWh Consumption | Uptime report → sum of `actual_kwh` per switch |
| L | Actual kWh Consumption % | Formula: `K / J` |

---

## Implementation notes

- The script sorts switches alphabetically by Switch ID (A* first, then H*) to match historical NDMC reports.
- Connected Load comes from the `totalwattage` JSON field, not from rated power.
- Operational and Uptime API calls are chunked into short (4-day) windows because the portal cannot handle large date ranges in a single call (it times out or returns 502).
- Chunks are fetched **concurrently**, throttled by a global semaphore to at most `MAX_CONCURRENCY` (default 3) in-flight requests at any moment. Memory stays bounded — at most 3 chunk responses are held at once, each aggregated then freed.
- Operational and uptime for a zone run **in parallel** (they're independent), sharing the same concurrency budget.
- All daily rows from chunks are aggregated per switch into Maps as they arrive (no raw-row accumulation).
- Substitution rules:
  - Connected Load = 0 → random 0.10–0.80 (per zone's `connected-load-zero-rule`).
  - Power Failure hours > 2.0 → random 0.80–1.50 (per `power-failure-substitution-rule`).
