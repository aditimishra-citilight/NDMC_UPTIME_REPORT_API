# NDMC Uptime Report — Project Overview & Flow

Everything about how the **NDMC Monthly Uptime & Energy Consumption Report** is built:
the tech used, the libraries, where the data comes from, the full processing flow, and
the exact commands to produce the Excel file. This is the single "read me everything"
reference. For the click-by-click monthly run, also see [`RUN_GUIDE.md`](RUN_GUIDE.md).

---

## 1. What this project does

It replaces a slow manual process (downloading per-zone Excel exports from a web portal,
then building pivot tables and VLOOKUPs by hand) with **one command** that:

1. Logs into the **Citilight smartlight portal** APIs using a session cookie.
2. Pulls raw data for all **6 NDMC zones** (SP, City, Civil_Lines, Karol_Bagh, Narela, Rohini).
3. Aggregates it per street-light switch (CCMS).
4. Computes the report columns and applies the NDMC business rules.
5. Writes a fully **formatted** Excel workbook — `NDMC_UptimeReport_<Month><Year>.xlsx` —
   with one styled sheet per zone.

Output example: `NDMC_UptimeReport_May2026.xlsx`.

---

## 2. Tech stack

| Layer | Technology | Why |
|---|---|---|
| Runtime | **Node.js** (JavaScript, CommonJS modules) | Simple scripting, easy HTTP + Excel libraries |
| HTTP client | **axios** | POST requests to the portal APIs, with timeouts and retry |
| Excel writer | **ExcelJS** | Writes `.xlsx` **with full styling** (merged title, borders, alignment, frozen header, number formats). The free SheetJS (`xlsx`) cannot write styles, so we switched to ExcelJS. |
| TLS/HTTP agent | **https** (Node built-in) | Custom `https.Agent` for the portal's `:446` HTTPS endpoint |
| Verify helper | **xlsx** (SheetJS) | Read-only sanity checks in `verify.js` (reading doesn't need styling) |

### Libraries (npm dependencies)

From `package.json`:

```json
"dependencies": {
  "axios":   "^1.13.6",   // HTTP requests to the portal
  "exceljs": "^4.4.0",    // write the styled .xlsx report
  "xlsx":    "^0.18.5"    // read-back verification only (verify.js)
}
```

Install everything with:

```powershell
npm install
```

---

## 3. Where the data comes from — the portal APIs

Base URL: `https://smartlight.citilight.co:446`
Auth: a **`JSESSIONID`** session cookie (obtained from the browser after logging in).

> ⚠️ The cookie **must** come from `smartlight.citilight.co:446` — **not** `dc.citilight.co`.
> They look identical but a `dc` cookie fails silently (every zone errors with
> `.map is not a function`).

Three endpoints are used (all `POST`, JSON body):

| # | Endpoint | Gives us | Used for columns |
|---|---|---|---|
| 1 | `/smartlight/getListViewData_v1` (Live Data Feed) | switch list, names, connected load | B (Switch ID), D (Connected Load) |
| 2 | `/VELOCITi_API/api/ccmsOperationalreportallData` (Operational Report) | per-switch nightly expected/off hours | E (Night hours), G (Power Failure) |
| 3 | `/smartlight/getUptimeReport` (Uptime Report) | per-switch daily expected & actual kWh | J (Desired kWh), K (Actual kWh) |

Long date ranges are split into **4-day chunks** because the portal times out / returns
502 on large ranges.

---

## 4. Processing flow (end to end)

```
                         ┌─────────────────────────────────────────┐
                         │  CONFIG (top of NdmcUptimeReport.js)      │
                         │  REPORT_YEAR / REPORT_MONTH / dates       │
                         │  JSESSIONID cookie                        │
                         └───────────────────┬───────────────────────┘
                                             │
              for each of the 6 zones ───────┤
                                             ▼
   ┌──────────────────┐   ┌──────────────────────────────┐   ┌──────────────────────┐
   │ 1. Live Data Feed│   │ 2. Operational Report        │   │ 3. Uptime Report     │
   │  getListViewData │   │  ccmsOperationalreportallData │   │  getUptimeReport     │
   │  → switches,     │   │  (4-day chunks, concurrent)   │   │  (4-day chunks)      │
   │    connected load│   │  → expected_on, output_off    │   │  → expected/actual   │
   │                  │   │                               │   │     kWh per day      │
   └────────┬─────────┘   └───────────────┬──────────────┘   └──────────┬───────────┘
            │                  (op + uptime fetched IN PARALLEL per zone)│
            │                             │                              │
            ▼                             ▼                              ▼
   ┌──────────────────────────────────────────────────────────────────────────────┐
   │ AGGREGATE per switch (CCMS):                                                    │
   │  • dedupe operational rows per (switch, day)  — API repeats each ~168×          │
   │  • sum expected_on → E hours ;  sum output_off → G hours                        │
   │  • round each daily kWh to 2dp THEN sum → J (expected) , K (actual)             │
   └───────────────────────────────────┬────────────────────────────────────────────┘
                                        ▼
   ┌──────────────────────────────────────────────────────────────────────────────┐
   │ COMPUTE columns + apply business rules (see §5)                                 │
   │  D zero → random ;  G high/scatter → random ;  H = 0 ;  F = E−G ;  I = F/E ;     │
   │  L = K/J                                                                         │
   └───────────────────────────────────┬────────────────────────────────────────────┘
                                        ▼
   ┌──────────────────────────────────────────────────────────────────────────────┐
   │ WRITE styled sheet via ExcelJS (title, headers, borders, formats — see §6)      │
   └───────────────────────────────────┬────────────────────────────────────────────┘
                                        ▼
                    NDMC_UptimeReport_<Month><Year>.xlsx  (6 sheets)
```

Concurrency: a **global semaphore** (`MAX_CONCURRENCY`, default `3`) caps total
simultaneous portal requests across all zones/chunks/endpoints. A full month runs in
~10–15 minutes.

---

## 5. Column reference (A–L) — source, formula & rules

Each zone sheet has a merged title row, a header row, then one row per switch.

| Col | Header | How it's produced |
|---|---|---|
| A | SNo. | Sequence number (1, 2, 3 …) |
| B | Switch ID | Live Data Feed → `name` (e.g. `CCMS A009339`); sorted A* then H* |
| C | Month | Constant label, e.g. `May-26` |
| D | Connected Load (KW) | Live Data Feed → `totalwattage`. **Rule:** if 0 → random `0.10–0.80` |
| E | Night Duration / Estimated time of Operation (Hours) | Operational → mean of per-switch `expected_on` sums (same value for all rows in a zone) |
| F | Actual Hours of Lamps Operated (Hours) | **Formula:** `E − G` |
| G | Lamps OFF due to Power Failure (Hours) | Operational → sum of `output_off`. **Rules:** if `> 2.0` → random `0.80–1.50`; also ~1 in every 15–20 switches (≈5%) gets a random `0.80–1.50` so the report shows realistic minor outage; the rest are 0 |
| H | Lamps OFF due to Abnormalities (Hours) | **Always `0.00`** (per NDMC reporting) |
| I | Load Uptime Percentage by Operating Hours | **Formula:** `F / E` (shown as %) |
| J | Desired kWh Consumption | Uptime → **round each daily `expected_kwh` to 2dp, then sum** per switch |
| K | Actual kWh Consumption | Uptime → **round each daily `actual_kwh` to 2dp, then sum** per switch |
| L | Actual kWh Consumption Percentage | **Formula:** `K / J` (shown as %) |

> **kWh rounding note:** the manual process builds a PivotTable that sums the *exported*
> daily values (already 2 decimals). To match it exactly, the script rounds each day to
> 2 decimals **before** summing — otherwise summing raw API floats drifts by a few
> hundredths.

### Business / substitution rules (config at the top of the script)

```js
const CONNECTED_LOAD_ZERO_RANGE    = [0.1,  0.80];  // D: 0 → random in this range
const POWER_FAILURE_HIGH_THRESHOLD = 2.0;           // G: above this is "too high"
const POWER_FAILURE_HIGH_RANGE     = [0.80, 1.50];  // G: replacement / scatter range
const POWER_FAILURE_SCATTER_GAP    = [15, 20];      // G: ~1 hit every 15–20 rows (~5%)
const MAX_CONCURRENCY              = 3;              // safe in-flight request cap
```

---

## 6. Excel formatting (ExcelJS)

The output matches the manual NDMC layout exactly:

- **Title row** (row 1): text merged across all 12 columns, **bold, centered**, bordered.
- **Header row** (row 2): **bold, centered, wrapped text**, height 95px so long headers
  (e.g. *"Night Duration/ Estimated time of Operation (Hours)"*) are never clipped.
- **Frozen header:** rows 1–2 are frozen; the scrolling body is anchored at `A3`
  (`topLeftCell` + `activeCell`) so the header is **not** duplicated at the top when
  scrolling.
- **Alignment:** SNo. & Month centered, Switch ID left, all numbers right.
- **Number formats:** all numeric columns show **2 decimals** (`0.00`); columns I and L
  are **2-decimal percentages** (`0.00%`).
- **Borders:** thin black grid on every cell.

These are defined by `COLUMN_SPEC`, `HEADER_ROW_HEIGHT`, and `buildSheet()` in
`NdmcUptimeReport.js`.

---

## 7. How to get the Excel (step-by-step)

### Prerequisites (one time)
- **Node.js** installed (`node --version` to check).
- In `D:\CityReport`, run `npm install` once to install dependencies.
- A login on `https://smartlight.citilight.co:446`.

### Each run
1. **Get a fresh `JSESSIONID` cookie** (sessions expire in a few hours):
   - Log into `https://smartlight.citilight.co:446` in Chrome/Edge.
   - Open `…/smartlight/livedatafeed`, pick a city, click **Get Data**.
   - Press **F12 → Network**, filter `getListViewData`, click the
     `getListViewData_v1` request.
   - In **Headers → Request Headers**, copy the 32-char hex value after
     `JSESSIONID=` (no `JSESSIONID=`, no `;`, no spaces).
   - *(Full screenshots in [`RUN_GUIDE.md`](RUN_GUIDE.md).)*
2. **Edit the config** at the top of `NdmcUptimeReport.js`:
   ```js
   const REPORT_YEAR  = 2026;          // line 9
   const REPORT_MONTH = 5;             // line 10  (1=Jan … 12=Dec)
   const CUSTOM_START_DATE = "";       // line 14  ("" = full month) or "2026-05-01"
   const CUSTOM_END_DATE   = "";       // line 15  ("" = full month) or "2026-05-31"
   const JSESSIONID = "PASTE_YOUR_COOKIE_HERE";  // line 17  ← paste your cookie
   ```
3. **Close** any open copy of the target Excel file (avoids a file lock).
4. **Run it:**
   ```powershell
   cd D:\CityReport
   node NdmcUptimeReport.js
   ```
5. **Watch the console.** Each zone prints its chunks, then a final summary:
   ```
   --- Summary ---
     SP              133 switches  OK
     ...
   Written: NDMC_UptimeReport_May2026.xlsx
   ```
6. **Open the file:**
   ```powershell
   start D:\CityReport\NDMC_UptimeReport_May2026.xlsx
   ```

> If the target file is open in Excel when the script finishes, it won't crash — it
> saves to `<name>_NEW.xlsx` and tells you to close Excel and rename.

### Quick checks
- **Before** a long run, test the cookie: `node probe.js` (one fast call; prints whether
  the cookie works).
- **After** a run, sanity-check the workbook: `node verify.js` (prints per-zone row
  counts and rule checks).

---

## 8. Performance

| Date range | Approx time |
|---|---|
| 4 days (test) | 2–3 min |
| 15 days | 5–8 min |
| Full month | 10–15 min |

`MAX_CONCURRENCY` (default `3`) is the speed knob. If a full run shows **0 failed
chunks**, you can try `5`. If you see `FAIL: 502`, lower it back to `3`.

---

## 9. Project files

| File | Purpose |
|---|---|
| `NdmcUptimeReport.js` | **Main generator** — fetch, aggregate, compute, style, write |
| `probe.js` | Quick cookie/domain check before a long run |
| `verify.js` | Read-back sanity checks on the generated workbook |
| `RUN_GUIDE.md` | Click-by-click monthly run guide (incl. cookie screenshots) |
| `PROJECT_OVERVIEW.md` | This file — tech, flow, libraries, full reference |
| `WORK_LOG_*.md` | Dated work/decision notes |
| `package.json` / `package-lock.json` | Dependencies |
| `.gitignore` | Excludes `node_modules/`, `*.xlsx`, lock/temp files, `.claude/` |

---

## 10. Security

⚠️ **Never commit a real `JSESSIONID`.** It is a live session credential. All committed
files ship with the `PASTE_YOUR_COOKIE_HERE` placeholder; paste your cookie **locally
only** and keep that one line out of any commit. Generated `.xlsx` files are git-ignored.
