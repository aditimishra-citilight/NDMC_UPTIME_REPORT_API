# NDMC_UPTIME_REPORT_API

Automates the **NDMC Monthly Uptime & Energy Consumption Report** — generates
`NDMC_UptimeReport_<Month><Year>.xlsx` with one sheet per zone (SP, City, Civil_Lines,
Karol_Bagh, Narela, Rohini), pulled directly from the Citilight smartlight portal APIs.

## What it does

In a single run it fetches, for all 6 zones:
- **Live Data Feed** → Switch ID + Connected Load (columns B, D)
- **Operational Report** → Night/expected hours, Power Failure (columns E, G)
- **Uptime Report** → Desired & Actual kWh (columns J, K)

It then computes the formula columns (F = E−G, I = F/E, L = K/J), applies the
business rules, and writes the finished, fully-styled workbook (via **ExcelJS**) —
no manual pivots or VLOOKUPs.

**📘 For the complete reference** — tech stack, every library, the end-to-end data flow,
API sources, column logic, Excel formatting, and all commands — see
[`PROJECT_OVERVIEW.md`](PROJECT_OVERVIEW.md).

## Usage

1. Get a fresh `JSESSIONID` cookie from `https://smartlight.citilight.co:446`
   (see `RUN_GUIDE.md` for the exact steps).
2. In `NdmcUptimeReport.js`, set `REPORT_YEAR`, `REPORT_MONTH`, the date range, and
   paste the cookie into `JSESSIONID`.
3. Run:
   ```powershell
   npm install
   node NdmcUptimeReport.js
   ```

Full step-by-step instructions, troubleshooting, and the column source mapping are in
[`RUN_GUIDE.md`](RUN_GUIDE.md).

## Performance

Requests are fetched concurrently (throttled by a global semaphore, default 3) with
operational + uptime running in parallel per zone. A full month runs in ~10–15 min.

## Security

⚠️ **Never commit a real `JSESSIONID`.** The cookie is a live session credential; the
files here ship with a `PASTE_YOUR_COOKIE_HERE` placeholder. Paste your cookie locally
and do not push it.

## Files

| File | Purpose |
|---|---|
| `NdmcUptimeReport.js` | Main report generator |
| `probe.js` | Quick cookie/domain check before a long run |
| `verify.js` | Sanity-check the generated workbook's columns |
| `PROJECT_OVERVIEW.md` | Complete reference: tech, libraries, flow, columns, commands |
| `RUN_GUIDE.md` | Full monthly run instructions |
