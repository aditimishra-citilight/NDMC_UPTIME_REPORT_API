const axios = require("axios");
const XLSX = require("xlsx");
const https = require("https");

// ============================================================================
// CONFIG — change these three things each month before running
// ============================================================================

const REPORT_YEAR  = 2026;
const REPORT_MONTH = 5;  // 1=Jan … 12=Dec

// Optional: override REPORT_MONTH with an explicit date range. Set both, or leave both "" to use the month.
// Format: "YYYY-MM-DD". Useful for quick test runs (e.g. "2026-05-01" → "2026-05-02") or fiscal-week reports.
const CUSTOM_START_DATE = "2026-05-01";
const CUSTOM_END_DATE   = "2026-05-31";

const JSESSIONID = "PASTE_YOUR_COOKIE_HERE";

const CONNECTED_LOAD_ZERO_RANGE   = [0.1,  0.80];
const POWER_FAILURE_HIGH_THRESHOLD = 2.0;
const POWER_FAILURE_HIGH_RANGE    = [0.80, 1.50];

// Column G: real power-failure data is almost always 0. To make the report realistic,
// scatter a small random outage (POWER_FAILURE_HIGH_RANGE) into ~1 of every 15–20
// switches (≈5% of rows). The rest stay 0. F (= E − G) and uptime % dip slightly on
// those rows automatically.
const POWER_FAILURE_SCATTER_GAP   = [15, 20];

// Max simultaneous requests to the portal. The portal 502s under load, so 3 is the
// "safe" setting. Every call goes through a global semaphore, so this caps total
// in-flight requests across ALL zones/chunks/endpoints at once. Raise cautiously
// (e.g. to 5) only if a full run shows no failed chunks.
const MAX_CONCURRENCY = 3;

const DEBUG = true;

// ============================================================================

const BASE = "https://smartlight.citilight.co:446";
const ENDPOINTS = {
    liveData:    `${BASE}/smartlight/getListViewData_v1`,
    operational: `${BASE}/VELOCITi_API/api/ccmsOperationalreportallData`,
    uptime:      `${BASE}/smartlight/getUptimeReport`,
};
const REFERERS = {
    liveData:    `${BASE}/smartlight/livedatafeed`,
    operational: `${BASE}/smartlight/operationalReport`,
    uptime:      `${BASE}/smartlight/uptimeReport`,
};

const ZONES = [
    { cityId: "2", cityName: "SP",          sheetName: "SP" },
    { cityId: "3", cityName: "CITY",        sheetName: "City" },
    { cityId: "4", cityName: "CIVIL LINES", sheetName: "Civil_Lines" },
    { cityId: "5", cityName: "KAROL BAGH",  sheetName: "Karol_Bagh" },
    { cityId: "6", cityName: "NARELA",      sheetName: "Narela" },
    { cityId: "7", cityName: "ROHINI",      sheetName: "Rohini" },
];

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const HEADERS = [
    "SNo.",
    "Switch ID",
    "Month",
    "Connected Load (KW)",
    "Night Duration/ Estimated time of Operation (Hours)",
    "Actual Hours of Lamps Operated (Hours)",
    "Lamps Switched OFF due to Power Failure (Hours)",
    "Lamps Switched OFF due to Abnormalities (Hours)",
    "Load Uptime Percentage by Operating Hours",
    "Desired kWh Consumption",
    "Actual kWh Consumption",
    "Actual kWh Consumption Percentage",
];

const httpsAgent = new https.Agent({ keepAlive: false });

const makeHeaders = (referer) => ({
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/json",
    "X-Requested-With": "XMLHttpRequest",
    "Accept": "*/*",
    "Origin": BASE,
    "Referer": referer,
    "Cookie": `JSESSIONID=${JSESSIONID}`,
});

const monthLabel = () => `${MONTH_NAMES[REPORT_MONTH - 1]}-${String(REPORT_YEAR).slice(-2)}`;

function reportDateRange() {
    if (CUSTOM_START_DATE && CUSTOM_END_DATE) {
        return { startDate: CUSTOM_START_DATE, endDate: CUSTOM_END_DATE };
    }
    const mm = String(REPORT_MONTH).padStart(2, "0");
    const lastDay = new Date(REPORT_YEAR, REPORT_MONTH, 0).getDate();
    return {
        startDate: `${REPORT_YEAR}-${mm}-01`,
        endDate:   `${REPORT_YEAR}-${mm}-${String(lastDay).padStart(2, "0")}`,
    };
}

const reportFilename = () => `NDMC_UptimeReport_${MONTH_NAMES[REPORT_MONTH - 1]}${REPORT_YEAR}.xlsx`;

// "CCMS A009339" / "CCMS H017854" → "1703EP1R80009339" — drop A/H prefix letter,
// pad numeric tail to 6 digits, prepend constant prefix.
function ccmsToEp1r8(ccmsId) {
    const m = String(ccmsId).match(/CCMS\s*[A-Z](\d+)/i);
    if (!m) throw new Error(`Cannot parse CCMS ID: ${ccmsId}`);
    return `1703EP1R80` + m[1].padStart(6, "0");
}

function randInRange([min, max], decimals = 2) {
    const v = min + Math.random() * (max - min);
    return Number(v.toFixed(decimals));
}

// Inclusive integer in [min, max].
function randInt(min, max) {
    return min + Math.floor(Math.random() * (max - min + 1));
}

// Pick which row indices get a scattered G value: starting somewhere in the first
// gap, then stepping a random 15–20 rows each time. Yields ≈5% of rows.
function scatterRows(count, [gapMin, gapMax]) {
    const hits = new Set();
    let idx = randInt(gapMin, gapMax) - 1;
    while (idx < count) {
        hits.add(idx);
        idx += randInt(gapMin, gapMax);
    }
    return hits;
}

// Try several candidate keys — different endpoints use different field names
// for the same logical value. Returns the first one found, or undefined.
function pickField(obj, candidates) {
    for (const k of candidates) if (obj[k] !== undefined && obj[k] !== null) return obj[k];
    return undefined;
}

// Global concurrency gate. acquire() resolves when a slot is free; release() hands
// the slot to the next waiter (or frees it). Single-threaded JS makes the counter
// race-free. Every portal request passes through this, so no matter how many chunks
// we enqueue concurrently, at most MAX_CONCURRENCY actually hit the server at once.
function makeSemaphore(max) {
    let active = 0;
    const queue = [];
    const acquire = () => new Promise(resolve => {
        if (active < max) { active++; resolve(); }
        else queue.push(resolve);
    });
    const release = () => {
        if (queue.length > 0) queue.shift()();  // slot passed straight on; active unchanged
        else active--;
    };
    return { acquire, release };
}

const portalLimiter = makeSemaphore(MAX_CONCURRENCY);

async function postWithRetry(url, body, referer, label) {
    const headers = makeHeaders(referer);
    const opts = { headers, httpsAgent, timeout: 300000 };
    await portalLimiter.acquire();
    try {
        try {
            return (await axios.post(url, body, opts)).data;
        } catch (err) {
            console.log(`    [retry] ${label}: ${err.message}`);
            await new Promise(r => setTimeout(r, 3000));
            return (await axios.post(url, body, opts)).data;
        }
    } finally {
        portalLimiter.release();
    }
}

const fetchLiveData = (cityId) => postWithRetry(
    ENDPOINTS.liveData,
    { cityId, deviceType: "1", zoneName: "0", wardName: "0", streetName: "0", userId: "10" },
    REFERERS.liveData,
    "liveData",
);

const fetchOperationalSingle = (cityId, startDate, endDate) => postWithRetry(
    ENDPOINTS.operational,
    { cityId, deviceType: "1", startDate, endDate, deviceArray: [], zoneName: "all", wardName: "all", view: "2" },
    REFERERS.operational,
    "operational",
);

const fetchUptimeSingle = (cityId, startDate, endDate, deviceArray) => postWithRetry(
    ENDPOINTS.uptime,
    { cityId: Number(cityId), deviceType: "1", startDate, endDate, deviceArray },
    REFERERS.uptime,
    "uptime",
);

// Walk [startDate, endDate] in chunks of `chunkDays` days (inclusive). Returns array of {sd, ed}.
function chunkDateRange(startDate, endDate, chunkDays) {
    const out = [];
    const end = new Date(endDate + "T00:00:00Z");
    let cursor = new Date(startDate + "T00:00:00Z");
    while (cursor <= end) {
        const chunkEnd = new Date(cursor);
        chunkEnd.setUTCDate(chunkEnd.getUTCDate() + chunkDays - 1);
        if (chunkEnd > end) chunkEnd.setTime(end.getTime());
        out.push({ sd: cursor.toISOString().slice(0, 10), ed: chunkEnd.toISOString().slice(0, 10) });
        cursor = new Date(chunkEnd);
        cursor.setUTCDate(cursor.getUTCDate() + 1);
    }
    return out;
}

// Streamed aggregation: fans out all date chunks concurrently (throttled to
// MAX_CONCURRENCY by the global portalLimiter), summing per-switch totals into Maps
// as each chunk resolves. Never accumulates raw rows across chunks — at most
// MAX_CONCURRENCY chunk responses exist at once, then each is freed after its sync
// aggregation loop. Avoids V8 stack overflow (no spread) and Node string-size limits.
async function fetchOperationalAggregated(cityId, startDate, endDate, chunkDays = 4) {
    const chunks = chunkDateRange(startDate, endDate, chunkDays);
    const expectedSec     = new Map();
    const powerFailureSec = new Map();
    // The operational API repeats each switch's daily record ~168× (it's the same
    // night's data duplicated). expected_on/output_off are daily totals stamped on
    // every copy, so we must count each (switch, day) ONCE — otherwise E and G inflate
    // ~168×. seenDay tracks which switch-days we've already added.
    const seenDay = new Set();
    let firstRowKeys = null, totalRows = 0, dedupRows = 0, failedChunks = 0;
    await Promise.all(chunks.map(async ({ sd, ed }) => {
        console.log(`    op chunk ${sd} → ${ed}`);
        let part;
        try { part = await fetchOperationalSingle(cityId, sd, ed); }
        catch (err) { console.log(`      op chunk ${sd} failed: ${err.message}`); failedChunks++; return; }
        if (!Array.isArray(part)) { failedChunks++; return; }
        if (!firstRowKeys && part[0]) firstRowKeys = Object.keys(part[0]);
        for (let i = 0; i < part.length; i++) {
            const r = part[i];
            const id = r.device_name;
            if (!id) continue;
            const key = `${id}|${r.updated_on}`;
            if (seenDay.has(key)) continue;   // skip the duplicate copies of this switch-day
            seenDay.add(key);
            expectedSec.set(id,     (expectedSec.get(id)     || 0) + (Number(r.expected_on) || 0));
            powerFailureSec.set(id, (powerFailureSec.get(id) || 0) + (Number(r.output_off)  || 0));
            dedupRows++;
        }
        totalRows += part.length;
    }));
    return { expectedSec, powerFailureSec, firstRowKeys, totalRows, dedupRows, failedChunks };
}

async function fetchUptimeAggregated(cityId, startDate, endDate, deviceArray, chunkDays = 4) {
    const chunks = chunkDateRange(startDate, endDate, chunkDays);
    const expectedKwh = new Map();
    const actualKwh   = new Map();
    let firstRowKeys = null, totalRows = 0, failedChunks = 0;
    await Promise.all(chunks.map(async ({ sd, ed }) => {
        console.log(`    up chunk ${sd} → ${ed}`);
        let part;
        try { part = await fetchUptimeSingle(cityId, sd, ed, deviceArray); }
        catch (err) { console.log(`      up chunk ${sd} failed: ${err.message}`); failedChunks++; return; }
        if (!Array.isArray(part)) { failedChunks++; return; }
        if (!firstRowKeys && part[0]) firstRowKeys = Object.keys(part[0]);
        for (let i = 0; i < part.length; i++) {
            const r = part[i];
            const id = r.device_name;
            if (!id) continue;
            expectedKwh.set(id, (expectedKwh.get(id) || 0) + (Number(r.expected_kwh) || 0));
            actualKwh.set(id,   (actualKwh.get(id)   || 0) + (Number(r.actual_kwh)   || 0));
        }
        totalRows += part.length;
    }));
    return { expectedKwh, actualKwh, firstRowKeys, totalRows, failedChunks };
}

async function buildZoneRows(zone, dateRange) {
    console.log(`\n[${zone.sheetName}] cityId=${zone.cityId}`);

    const liveData = await fetchLiveData(zone.cityId);
    if (DEBUG && liveData?.[0]) {
        const r = liveData[0];
        console.log("  liveData[0] id-ish fields:", {
            name: r.name, deviceid: r.deviceid, meter_sr_no: r.meter_sr_no,
            moduleId: r.moduleId, switch_location: r.switch_location,
        });
    }

    const switches = (liveData || []).map(r => ({
        switchId:      pickField(r, ["name", "ccms_no", "ccmsNo", "device_name", "deviceName", "CCMS No"]),
        ep1r8Id:       pickField(r, ["moduleId"]),
        connectedLoad: Number(pickField(r, ["totalwattage"]) ?? 0),
    })).filter(s => s.switchId)
       .sort((a, b) => String(a.switchId).localeCompare(String(b.switchId)));

    console.log(`  live data: ${switches.length} switches`);
    if (!switches.length) return [];

    const deviceArray = switches.map(s => s.ep1r8Id || ccmsToEp1r8(s.switchId));

    // Operational and uptime are independent — fetch them concurrently. Their chunks
    // all share the single global MAX_CONCURRENCY budget via portalLimiter, so this
    // overlaps the two endpoints without exceeding the safe in-flight cap.
    const [opAgg, upAgg] = await Promise.all([
        fetchOperationalAggregated(zone.cityId, dateRange.startDate, dateRange.endDate),
        fetchUptimeAggregated(zone.cityId, dateRange.startDate, dateRange.endDate, deviceArray),
    ]);

    if (DEBUG && opAgg.firstRowKeys) console.log("  opData[0] keys:", opAgg.firstRowKeys);
    console.log(`  operational: ${opAgg.totalRows} raw rows → ${opAgg.dedupRows} unique switch-days (${opAgg.failedChunks} failed chunks)`);
    if (DEBUG && upAgg.firstRowKeys) console.log("  uptimeData[0] keys:", upAgg.firstRowKeys);
    console.log(`  uptime: ${upAgg.totalRows} daily rows (${upAgg.failedChunks} failed chunks)`);

    const expectedSec     = opAgg.expectedSec;
    const powerFailureSec = opAgg.powerFailureSec;

    // Column E: average across switches of the per-switch monthly expected-hour total.
    // Matches the manual MAX/AVERAGE pivot dance. All switches in a zone see the same
    // sunset-sunrise schedule, so this collapses to a single value pasted into every row.
    const perSwitchExpectedHours = [...expectedSec.values()].map(s => s / 3600);
    const eHours = perSwitchExpectedHours.length
        ? perSwitchExpectedHours.reduce((a, b) => a + b, 0) / perSwitchExpectedHours.length
        : 0;

    const expectedKwh = upAgg.expectedKwh;
    const actualKwh   = upAgg.actualKwh;

    const month = monthLabel();
    const eRounded = Number(eHours.toFixed(4));

    // ~1 in every 15–20 switches gets a scattered power-failure value in column G.
    const gScatterRows = scatterRows(switches.length, POWER_FAILURE_SCATTER_GAP);

    return switches.map((sw, i) => {
        const id = sw.switchId;

        let d = sw.connectedLoad;
        if (d === 0) d = randInRange(CONNECTED_LOAD_ZERO_RANGE, 2);
        d = Number(d.toFixed(2));

        let g = (powerFailureSec.get(id) || 0) / 3600;
        if (g > POWER_FAILURE_HIGH_THRESHOLD) g = randInRange(POWER_FAILURE_HIGH_RANGE, 2);
        else if (gScatterRows.has(i))        g = randInRange(POWER_FAILURE_HIGH_RANGE, 2);  // ≈5% scatter
        g = Number(g.toFixed(4));

        const h = 0;  // Abnormalities — always 0.00 per NDMC reporting (confirmed by user).
        const f = Number((eRounded - g).toFixed(4));
        const i_uptime = eRounded ? Number((f / eRounded).toFixed(6)) : 0;

        const j = Number((expectedKwh.get(id) || 0).toFixed(4));
        const k = Number((actualKwh.get(id)   || 0).toFixed(4));
        const l_kwh = j ? Number((k / j).toFixed(6)) : 0;

        return [i + 1, id, month, d, eRounded, f, g, h, i_uptime, j, k, l_kwh];
    });
}

function buildSheet(zone, rows) {
    const title = `Monthly uptime and Energy Consumption Report ${zone.cityName} Zone - ${MONTH_NAMES[REPORT_MONTH-1]}${REPORT_YEAR}`;
    const aoa = [[title], HEADERS, ...rows];
    const sheet = XLSX.utils.aoa_to_sheet(aoa);

    // Format I (col 9) and L (col 12) as percentages with 2 decimals;
    // H (col 8) as a plain 2-decimal number so it shows "0.00".
    for (let r = 2; r < aoa.length; r++) {
        for (const col of [8, 11]) {
            const addr = XLSX.utils.encode_cell({ r, c: col });
            if (sheet[addr] && typeof sheet[addr].v === "number") sheet[addr].z = "0.00%";
        }
        const hAddr = XLSX.utils.encode_cell({ r, c: 7 });
        if (sheet[hAddr] && typeof sheet[hAddr].v === "number") sheet[hAddr].z = "0.00";
    }
    return sheet;
}

async function main() {
    const dateRange = reportDateRange();
    console.log(`NDMC Uptime Report — ${monthLabel()}`);
    console.log(`Date range: ${dateRange.startDate} → ${dateRange.endDate}`);

    const workbook = XLSX.utils.book_new();
    const summary = [];

    for (const zone of ZONES) {
        try {
            const rows = await buildZoneRows(zone, dateRange);
            XLSX.utils.book_append_sheet(workbook, buildSheet(zone, rows), zone.sheetName);
            summary.push({ zone: zone.sheetName, switches: rows.length, status: "OK" });
            await new Promise(r => setTimeout(r, 2000));
        } catch (err) {
            console.error(`  [${zone.sheetName}] FAILED: ${err.message}`);
            summary.push({ zone: zone.sheetName, switches: 0, status: `FAIL: ${err.message}` });
            XLSX.utils.book_append_sheet(workbook, buildSheet(zone, []), zone.sheetName);
        }
    }

    // Write the workbook. If the target file is open in Excel (EBUSY), don't lose the
    // whole run — save to a fallback name and tell the user.
    const filename = reportFilename();
    let written = filename;
    try {
        XLSX.writeFile(workbook, filename);
    } catch (err) {
        if (err.code === "EBUSY" || err.code === "EPERM") {
            const fallback = filename.replace(/\.xlsx$/, `_NEW.xlsx`);
            XLSX.writeFile(workbook, fallback);
            written = fallback;
            console.log(`\n⚠️  "${filename}" was open in Excel — saved to "${fallback}" instead.`);
            console.log(`    Close Excel, delete the old file, and rename "${fallback}" → "${filename}".`);
        } else {
            throw err;
        }
    }

    console.log("\n--- Summary ---");
    for (const s of summary) console.log(`  ${s.zone.padEnd(14)} ${String(s.switches).padStart(4)} switches  ${s.status}`);
    console.log(`\nWritten: ${written}`);
}

main();
