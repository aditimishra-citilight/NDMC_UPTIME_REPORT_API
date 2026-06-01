// Quick cookie/domain probe — one fast call to the Live Data Feed for SP.
// Prints status, content-type, and a hint so we know the cookie works BEFORE a long run.
const axios = require("axios");
const https = require("https");

const JSESSIONID = "PASTE_YOUR_COOKIE_HERE";
const BASE = "https://smartlight.citilight.co:446";
const httpsAgent = new https.Agent({ keepAlive: false });

(async () => {
    try {
        const res = await axios.post(
            `${BASE}/smartlight/getListViewData_v1`,
            { cityId: "2", deviceType: "1", zoneName: "0", wardName: "0", streetName: "0", userId: "10" },
            {
                headers: {
                    "User-Agent": "Mozilla/5.0",
                    "Content-Type": "application/json",
                    "X-Requested-With": "XMLHttpRequest",
                    "Accept": "*/*",
                    "Origin": BASE,
                    "Referer": `${BASE}/smartlight/livedatafeed`,
                    "Cookie": `JSESSIONID=${JSESSIONID}`,
                },
                httpsAgent,
                timeout: 60000,
                maxRedirects: 0,
                validateStatus: () => true,
            },
        );
        const ct = res.headers["content-type"] || "";
        console.log("HTTP status :", res.status);
        console.log("content-type:", ct);
        console.log("location    :", res.headers["location"] || "(none)");
        if (Array.isArray(res.data)) {
            console.log("RESULT      : ✅ COOKIE GOOD — got", res.data.length, "switches for SP");
            const r0 = res.data[0] || {};
            console.log("sample id   :", r0.name, "| totalwattage:", r0.totalwattage);
        } else if (typeof res.data === "string" && res.data.toLowerCase().includes("login")) {
            console.log("RESULT      : ❌ COOKIE BAD — got a login page (expired or wrong domain)");
        } else {
            console.log("RESULT      : ⚠️ Unexpected response shape:", typeof res.data);
            console.log(String(JSON.stringify(res.data)).slice(0, 200));
        }
    } catch (err) {
        console.log("RESULT      : ❌ Request error:", err.message);
    }
})();
