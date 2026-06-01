const axios = require("axios");
const XLSX = require("xlsx");
const https = require("https");

const API_URL = "https://smartlight.citilight.co:446/VELOCITi_API/api/ccmsOperationalreportallData";

const httpsAgent = new https.Agent({
    keepAlive: false
});

const headers = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/json",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": "https://smartlight.citilight.co:446",
    "Referer": "https://smartlight.citilight.co:446/smartlight/operationalReport",
    "Cookie": "JSESSIONID=PASTE_YOUR_COOKIE_HERE"
};

function secondsToHHMMSS(seconds) {
    if (!seconds) return "00:00:00";

    const h = Math.floor(seconds / 3600);
    const m = Math.floor((seconds % 3600) / 60);
    const s = seconds % 60;

    return [h, m, s].map(v => String(v).padStart(2, "0")).join(":");
}

async function fetchReport(cityId, startDate, endDate) {

    const body = {
        cityId: cityId,
        deviceType: "1",
        startDate: startDate,
        endDate: endDate,
        deviceArray: [],
        zoneName: "all",
        wardName: "all",
        view: "2"
    };

    try {

        const response = await axios.post(API_URL, body, {
            headers,
            httpsAgent,
            timeout: 60000
        });

        return response.data;

    } catch (err) {

        console.log(`Retrying city ${cityId} due to network issue...`);

        await new Promise(r => setTimeout(r, 3000));

        const response = await axios.post(API_URL, body, {
            headers,
            httpsAgent,
            timeout: 60000
        });

        return response.data;
    }
}

const date = {
    startDate: "2026-02-01",
    endDate: "2026-02-02"
}

// const cityData = [
//     { "mongoID": "647eb6f7364f09234fd9404d", "cityName": "SP", "data": { "Zone": [{ "Name": "" }] }, "cityId": "2" },
//     { "mongoID": "647eb6ff364f09234fd9404f", "cityName": "CITY", "data": { "Zone": [{ "Name": "" }] }, "cityId": "3" },
//     { "mongoID": "647eb7e7364f09234fd94051", "cityName": "CIVIL LINES", "data": { "Zone": [{ "Name": "" }] }, "cityId": "4" },
//     { "mongoID": "647eb807364f09234fd94053", "cityName": "KAROL BAGH", "data": { "Zone": [{ "Name": "" }] }, "cityId": "5" },
//     { "mongoID": "647eb83c364f09234fd94055", "cityName": "NARELA", "data": { "Zone": [{ "Name": "" }] }, "cityId": "6" },
//     { "mongoID": "647eb847364f09234fd94057", "cityName": "ROHINI", "data": { "Zone": [{ "Name": "" }] }, "cityId": "7" }
// ];



const cityData = [
    // { "mongoID": "647d8121a31faa711150411b", "cityName": "TEST", "data": { "Zone": [{ "Name": "" }] }, "cityId": "1" },
    { "mongoID": "647eb6f7364f09234fd9404d", "cityName": "SP", "data": { "Zone": [{ "Name": "" }] }, "cityId": "2" },
    { "mongoID": "647eb6ff364f09234fd9404f", "cityName": "CITY", "data": { "Zone": [{ "Name": "" }] }, "cityId": "3" },
    { "mongoID": "647eb7e7364f09234fd94051", "cityName": "CIVIL LINES", "data": { "Zone": [{ "Name": "" }] }, "cityId": "4" },
    { "mongoID": "647eb807364f09234fd94053", "cityName": "KAROL BAGH", "data": { "Zone": [{ "Name": "" }] }, "cityId": "5" },
    { "mongoID": "647eb83c364f09234fd94055", "cityName": "NARELA", "data": { "Zone": [{ "Name": "" }] }, "cityId": "6" },
    { "mongoID": "647eb847364f09234fd94057", "cityName": "ROHINI", "data": { "Zone": [{ "Name": "" }] }, "cityId": "7" },
    // { "mongoID": "647eb857364f09234fd94059", "cityName": "MANGOLPURI", "data": { "Zone": [{ "Name": "" }] }, "cityId": "8" },
    // { "mongoID": "648950131a29e4145f2c5409", "cityName": "ABC", "data": { "Zone": [{ "Name": "" }] }, "cityId": "9" }
]




async function generateExcel() {

    const workbook = XLSX.utils.book_new();

    for (const city of cityData) {

        console.log(`Fetching City ${city.cityName} (${city.cityId}) startDate ${date.startDate} endDate ${date.endDate}`);

        const data = await fetchReport(city.cityId, date.startDate, date.endDate);

        if (!Array.isArray(data)) {
            console.log(`No data received for ${city.cityName}`);
            continue;
        }

        const formatted = data.map(row => ({
            "Switch Point Name": row.device_name,
            "Date": row.updated_on,
            "Location": row.switch_location,
            "On Hours": secondsToHHMMSS(row.actual_on_seconds),
            "OFF Hours": secondsToHHMMSS(row.actual_off_seconds),
            "Output OFF Hours": secondsToHHMMSS(row.output_off),
            "Expected ON Hour": secondsToHHMMSS(row.expected_on),
            "Uptime": row.uptime + "%"
        }));

        const sheet = XLSX.utils.json_to_sheet(formatted);

        const sheetName = `${city.cityName}`;

        XLSX.utils.book_append_sheet(workbook, sheet, sheetName);

        // Small delay so server doesn't reset connection
        await new Promise(r => setTimeout(r, 2000));
    }

    XLSX.writeFile(workbook, "ccms_report.xlsx");

    console.log("Excel generated: ccms_report.xlsx");
}

generateExcel();