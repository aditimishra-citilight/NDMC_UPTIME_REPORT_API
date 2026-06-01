// Read the generated workbook and sanity-check the columns.
const XLSX = require("xlsx");
const wb = XLSX.readFile("NDMC_UptimeReport_May2026.xlsx");

for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
    const data = rows.slice(2).filter(r => r[1]); // skip title+header, keep rows with a Switch ID
    const col = (i) => data.map(r => r[i]).filter(v => typeof v === "number");
    const E = data[0] ? data[0][4] : null;
    const eAllSame = data.every(r => r[4] === E);
    const gMax = Math.max(...col(6));
    const dZeros = col(3).filter(v => v === 0).length;
    console.log(`\n[${name}] ${data.length} rows`);
    console.log(`  sample row 1: B=${data[0][1]} C=${data[0][2]} D=${data[0][3]} E=${data[0][4]} F=${data[0][5]} G=${data[0][6]} H=${data[0][7]} I=${data[0][8]} J=${data[0][9]} K=${data[0][10]} L=${data[0][11]}`);
    console.log(`  E same in all rows: ${eAllSame} (E=${E})`);
    console.log(`  G max: ${gMax.toFixed(3)} (rule: must be <= 2.0)`);
    const hAllZero = data.every(r => r[7] === 0);
    console.log(`  H all 0.00: ${hAllZero} (rule: must be true)`);
    console.log(`  E hours: ${E} (rule: ~21 for a 2-day test, ~325 for a month)`);
    console.log(`  D zeros remaining: ${dZeros} (rule: must be 0)`);
    console.log(`  J>0 rows: ${col(9).filter(v => v > 0).length}/${data.length},  K>0 rows: ${col(10).filter(v => v > 0).length}/${data.length}`);
}
