// Inspect how Weekly draft sheet gets parsed via sheet_to_json (default mode used by app.js)
const XLSX = require('xlsx');
const path = require('path');

const wb = XLSX.readFile(path.join(__dirname, '2026 Global Rev.01.xlsx'));
const ws = wb.Sheets['Weekly draft'];
const json = XLSX.utils.sheet_to_json(ws, { defval: '', cellDates: true });

console.log('--- KEYS of first row ---');
console.log(Object.keys(json[0] || {}));
console.log('\n--- First 18 rows (with all keys) ---');
for (let i = 0; i < Math.min(18, json.length); i++) {
    const r = json[i];
    const compact = {};
    Object.keys(r).forEach(k => {
        const v = String(r[k]);
        if (v !== '') compact[k] = v.substring(0, 50);
    });
    if (Object.keys(compact).length > 0) console.log(`R${i+2}:`, JSON.stringify(compact));
}
