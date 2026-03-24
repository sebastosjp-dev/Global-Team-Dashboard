const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
const sheet = workbook.Sheets['POC'];
if (sheet) {
    const json = XLSX.utils.sheet_to_json(sheet);
    if (json.length > 0) {
        console.log('POC Headers:', Object.keys(json[0]));
        // Filter keys that look like 'Working Days'
        const keys = Object.keys(json[0]);
        console.log('Working Days matches:', keys.filter(k => k.toLowerCase().includes('working')));
    }
}
