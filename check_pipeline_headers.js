const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
const sheet = workbook.Sheets['PIPELINE'];
if (sheet) {
    const json = XLSX.utils.sheet_to_json(sheet);
    if (json.length > 0) {
        console.log('PIPELINE Headers:', Object.keys(json[0]));
    }
}
