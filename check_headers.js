const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
const sheet = workbook.Sheets['ORDER SHEET'];
const json = XLSX.utils.sheet_to_json(sheet);
if (json.length > 0) {
    console.log('Headers:', Object.keys(json[0]));
    console.log('Sample Row:', json[0]);
} else {
    console.log('No data found in ORDER SHEET');
}
