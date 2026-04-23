const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
['POC', 'Sheet9'].forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    if (sheet) {
        const json = XLSX.utils.sheet_to_json(sheet);
        if (json.length > 0) {
            console.log(`\n--- ${sheetName} Headers ---`);
            console.log(Object.keys(json[0]));
            console.log(`\n--- ${sheetName} Sample Row ---`);
            console.log(json[0]);
        } else {
            console.log(`${sheetName} sheet is empty`);
        }
    } else {
        console.log(`${sheetName} sheet not found`);
    }
});
