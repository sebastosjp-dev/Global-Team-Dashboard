const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
    if (json.length > 0) {
        console.log(`Sheet: ${sheetName}`);
        console.log(`Headers (Row 1):`, json[0]);
        if (json.length > 1) {
            console.log(`Headers (Row 2):`, json[1]);
        }
        console.log('---');
    }
});
