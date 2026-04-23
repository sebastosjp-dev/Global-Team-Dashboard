const XLSX = require('xlsx');
const fs = require('fs');

function searchInFile(filePath) {
    if (!fs.existsSync(filePath)) return;
    const workbook = XLSX.readFile(filePath);
    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        rows.forEach((row, idx) => {
            const rowStr = JSON.stringify(row);
            if (rowStr.includes('Andy')) {
                console.log(`Found 'Andy' in ${filePath} -> ${sheetName} at row ${idx + 2}`);
            }
        });
    });
}

// I need to find where XLSX is if it's not in node_modules.
// In the browser it's loaded from CDN.
// I'll try to find it in the directory.
searchInFile('2026 Global Rev.01.xlsx');
searchInFile('Global MRR ARR.xlsx');
