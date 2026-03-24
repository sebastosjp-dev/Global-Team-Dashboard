const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
const sheet = workbook.Sheets['POC'];
if (sheet) {
    const json = XLSX.utils.sheet_to_json(sheet);
    if (json.length > 0) {
        const statusKey = Object.keys(json[0]).find(k => k.toLowerCase().includes('status'));
        if (statusKey) {
            const statuses = new Set();
            json.forEach(r => {
                if (r[statusKey]) statuses.add(String(r[statusKey]).trim());
            });
            console.log('Unique Statuses:', Array.from(statuses));
        } else {
            console.log('Status key not found. Headers:', Object.keys(json[0]));
        }
    }
}
