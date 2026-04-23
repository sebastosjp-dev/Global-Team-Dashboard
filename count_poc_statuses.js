const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
const sheet = workbook.Sheets['POC'];
if (sheet) {
    const json = XLSX.utils.sheet_to_json(sheet);
    if (json.length > 0) {
        // Find the "Current Status" key
        const statusKey = Object.keys(json[0]).find(k => k.trim() === 'Current Status' || k.toLowerCase().includes('status'));
        console.log('Using status key:', statusKey);
        
        if (statusKey) {
            const counts = {};
            json.forEach(r => {
                let status = r[statusKey];
                if (status) {
                    status = String(status).trim();
                    counts[status] = (counts[status] || 0) + 1;
                }
            });
            console.log('Status Counts:', JSON.stringify(counts, null, 2));
        } else {
            console.log('Status key not found. Headers:', Object.keys(json[0]));
        }
    } else {
        console.log('No data found in sheet.');
    }
} else {
    console.log('Sheet "POC" not found.');
}
