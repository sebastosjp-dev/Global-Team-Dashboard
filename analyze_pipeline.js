const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
const sheet = workbook.Sheets['PIPELINE'];
const data = XLSX.utils.sheet_to_json(sheet, { cellDates: true });

const stats = {
    'Q1': { all: 0, y2026: 0 },
    'Q2': { all: 0, y2026: 0 },
    'Q3': { all: 0, y2026: 0 },
    'Q4': { all: 0, y2026: 0 }
};

data.forEach(r => {
    const tcv = parseFloat(r['KOR TCV (USD)']) || 0;
    const qValue = String(r['Quarter'] || '').toUpperCase();
    const closeDate = r['Close Date'];
    
    let q = '';
    if (qValue.includes('Q1')) q = 'Q1';
    else if (qValue.includes('Q2')) q = 'Q2';
    else if (qValue.includes('Q3')) q = 'Q3';
    else if (qValue.includes('Q4')) q = 'Q4';
    
    if (q) {
        stats[q].all += tcv;
        if (closeDate instanceof Date && closeDate.getFullYear() === 2026) {
            stats[q].y2026 += tcv;
        }
    }
});

console.log('Quarterly Stats:');
console.log(JSON.stringify(stats, null, 2));
