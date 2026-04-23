const XLSX = require('xlsx');
const workbook = XLSX.readFile('2026 Global Rev.01.xlsx');
const pocData = XLSX.utils.sheet_to_json(workbook.Sheets['POC']);

const analysisYear = 2026;
const now = new Date('2026-03-24');

let wonCount = 0;
let oldWonCount = 0;

pocData.forEach(r => {
    // Old Logic
    const statusKey = Object.keys(r).find(k => k.toLowerCase().includes('status'));
    const curStatus = String(r[statusKey] || '').trim().toLowerCase();
    if (curStatus.includes('won') || curStatus.includes('complete') || curStatus.includes('success')) {
        oldWonCount++;
    }

    // New Logic
    const pnKey = Object.keys(r).find(k => k.toLowerCase().replace(/\s/g, '') === 'pocnotes');
    const peKey = Object.keys(r).find(k => k.toLowerCase().replace(/\s/g, '') === 'pocend');
    const notesStr = String(r[pnKey] || '').trim().toLowerCase();
    
    let pEndDate = null;
    const val = r[peKey];
    if (val) {
        if (val instanceof Date) pEndDate = val;
        else if (typeof val === 'number') pEndDate = new Date(Math.round((val - 25569) * 86400 * 1000));
        else pEndDate = new Date(val);
    }
    
    const isActuallyWon = notesStr.includes('won') && 
                        pEndDate && !isNaN(pEndDate.getTime()) && 
                        pEndDate.getFullYear() === analysisYear && 
                        pEndDate <= now;

    if (isActuallyWon) {
        wonCount++;
    }
});

console.log(`Old Won Count (Current Status 기반): ${oldWonCount}`);
console.log(`New Won Count (POC Notes + Date 기반): ${wonCount}`);
console.log(`Difference: ${wonCount - oldWonCount}`);
