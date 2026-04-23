const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Mocking some internal logic to understand how data is processed
function parseCurrency(val) {
    if (!val) return 0;
    if (typeof val === 'number') return val;
    return parseFloat(String(val).replace(/[$,\s]/g, '')) || 0;
}

const workbook = xlsx.readFile('2026 Global Rev.01.xlsx');
const pocData = xlsx.utils.sheet_to_json(workbook.Sheets['POC']);

const targetName = '[SimasJiwa][On Prem_Server,DB]';
const found = pocData.find(r => {
    const keys = Object.keys(r);
    const pocNameKey = keys.find(k => k.toLowerCase().replace(/[^a-z]/g, '') === 'crmpocname');
    return r[pocNameKey] === targetName;
});

if (found) {
    console.log('Found POC:', JSON.stringify(found, null, 2));
    const keys = Object.keys(found);
    const statusKey = keys.find(k => k.toLowerCase().includes('current status')) || keys.find(k => k.toLowerCase().includes('status'));
    const curStatus = String(found[statusKey] || '').trim().toLowerCase();
    const wdKey = keys.find(k => k.toLowerCase().includes('working days')) || keys.find(k => k.toLowerCase().includes('workingdays'));
    const runningDays = Number(found[wdKey] || 0);

    const pnKey = keys.find(k => k.toLowerCase().replace(/\s/g, '') === 'pocnotes');
    const notesStr = String(found[pnKey] || '').trim().toLowerCase();

    console.log('Status:', curStatus);
    console.log('Working Days:', runningDays);
    console.log('Notes:', notesStr);
    
    const analysisYear = 2026;
    const now = new Date('2026-03-24');
    const peKey = keys.find(k => k.toLowerCase().replace(/\s/g, '') === 'pocend');
    let pEndDate = null;
    if (found[peKey]) {
        if (found[peKey] instanceof Date) pEndDate = found[peKey];
        else if (typeof found[peKey] === 'number') pEndDate = new Date(Math.round((found[peKey] - 25569) * 86400 * 1000));
        else pEndDate = new Date(found[peKey]);
    }

    const isActuallyWon = notesStr.includes('won') && pEndDate && !isNaN(pEndDate.getTime()) && pEndDate.getFullYear() === analysisYear && pEndDate <= now;
    console.log('Is Actually Won (logic from services.js line 395):', isActuallyWon);
} else {
    console.log('POC not found:', targetName);
}
