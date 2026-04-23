const xlsx = require('xlsx');
const fs = require('fs');
const { getPipelineStats } = require('./services.js');
const { isCountryMatch } = require('./utils.js');

try {
    const workbook = xlsx.readFile('./2026 Global Rev.01.xlsx', { type: 'buffer', cellDates: true });
    
    let pData = xlsx.utils.sheet_to_json(workbook.Sheets['PIPELINE'] || workbook.Sheets[workbook.SheetNames.find(s => s.toLowerCase().includes('pipeline'))], { defval: "" });
    let oData = xlsx.utils.sheet_to_json(workbook.Sheets['ORDER SHEET'] || workbook.Sheets[workbook.SheetNames.find(s => s.toLowerCase().includes('order'))], { defval: "" });

    const filterCountry = 'Thailand';
    
    const filteredPData = pData.filter(r => isCountryMatch(r, filterCountry));
    const filteredOData = oData.filter(r => isCountryMatch(r, filterCountry));

    const stats = getPipelineStats(filteredPData, filteredOData);

    console.log("Q1 Thailand Deals:", stats.pipelineByQuarter['Q1'].countries['Thailand']);
    console.log("Overall Q1 deals:", stats.pipelineByQuarter['Q1']);
    console.log("Total global TCV computed:", stats.globalTotalTcv);
} catch (err) {
    console.error(err);
}
