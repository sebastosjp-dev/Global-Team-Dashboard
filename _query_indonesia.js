const XLSX = require('xlsx');
const wb = XLSX.readFile('2026 Global Rev.01.xlsx');

// Helpers
function parseNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(/[, $]/g, ''));
  return isFinite(n) ? n : 0;
}
function parseContractYears(dealName) {
  if (!dealName) return 1;
  const s = String(dealName);
  let m = s.match(/(\d+)\s*(?:yr|years?)\b/i);
  if (m) { const y = parseInt(m[1],10); if (y>=1 && y<=10) return y; }
  m = s.match(/\b(\d{2})\.\d{1,2}(?:\.\d{1,2})?\s*[-~]\s*(\d{2})\.\d{1,2}/);
  if (m) { const diff = parseInt(m[2],10)-parseInt(m[1],10); if (diff>=1 && diff<=10) return diff; }
  m = s.match(/\b(20\d{2})\s*[-~]\s*(20\d{2})\b/);
  if (m) { const diff = parseInt(m[2],10)-parseInt(m[1],10); if (diff>=1 && diff<=10) return diff; }
  return 1;
}

const pipelineRaw = XLSX.utils.sheet_to_json(wb.Sheets['PIPELINE'], { defval: '' });
const orderRaw = XLSX.utils.sheet_to_json(wb.Sheets['ORDER SHEET'], { defval: '' });
const pocRaw = XLSX.utils.sheet_to_json(wb.Sheets['POC'], { defval: '' });

console.log('\n========== PIPELINE — INDONESIA — Q2/Q3/Q4 ==========');
const idnPipe = pipelineRaw.filter(r => String(r['Country']||'').toUpperCase().includes('IDN'));
console.log(`Total IDN pipeline rows: ${idnPipe.length}`);

const byQ = { Q2: [], Q3: [], Q4: [] };
idnPipe.forEach(r => {
  const q = String(r['Quarter']||'').toUpperCase();
  let key = '';
  if (q.includes('Q2')) key='Q2';
  else if (q.includes('Q3')) key='Q3';
  else if (q.includes('Q4')) key='Q4';
  if (key) byQ[key].push(r);
});

for (const q of ['Q2','Q3','Q4']) {
  console.log(`\n--- ${q} (${byQ[q].length} deals) ---`);
  let sumTCV=0, sumW=0, sumARR=0;
  byQ[q].forEach(r => {
    const name = r['Deal Name'];
    const stage = r['Deal Stage'];
    const tcv = parseNum(r['KOR TCV (USD)']);
    const w = parseNum(r['Weighted KOR TCV (USD)']);
    const years = parseContractYears(name);
    const arr = years > 0 ? tcv/years : tcv;
    const isRenewal = /renewal/i.test(name);
    sumTCV += tcv; sumW += w; sumARR += arr;
    console.log(`  [${isRenewal?'RENEW':'NEW  '}] ${stage.padEnd(14)} | TCV ${tcv.toLocaleString().padStart(9)} | W ${w.toLocaleString().padStart(9)} | yrs=${years} ARR ${arr.toFixed(0).padStart(9)} | ${name} | ${r['Partner']||''} | close=${r['Close Date']||''}`);
  });
  console.log(`  TOTAL: TCV=${sumTCV.toLocaleString()} | Weighted TCV=${sumW.toLocaleString()} | ARR≈${sumARR.toFixed(0)}`);
}

console.log('\n========== ORDER SHEET — INDONESIA ==========');
const idnOrder = orderRaw.filter(r => String(r['Country']||'').toUpperCase().includes('IDN'));
console.log(`Total IDN order rows: ${idnOrder.length}`);
idnOrder.forEach(r => {
  console.log(`  No.${r['No.']} | ${r['End User']||''} | ${r['CRM Deal Name']||''} | ${r['Partner']||''} | start=${r['Contract Start']} end=${r['Contract End']} | yrs=${r['Contract Yr']} | KOR_TCV=${r['KOR TCV']} ARR=${r['KOR ARR']}`);
});

console.log('\n========== POC — INDONESIA ==========');
const idnPoc = pocRaw.filter(r => String(r['Country']||'').toUpperCase().includes('IDN'));
console.log(`Total IDN POC rows: ${idnPoc.length}`);
idnPoc.forEach(r => {
  console.log(`  No.${r['No.']} | [${r['Current Status']||''}] | ${r['CRM POC Name']||''} | Partner=${r['Partner']||''} | Est=${r['Estimated Value KOR (USD)']||0} W=${r['Weighted Value KOR (USD)']||0} | ${r['POC License Start']||''}~${r['POC License End']||''} | report=${r['POC Report']||''} | tech=${r['Technical Comment']||''} | sales=${r['Sales Comment']||''}`);
});
