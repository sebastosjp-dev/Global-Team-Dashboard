const XLSX = require('xlsx');
const wb = XLSX.readFile('2026 Global Rev.01.xlsx');
console.log('=== SHEETS ===');
console.log(wb.SheetNames);

['ORDER SHEET','PIPELINE','POC'].forEach(name => {
  const sh = wb.Sheets[name];
  if (!sh) return;
  const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: null, raw: false });
  console.log(`\n=== ${name} | rows: ${rows.length} ===`);
  // Print first 3 header-like rows
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    console.log(`R${i}:`, JSON.stringify(rows[i]).slice(0, 800));
  }
});
