// Simulate how the dashboard parses & displays Contract Start in browser
const xlsx = require('xlsx');
const wb = xlsx.readFile('./2026 Global Rev.01.xlsx', { cellDates: true });
const col = xlsx.utils.sheet_to_json(wb.Sheets['COLLECTION'], { defval: '' });

const targets = ['K8S LOG APM 24.1.24', 'APM SERVER 24.05.30', '24.12.07-25.12.06', '25.01.03~26.04.02'];

console.log('Process TZ env =', process.env.TZ || '(not set; using system)');
console.log('new Date().toString() ->', new Date().toString());
console.log('---');
for (const t of targets) {
  const row = col.find(r => String(r['CRM Deal Name'] || '').includes(t));
  if (!row) { console.log('NOT FOUND', t); continue; }
  const cs = row['Contract Start'];
  console.log(`Deal contains "${t}":`);
  console.log(`  contractStart raw    : ${cs}  (type=${typeof cs}, instanceof Date=${cs instanceof Date})`);
  if (cs instanceof Date) {
    console.log(`  .toISOString()       : ${cs.toISOString()}`);
    console.log(`  .toISOString().slice : ${cs.toISOString().slice(0,10)}`);
    console.log(`  local YYYY-MM-DD     : ${cs.getFullYear()}-${String(cs.getMonth()+1).padStart(2,'0')}-${String(cs.getDate()).padStart(2,'0')}`);
  }
  console.log(`  2025 RECEIVED        : ${row['2025 RECEIVED']}`);
}
