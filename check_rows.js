const xlsx = require('xlsx');
const wb = xlsx.readFile('./2026 Global Rev.01.xlsx');  // no cellDates -> raw serials
const sh = wb.Sheets['COLLECTION'];

// Walk rows; for each target deal print raw Contract Start cell + parsed value
const targets = [
  { name:'FIF',       hint:'K8S LOG APM 24.1.24',  expectStart:'2023-12-19', expectQ:2334 },
  { name:'Smartfren', hint:'APM SERVER 24.05.30',  expectStart:'2024-04-27', expectQ:34349 },
  { name:'EDTS',      hint:'24.12.07-25.12.06',    expectStart:'2024-12-23', expectQ:7459 },
  { name:'Eksad',     hint:'25.01.03~26.04.02',    expectStart:'2024-12-31', expectQ:2310 },
];

const range = xlsx.utils.decode_range(sh['!ref']);
function cellStr(r,c){ const a = xlsx.utils.encode_cell({r,c}); return sh[a]; }
// header at r=0; data starts r=1
for (const t of targets) {
  for (let r = 1; r <= range.e.r; r++) {
    const deal = cellStr(r, 8); // I
    if (!deal || !String(deal.v||'').toLowerCase().includes(t.hint.toLowerCase())) continue;
    const eu = cellStr(r,4);
    const terms = cellStr(r,7);
    const start = cellStr(r,10); // K
    const tcv = cellStr(r,12);   // M
    const y25 = cellStr(r,16);   // Q
    console.log(`\n[${t.name}] row=${r+1} EndUser=${eu&&eu.v} Terms=${terms&&terms.v}`);
    console.log(`  Contract Start cell K${r+1}: t=${start&&start.t} v=${start&&start.v} w='${start&&start.w}'`);
    if (start && start.t === 'n') {
      const d = xlsx.SSF.parse_date_code(start.v);
      console.log(`    parsed serial -> ${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`);
    }
    console.log(`  KOR TCV  M${r+1}: ${tcv&&tcv.v}  (expected total ~${t.expectQ === 7459 ? '11,700 quarterly' : t.expectQ})`);
    console.log(`  2025 RCV Q${r+1}: ${y25&&y25.v}  (user says ${t.expectQ})`);
  }
}
