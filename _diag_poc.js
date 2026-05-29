const XLSX = require('./node_modules/xlsx');
const wb = XLSX.readFile('2026 Global Rev.01.xlsx', { cellDates: true });
console.log('Sheets:', wb.SheetNames.join(' | '));
function dump(name) {
  const ws = wb.Sheets[name];
  if (!ws) { console.log('NO SHEET', name); return; }
  const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
  console.log('\n==== ' + name + ' headers:', Object.keys(json[0]||{}).join(' || '));
  const findName = r => { const k = Object.keys(r).find(k=>/poc name/i.test(k)||/crm/i.test(k)); return k?r[k]:''; };
  json.forEach(r => {
    const nm = String(findName(r));
    if (/BNI/i.test(nm)) {
      console.log('--- ROW BNI in', name, '---');
      Object.keys(r).forEach(k=>{ if(/note/i.test(k)) console.log('   ['+k+'] =', JSON.stringify(String(r[k]).slice(0,200))); });
    }
  });
}
dump('POC');
dump('Sheet9');
