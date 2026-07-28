const xlsx = require('xlsx');
const wb = xlsx.readFile('E:\\crm extractor\\Periods\\recent\\July 2026 - Focus 1.xlsx');
if (wb.Sheets['RAW DATA']) {
  const data = xlsx.utils.sheet_to_json(wb.Sheets['RAW DATA'], {defval: null});
  console.log('Total RAW DATA Rows:', data.length);
  const userRows = data.filter(r => String(r.user || r.User || '').toLowerCase().includes('abdelrahman arafat'));
  console.log('Abdelrahman Arafat Rows:', userRows.length);
  if (userRows.length > 0) {
    console.log('Sample row keys:', Object.keys(userRows[0]));
    const dates = [...new Set(userRows.map(r => r.date))];
    console.log('Unique dates for Abdelrahman Arafat:', dates.slice(0, 10));
    const targetDateRows = userRows.filter(r => String(r.date).includes('2026-07-08') || String(r.date).includes('2026-08-07'));
    console.log('Rows on 2026-07-08 / 2026-08-07:', targetDateRows.length);
    targetDateRows.forEach(r => console.log(r.date, r.shift, r.acc_name, r.doctor_name));
  }
}
