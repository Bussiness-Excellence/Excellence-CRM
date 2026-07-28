const xlsx = require('xlsx');
const wb = xlsx.readFile('E:\\crm extractor\\Periods\\recent\\July 2026 - Eagles 1.xlsx');
console.log('All Sheet Names:', wb.SheetNames);
if (wb.Sheets['RAW DATA']) {
  const data = xlsx.utils.sheet_to_json(wb.Sheets['RAW DATA'], {defval: null});
  console.log('RAW DATA Rows:', data.length);
  if (data.length > 0) {
    console.log('RAW DATA Columns:', Object.keys(data[0]).join(' | '));
    console.log('First 2 rows:', JSON.stringify(data.slice(0, 2), null, 2));
  }
} else {
  console.log('RAW DATA sheet not found. Inspecting first sheet:', wb.SheetNames[0]);
  const data = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval: null});
  if (data.length > 0) console.log('Columns:', Object.keys(data[0]).join(' | '));
}
