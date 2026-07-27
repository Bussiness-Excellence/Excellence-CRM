const xlsx = require('xlsx');
const wb = xlsx.readFile('E:\\crm extractor\\hierarchy\\hierarchy_export.xlsx');
const sheet = wb.Sheets[wb.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(sheet, {defval: null});
console.log('Sheet:', wb.SheetNames[0]);
console.log('Rows:', data.length);
console.log('Columns:', Object.keys(data[0]).join(', '));
console.log('First 3 rows:');
data.slice(0,3).forEach((r,i) => console.log(JSON.stringify(r)));
