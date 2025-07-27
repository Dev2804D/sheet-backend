const XLSX = require('xlsx');
const fs = require('fs');

// Initial headers and one row
const initialData = [
  {Id: 12345, UPI: 'test@upi', Mobile: '9999999999', Timestamp: new Date().toISOString() }
];

const ws = XLSX.utils.json_to_sheet(initialData);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
XLSX.writeFile(wb, 'upstock.xlsx');

console.log('✅ Excel data is created');
