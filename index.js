const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');


const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(bodyParser.json());

const FILE_PATH = './upstock.xlsx'; 

if (!fs.existsSync(FILE_PATH)) {
  const ws = XLSX.utils.json_to_sheet([]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, FILE_PATH);
}

console.log(`📝 Excel file path: ${FILE_PATH}`);


app.post('/submit', async (req, res) => {
  const {Id, UPI, Mobile,Offer, Timestamp } = req.body;

  try {
    // 1. Read existing Excel file
    const workbook = XLSX.readFile(FILE_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 2. Convert existing sheet to JSON
    const data = XLSX.utils.sheet_to_json(worksheet);

    // 3. Append new entry
    data.push({Id, UPI, Mobile,Offer, Timestamp });

    // 4. Convert JSON back to sheet
    const updatedSheet = XLSX.utils.json_to_sheet(data);

    // 5. Update workbook and write back to file
    workbook.Sheets[sheetName] = updatedSheet;
    XLSX.writeFile(workbook, FILE_PATH);
    console.log('✅ Excel file updated');

    res.status(200).json({ status: 'success', message: 'Data appended to Excel' });
  } catch (err) {
    console.error('❌ Excel Error:', err.message);
    res.status(500).json({ status: 'error', message: err.message });
  }
});




app.get('/view-excel', (req, res) => {
  const workbook = XLSX.readFile(FILE_PATH);
  const sheetName = workbook.SheetNames[0];
  const tableHTML = XLSX.utils.sheet_to_html(workbook.Sheets[sheetName]);

  const styledHTML = `
    <html>
      <head>
        <title>Excel View</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            background: #f5f5f5;
            padding: 40px;
          }
          .search-box {
            margin-bottom: 20px;
            padding: 10px;
            font-size: 1rem;
            width: 100%;
            max-width: 400px;
            border: 1px solid #ccc;
            border-radius: 6px;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            background: white;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
          }
          th, td {
            border: 1px solid #ccc;
            padding: 12px 16px;
            text-align: left;
          }
          th {
            background-color: #007bff;
            color: white;
          }
          tr:nth-child(even) {
            background-color: #f2f2f2;
          }
          caption {
            caption-side: top;
            font-size: 1.5rem;
            margin-bottom: 15px;
            font-weight: bold;
            text-align: left;
          }
        </style>
      </head>
      <body>
        <input type="text" class="search-box" placeholder="Search by any value..." onkeyup="searchTable()" />
        ${tableHTML}
        <script>
          function searchTable() {
            const input = document.querySelector('.search-box');
            const filter = input.value.toLowerCase();
            const rows = document.querySelectorAll('table tr');

            rows.forEach((row, index) => {
              if (index === 0) return; // Skip header
              const cells = row.querySelectorAll('td');
              const match = Array.from(cells).some(cell => cell.textContent.toLowerCase().includes(filter));
              row.style.display = match ? '' : 'none';
            });
          }
        </script>
      </body>
    </html>
  `;

  res.send(styledHTML);
});


app.listen(PORT, () => {
  console.log(`✅ Server running on port ${PORT}`);
});
