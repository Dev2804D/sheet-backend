const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors());
app.use(bodyParser.json());

const FILE_PATH = './data.xlsx'; // Your local Excel file

app.post('/submit', async (req, res) => {
  const { UPI, Mobile, Timestamp } = req.body;

  try {
    // 1. Read existing Excel file
    const workbook = XLSX.readFile(FILE_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 2. Convert existing sheet to JSON
    const data = XLSX.utils.sheet_to_json(worksheet);

    // 3. Append new entry
    data.push({ UPI, Mobile, Timestamp });

    // 4. Convert JSON back to sheet
    const updatedSheet = XLSX.utils.json_to_sheet(data);

    // 5. Update workbook and write back to file
    workbook.Sheets[sheetName] = updatedSheet;
    XLSX.writeFile(workbook, FILE_PATH);

    res.status(200).json({ status: 'success', message: 'Data appended to Excel' });
  } catch (err) {
    console.error('❌ Excel Error:', err.message);
    res.status(500).json({ status: 'error', message: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`✅ Server running on port ${PORT}`);
});
