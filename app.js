const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;

// Set up the storage for uploaded files
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.use(express.static('public'));

// Serve an HTML form for uploading Excel files
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

// Handle Excel file upload and parsing
app.post('/upload', upload.single('sampleFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  const buffer = req.file.buffer;
  const workbook = XLSX.read(buffer, { type: 'buffer' });

  // Assuming the data is in the first sheet
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Convert sheet data to an array of objects
  const jsonData = XLSX.utils.sheet_to_json(sheet);

  // Extract the fields dynamically from the first row
  const fields = Object.keys(jsonData[0]);

  // Create an array of objects with dynamic fields and values
  const dataArray = jsonData.map((row) => {
    const dataObject = {};
    fields.forEach((field) => {
      dataObject[field] = row[field];
    });
    return dataObject;
  });

  res.json(dataArray);
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
