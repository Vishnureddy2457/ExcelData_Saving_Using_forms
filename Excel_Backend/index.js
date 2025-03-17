const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = 5000;

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Path to the Excel file
const filePath = path.join(__dirname, "data.xlsx");

// Function to read or create the Excel file
function getWorkbook() {
  if (fs.existsSync(filePath)) {
    return xlsx.readFile(filePath);
  } else {
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet([]);
    xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    return workbook;
  }
}

// API endpoint to handle form submission
app.post("/submit", (req, res) => {
  const { name, email, age } = req.body;

  // Read or create the workbook
  const workbook = getWorkbook();
  const worksheet = workbook.Sheets["Sheet1"];

  // Convert worksheet to JSON
  let data = [];
  if (worksheet) {
    data = xlsx.utils.sheet_to_json(worksheet);
  }

  // Append new data
  data.push({ Name: name, Email: email, Age: age });

  // Update the worksheet
  const newWorksheet = xlsx.utils.json_to_sheet(data);
  workbook.Sheets["Sheet1"] = newWorksheet;

  // Write to the file
  xlsx.writeFile(workbook, filePath);

  res.status(200).send("Data saved successfully!");
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});