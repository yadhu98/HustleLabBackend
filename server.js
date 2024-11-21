const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const fs = require("fs");
const XLSX = require("xlsx");
const axios = require("axios"); 

const app = express();
const port = 3001;

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Temporary data store (in-memory)
let dataStore = [];
// const mongodb_url = "mongodb://hustlelabsolutions:<db_password>@<hostname>/?ssl=true&replicaSet=atlas-4j2l4h-shard-0&authSource=admin&retryWrites=true&w=majority&appName=Cluster0"
// Route to post details
// Replace with your Apps Script URL
const googleScriptURL = "https://script.google.com/macros/s/AKfycbw-D6PAw9wGNQUJn2BBGTJkmd_FlOayIEUhCSOfKfdWwENunVPYm8KxICqKQFuZUyTU/exec";

// Route to post details
app.post("/add-details", async (req, res) => {
  const { fname, lname, email, phone, message, heardus } = req.body;

  if (!fname || !lname || !email || !phone || !message || !heardus) {
    return res.status(400).json({ error: "All fields are required" });
  }

  try {
    // Send data to Google Apps Script
    const response = await axios.post(googleScriptURL, {
      fname,
      lname,
      email,
      phone,
      message,
      heardus,
    });

    res.json({ message: "Details added successfully", response: response.data });
  } catch (error) {
    console.error("Error sending data to Google Sheets:", error.message);
    res.status(500).json({ error: "Failed to add details" });
  }
});

// Route to download spreadsheet
app.get("/download-data", (req, res) => {
  if (dataStore.length === 0) {
    return res.status(400).json({ error: "No data available to download" });
  }

  // Create a worksheet and workbook
  const ws = XLSX.utils.json_to_sheet(dataStore);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  // Save the file
  const filePath = "./data.xlsx";
  XLSX.writeFile(wb, filePath);

  // Send the file for download
  res.download(filePath, "data.xlsx", (err) => {
    if (err) {
      console.error("Error downloading file:", err);
      res.status(500).send("Error downloading file");
    }

    // Clean up the file
    fs.unlinkSync(filePath);
  });
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
