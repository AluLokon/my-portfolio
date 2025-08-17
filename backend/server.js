const express = require("express");
const bodyParser = require("body-parser");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public"))); // index.html serve karega

// Excel file ka path
const filePath = path.join(__dirname, "contact_data.xlsx");

// Agar Excel file nahi hai to nayi banao
if (!fs.existsSync(filePath)) {
  const wb = xlsx.utils.book_new();
  const ws = xlsx.utils.json_to_sheet([]);
  xlsx.utils.book_append_sheet(wb, ws, "Contacts");
  xlsx.writeFile(wb, filePath);
}

// POST route
app.post("/save-contact", (req, res) => {
  const { name, phone, problem } = req.body;

  // Purana data padho
  const wb = xlsx.readFile(filePath);
  const ws = wb.Sheets["Contacts"];
  const data = xlsx.utils.sheet_to_json(ws);

  // Naya data add karo
  data.push({ Name: name, Phone: phone, Problem: problem });

  // Wapas Excel me likho
  const newWS = xlsx.utils.json_to_sheet(data);
  wb.Sheets["Contacts"] = newWS;
  xlsx.writeFile(wb, filePath);

  console.log("✅ Data saved:", { name, phone, problem });
  res.send("<h2>Thanks! Your details are saved.</h2><a href='/'>Go Back</a>");
});

// Server start
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
