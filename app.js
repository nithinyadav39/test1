const express = require("express"); 
const cors = require("cors");
const bodyParser = require("body-parser");
const XLSX = require("xlsx");
const multer = require("multer");
const path = require("path");
const Fuse = require("fuse.js");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 8080;

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, "public")));
app.use(express.json());  // Enables JSON body parsing
app.use(express.urlencoded({ extended: true }));

// Ensure 'uploads' directory exists
const uploadDir = path.join(__dirname, "uploads");
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

// Persistent storage for script mappings
const scriptFilePath = "script_mappings.json";
let scriptMappings = fs.existsSync(scriptFilePath) ? JSON.parse(fs.readFileSync(scriptFilePath, "utf-8")) : {};

// In-memory storage for uploaded data
const excelData = {};

// Multer File Upload Setup
const storage = multer.diskStorage({
  destination: "uploads/",
  filename: (req, file, cb) => cb(null, file.originalname), // Keep original filename
});
const upload = multer({ storage });

// ✅ Upload and Process Excel Files
app.post("/upload", upload.single("file"), (req, res) => {
  if (!req.file || !req.body.clientName) {
    return res.status(400).json({ error: "File and client name are required." });
  }

  let fileName = req.file.originalname;
  let filePath = req.file.path;
  let clientName = req.body.clientName.trim();

  // Check if client name already exists
  const existingClients = new Set(Object.values(scriptMappings).map(entry => entry.clientName));
  if (existingClients.has(clientName)) {
    return res.status(400).json({ error: "Client name already exists. Please choose a different name." });
  }

  // Generate script ID and URL
  let scriptId = Date.now().toString();
  let redirectUrl = `/ask/${scriptId}`;

  // Update scriptMappings and persist
  scriptMappings[fileName] = { scriptId, redirectUrl, clientName };
  fs.writeFileSync(scriptFilePath, JSON.stringify(scriptMappings, null, 2));

  // Read and process Excel file
  const workbook = XLSX.readFile(filePath);
  const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);

  if (!sheet.length || !sheet[0].Question || !sheet[0].Answer) {
    return res.status(400).json({ error: "The uploaded file is empty or missing required columns." });
  }

  // Store data for search
  excelData[scriptId] = {
    sheet,
    fuse: new Fuse(sheet, { keys: ["Question"], threshold: 0.4 }),
  };

  // Save details to script_links.txt
  fs.appendFile("script_links.txt", `Client: ${clientName}, Script ID: ${scriptId}, File: ${fileName}, URL: http://65.1.176.171:8080${redirectUrl}\n`, (err) => {
    if (err) console.error("Error saving script link:", err);
  });

  res.json({ scriptId, fileName, clientName, redirectUrl });
});


// ✅ Retrieve Script Links
app.get("/script-links", (req, res) => {
  if (!fs.existsSync("script_links.txt")) return res.json({ scripts: [] });
  
  fs.readFile("script_links.txt", "utf-8", (err, data) => {
    if (err) return res.status(500).json({ error: "Error reading script links." });

    const scripts = data.trim().split("\n").map((line) => {
      const parts = line.match(/Client: (.*?), Script ID: (.*?), File: (.*?), URL: (.*)/);
      return parts ? { client: parts[1], scriptId: parts[2], fileName: parts[3], url: parts[4] } : null;
    }).filter(Boolean);

    res.json({ scripts });
  });
});

// ✅ Process Speech Input
app.post("/process-speech/:id", (req, res) => {
  const scriptId = req.params.id;
  const userQuestion = req.body.question?.toLowerCase();

  if (!excelData[scriptId]) return res.status(404).json({ answer: "No data found for this script." });

  const { fuse } = excelData[scriptId];
  const result = fuse.search(userQuestion);
  
  res.json({ answer: result.length ? result[0].item.Answer : "Sorry, please ask related questions." });
});

// ✅ Retrieve Excel Data
app.get("/get-excel/:id", (req, res) => {
  const scriptId = req.params.id;
  const fileEntry = Object.entries(scriptMappings).find(([_, val]) => val.scriptId === scriptId);
  if (!fileEntry) return res.status(404).json({ error: "File not found." });
  
  try {
    const filePath = path.join(__dirname, "uploads", fileEntry[0]);
    const workbook = XLSX.readFile(filePath);
    const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    res.json({ sheet });
  } catch (error) {
    res.status(500).json({ error: "Error reading Excel file." });
  }
});

// ✅ Update Excel Data and Reload in Memory
app.put("/update-excel", (req, res) => {
  try {
    const { scriptId, sheet } = req.body;
    if (!scriptId || !Array.isArray(sheet)) return res.status(400).json({ error: "Invalid data format." });

    // Find the file associated with the script ID
    const fileEntry = Object.entries(scriptMappings).find(([_, val]) => val.scriptId === scriptId);
    if (!fileEntry) return res.status(404).json({ error: "File not found." });

    const filePath = path.join(__dirname, "uploads", fileEntry[0]);

    // ✅ Save updated data to the Excel file
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(sheet);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Sheet1");
    XLSX.writeFile(newWorkbook, filePath);

    // ✅ Reload data in memory
    const updatedWorkbook = XLSX.readFile(filePath);
    const updatedSheet = XLSX.utils.sheet_to_json(updatedWorkbook.Sheets[updatedWorkbook.SheetNames[0]]);

    // ✅ Update in-memory data for voice assistant
    excelData[scriptId] = {
      sheet: updatedSheet,
      fuse: new Fuse(updatedSheet, { keys: ["Question"], threshold: 0.4 }), // Rebuild Fuse.js search
    };

    res.json({ message: "Excel data updated successfully and reloaded for voice assistant!" });
  } catch (error) {
    console.error("Error updating Excel data:", error);
    res.status(500).json({ error: "Server error." });
  }
});


app.delete("/delete/:scriptId", (req, res) => {
  const scriptId = req.params.scriptId;

  // Find the file associated with the script ID
  const fileEntry = Object.entries(scriptMappings).find(([_, val]) => val.scriptId === scriptId);
  if (!fileEntry) return res.status(404).json({ error: "Script not found." });

  const [fileName, scriptData] = fileEntry;
  const filePath = path.join(__dirname, "uploads", fileName);

  try {
    // Delete the Excel file
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);

    // Remove from scriptMappings
    delete scriptMappings[fileName];
    fs.writeFileSync(scriptFilePath, JSON.stringify(scriptMappings, null, 2));

    // Remove from in-memory storage
    delete excelData[scriptId];

    // Remove the script link from script_links.txt
    let scriptLinks = fs.readFileSync("script_links.txt", "utf-8").split("\n").filter(line => !line.includes(scriptId));
    fs.writeFileSync("script_links.txt", scriptLinks.join("\n"));

    res.json({ message: "Script deleted successfully!" });
  } catch (error) {
    console.error("Error deleting script:", error);
    res.status(500).json({ error: "Server error." });
  }
});


// ✅ Serve the Ask Page
app.get("/ask/:id", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "ask.html"));
});

// ✅ Catch-All Route for index.html
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// Route to serve script.html
app.get("/script", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "scripts.html"));
});

// Start Server
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
