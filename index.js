const express = require("express");
const bodyParser = require("body-parser");
const multer = require("multer");
const path = require("path");
const xlsx = require("xlsx");
const fs = require("fs");
const { app: electronApp } = require("electron");

const app = express();
const PORT = 8000;

const isPackaged = electronApp ? electronApp.isPackaged : false;
const baseDir = isPackaged ? path.join(process.resourcesPath, "app") : __dirname;
const uploadsDir = isPackaged
  ? path.join(electronApp.getPath("userData"), "uploads")
  : path.join(baseDir, "uploads");
const downloadsDir = isPackaged
  ? path.join(electronApp.getPath("userData"), "downloads")
  : path.join(baseDir, "downloads");

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static(path.join(baseDir, "public")));

// Multer Storage Configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    if (!fs.existsSync(uploadsDir)) {
      fs.mkdirSync(uploadsDir, { recursive: true });
    }
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});
const upload = multer({ storage: storage });

// View Engine Setup
app.set("view engine", "ejs");
app.set("views", path.join(baseDir, "views"));

// Routes
app.get("/", (req, res) => {
  res.render("index.ejs");
});

app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    const excelFilePath = path.join(uploadsDir, req.file.filename);

    const workbook = xlsx.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    let transformedData = [["RollNo", "CourseSlot Name", "CourseCode", "CourseName"]];

    for (let i = 5; i < worksheet.length; i++) {
      const row = worksheet[i];

      let rollNumber = row[1]; // Assuming Roll No. is in the second column
      let totalCreditIndex = worksheet[0].indexOf("TOTAL CREDIT");

      for (let col = 3; col < totalCreditIndex; col += 5) {
        if (row[col + 3] && row[col + 3] !== "") { // Check if grade exists
          transformedData.push([
            rollNumber,
            worksheet[3][col],  // Credit is in the 3rd row
            worksheet[1][col],  // Course Code is in the 2nd row
            worksheet[0][col],  // Subject Name is in the 1st row
          ]);
        }
      }
    }

    const transformedWorkbook = xlsx.utils.book_new();
    const transformedSheet = xlsx.utils.aoa_to_sheet(transformedData);
    xlsx.utils.book_append_sheet(transformedWorkbook, transformedSheet, "TransformedData");

    if (!fs.existsSync(downloadsDir)) {
      fs.mkdirSync(downloadsDir, { recursive: true });
    }

    const outputFilePath = path.join(downloadsDir, "transformed_data.xls");
    xlsx.writeFile(transformedWorkbook, outputFilePath);

    res.download(outputFilePath, "transformed_data.xls", (err) => {
      if (err) {
        console.error("Error sending file:", err);
      }
      fs.unlinkSync(excelFilePath); // Clean up uploaded file
      fs.unlinkSync(outputFilePath); // Clean up generated file
    });

  } catch (error) {
    console.error("Error processing the file:", error);
    res.status(500).send("Internal Server Error");
  }
});

// Export Express App
module.exports = app;
