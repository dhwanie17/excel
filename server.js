const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const { Pool } = require("pg");

const app = express();
const port = 3000;

// Set up the file upload
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, "uploads");
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + "-" + file.originalname);
  },
});
const upload = multer({ storage: storage });

// Set up the PostgreSQL connection pool
const pool = new Pool({
  username: "postgres",
  host: "127.0.0.1",
  database: "excel",
  password: "1234",
  port: 5432,
});

// Set up the homepage
app.get("/", (req, res) => {
  res.sendFile(__dirname + "/index.html");
});

// Handle the file upload
app.post("/upload", upload.single("file"), (req, res) => {
  // Read the Excel file
  const workbook = xlsx.readFile(req.file.path);
  const sheetName = workbook.SheetNames[0];
  const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  console.log(`Found ${sheetData.length} rows in the Excel file`);
  console.log(`Columns in the Excel file: ${Object.keys(sheetData[0])}`);

  // Store the data in the database
  sheetData.forEach(async (row) => {
    console.log(`Processing row: ${JSON.stringify(row)}`);

    const sql = "INSERT INTO employee (name, email, phone) VALUES ($1, $2, $3)";
    const values = [row.name, row.email, row.phone];

    pool.query(sql, values, (err, result) => {
      if (err) {
        console.error(err);
      } else {
        console.log(`Inserted ${result.rowCount} row(s)`);
      }
    });
  });

  res.send("File uploaded and data stored in the database.");
});

// Start the server
app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
