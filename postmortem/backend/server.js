// backend/server.js
const express    = require("express");
const XLSX       = require("xlsx");
const bodyParser = require("body-parser");
const cors       = require("cors");
const path       = require("path");

const app = express();
app.use(cors());
app.use(bodyParser.json());

const RESP_FILE = path.join(__dirname, "responses.xlsx");
const SHEET_NAME = "Responses";

// POST /submit receives { "1": "Yes", "2": "Some text", … }
app.post("/submit", (req, res) => {
  let wb, data;
  try {
    wb   = XLSX.readFile(RESP_FILE);
    data = XLSX.utils.sheet_to_json(wb.Sheets[SHEET_NAME]);
  } catch {
    wb   = XLSX.utils.book_new();
    data = [];
  }

  data.push(req.body);  // append the new response row

  // write updated data back to the sheet
  const ws = XLSX.utils.json_to_sheet(data, { header: Object.keys(data[0]) });
  XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME, true);
  XLSX.writeFile(wb, RESP_FILE);

  res.sendStatus(200);
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`API listening on http://localhost:${PORT}`));

//email the downloaded copy


