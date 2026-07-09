const express = require("express");
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();

app.use(cors());
app.use(express.json());

app.post("/api/excel/generate", (req, res) => {
  try {
    const data = req.body.data || [];
    const issueKey = req.body.issue_key || "TestCases";

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(workbook, worksheet, "Test Cases");

    const fileName = `${issueKey}_TestCases.xlsx`;

    XLSX.writeFile(workbook, fileName);

    res.json({
      status: "success",
      file: fileName
    });

  } catch (err) {
    res.status(500).json({
      status: "error",
      message: err.message
    });
  }
});

app.listen(5002, () => {
  console.log("✅ Excel Service running on http://localhost:5002");
});