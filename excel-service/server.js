const express = require("express");
const XLSX = require("xlsx");
const cors = require("cors");
const fs = require("fs");
const path = require("path");

const app = express();

app.use(cors());
app.use(express.json({ limit: "50mb" }));

// Downloads folder
const DOWNLOAD_DIR = path.join(__dirname, "downloads");

if (!fs.existsSync(DOWNLOAD_DIR)) {
    fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
}

// Static download route
app.use("/downloads", express.static(DOWNLOAD_DIR));

// Health check
app.get("/", (req, res) => {
    res.send("Excel Service Running");
});

// Generate Excel
app.post("/api/excel/generate", (req, res) => {
    try {

        const data = req.body.data || [];
        const issueKey = req.body.issue_key || "TestCases";

        if (!Array.isArray(data) || data.length === 0) {
            return res.status(400).json({
                status: "error",
                message: "No data received"
            });
        }

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Test Cases");

        const fileName = `${issueKey}_TestCases.xlsx`;
        const filePath = path.join(DOWNLOAD_DIR, fileName);

        XLSX.writeFile(wb, filePath);

        const downloadUrl =
            `${req.protocol}://${req.get("host")}/downloads/${fileName}`;

        res.json({
            status: "success",
            fileName,
            downloadUrl
        });

    } catch (err) {

        console.error(err);

        res.status(500).json({
            status: "error",
            message: err.message
        });

    }
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
    console.log(`Excel Service running on ${PORT}`);
});