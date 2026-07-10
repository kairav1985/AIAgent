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

// Download files
app.use("/downloads", express.static(DOWNLOAD_DIR));

// Health Check
app.get("/", (req, res) => {
    res.send("Excel Service is Running 🚀");
});

// Excel API
app.post("/api/excel/generate", (req, res) => {
    try {

        console.log("========== REQUEST ==========");
        console.log(JSON.stringify(req.body, null, 2));
        console.log("=============================");

        let data = req.body.data;
        const issueKey = req.body.issue_key || "TestCases";

        // If n8n sends string instead of array
        if (typeof data === "string") {
            data = JSON.parse(data);
        }

        if (!Array.isArray(data) || data.length === 0) {
            return res.status(400).json({
                status: "error",
                message: "No data received"
            });
        }

        console.log("Rows:", data.length);

        // Create workbook
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(workbook, worksheet, "Test Cases");

        // Save file
        const fileName = `${issueKey}_TestCases.xlsx`;
        const filePath = path.join(DOWNLOAD_DIR, fileName);

        XLSX.writeFile(workbook, filePath);

        console.log("Saved:", filePath);

        res.json({
            success: true,
            fileName,
            downloadUrl: `${req.protocol}://${req.get("host")}/downloads/${fileName}`
        });

    } catch (err) {

        console.error(err);

        res.status(500).json({
            success: false,
            message: err.message
        });

    }
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
    console.log(`✅ Excel Service running on ${PORT}`);
});