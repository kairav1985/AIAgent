const express = require("express");
const XLSX = require("xlsx");
const cors = require("cors");
const fs = require("fs");
const path = require("path");

const app = express();

app.use(cors());
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true }));

// ===============================
// Downloads Folder
// ===============================
const DOWNLOAD_DIR = path.join(__dirname, "downloads");

if (!fs.existsSync(DOWNLOAD_DIR)) {
    fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
}

// Serve generated files
app.use("/downloads", express.static(DOWNLOAD_DIR));

// ===============================
// Health Check
// ===============================
app.get("/", (req, res) => {
    res.send("✅ Excel Service Running");
});

// ===============================
// Generate Excel API
// ===============================
app.post("/api/excel/generate", (req, res) => {

    try {

        console.log("======================================");
        console.log("Incoming Request");
        console.log(JSON.stringify(req.body, null, 2));
        console.log("======================================");

        let {
            issue_key,
            data,
            testCases
        } = req.body;

        issue_key = issue_key || "TestCases";

        // -----------------------------
        // Support AI output
        // -----------------------------
        if (!data && Array.isArray(testCases)) {
            data = testCases;
        }

        // Support stringified JSON
        if (typeof data === "string") {
            data = JSON.parse(data);
        }

        if (!Array.isArray(data) || data.length === 0) {
            return res.status(400).json({
                success: false,
                message: "No data received"
            });
        }

        console.log("Rows :", data.length);

        // Create worksheet
        const worksheet = XLSX.utils.json_to_sheet(data);

        // Create workbook
        const workbook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(
            workbook,
            worksheet,
            "Test Cases"
        );

        // File Name
        const fileName = `${issue_key}_TestCases.xlsx`;

        const filePath = path.join(
            DOWNLOAD_DIR,
            fileName
        );

        XLSX.writeFile(workbook, filePath);

        console.log("Excel Saved :", filePath);

        const downloadUrl =
            `https://${req.get("host")}/downloads/${fileName}`;

        return res.json({
            success: true,
            issue_key,
            totalRows: data.length,
            fileName,
            downloadUrl
        });

    } catch (err) {

        console.error(err);

        return res.status(500).json({
            success: false,
            message: err.message
        });

    }

});

// ===============================
// List Files
// ===============================
app.get("/api/files", (req, res) => {

    try {

        const files = fs.readdirSync(DOWNLOAD_DIR);

        const result = files.map(file => ({
            file,
            url: `https://${req.get("host")}/downloads/${file}`
        }));

        res.json(result);

    } catch (err) {

        res.status(500).json({
            success: false,
            message: err.message
        });

    }

});

// ===============================
// Delete File
// ===============================
app.delete("/api/files/:file", (req, res) => {

    try {

        const filePath = path.join(
            DOWNLOAD_DIR,
            req.params.file
        );

        if (!fs.existsSync(filePath)) {

            return res.status(404).json({
                success: false,
                message: "File not found"
            });

        }

        fs.unlinkSync(filePath);

        res.json({
            success: true,
            message: "File deleted"
        });

    } catch (err) {

        res.status(500).json({
            success: false,
            message: err.message
        });

    }

});

// ===============================
// Start Server
// ===============================
const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {

    console.log("=================================");
    console.log(`✅ Excel Service running on ${PORT}`);
    console.log("=================================");

});