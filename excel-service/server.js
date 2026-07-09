const express = require("express");
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();

app.use(cors());
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true }));

// Health Check
app.get("/", (req, res) => {
    res.send("✅ Excel Service is Running");
});

// Excel Generation API
app.post("/api/excel/generate", (req, res) => {
    try {

        console.log("========== REQUEST RECEIVED ==========");
        console.log(JSON.stringify(req.body, null, 2));

        const issueKey = req.body.issue_key || "TestCases";

        let data = req.body.data;

        // String -> Array
        if (typeof data === "string") {
            data = JSON.parse(data);
        }

        // Object -> Array
        if (data && !Array.isArray(data)) {
            data = [data];
        }

        // Empty check
        if (!Array.isArray(data) || data.length === 0) {
            return res.status(400).json({
                status: "error",
                message: "No test case data received"
            });
        }

        console.log("Rows:", data.length);

        // Create worksheet
        const worksheet = XLSX.utils.json_to_sheet(data);

        // Create workbook
        const workbook = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(
            workbook,
            worksheet,
            "Test Cases"
        );

        // Create Excel Buffer
        const buffer = XLSX.write(workbook, {
            type: "buffer",
            bookType: "xlsx"
        });

        res.setHeader(
            "Content-Disposition",
            `attachment; filename=${issueKey}_TestCases.xlsx`
        );

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );

        return res.send(buffer);

    } catch (err) {

        console.error("========== ERROR ==========");
        console.error(err);

        return res.status(500).json({
            status: "error",
            message: err.message
        });

    }
});

const PORT = process.env.PORT || 10000;

app.listen(PORT, () => {
    console.log(`✅ Excel Service running on port ${PORT}`);
});
