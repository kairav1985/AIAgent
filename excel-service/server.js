const express = require("express");
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();

app.use(cors());
app.use(express.json());

// Health Check
app.get("/", (req, res) => {
    res.send("Excel Service is Running 🚀");
});

// Excel Generation API
app.post("/api/excel/generate", (req, res) => {
    try {

        const data = req.body.data || [];
        const issueKey = req.body.issue_key || "TestCases";

        if (!data.length) {
            return res.status(400).json({
                status: "error",
                message: "No data received"
            });
        }

        const ws = XLSX.utils.json_to_sheet(data);

        const wb = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(wb, ws, "Test Cases");

        // Create Excel in memory
        const buffer = XLSX.write(wb, {
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

        res.send(buffer);

    } catch (err) {

        res.status(500).json({
            status: "error",
            message: err.message
        });

    }
});

const PORT = process.env.PORT || 5002;

app.listen(PORT, () => {
    console.log(`✅ Excel Service running on port ${PORT}`);
});
